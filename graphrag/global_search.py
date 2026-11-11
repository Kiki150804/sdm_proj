"""Map-reduce global answering over community summaries.

Map: every community summary at the chosen level is turned into a partial
answer with a 0-100 helpfulness score. With Gemini, the LLM writes the
partial answer and rates it, as in the paper; offline, the score comes from
TF-IDF relevance and the partial answer is the best-matching summary
sentences. Partial answers scoring 0 are dropped.

Reduce: the remaining partial answers are sorted by score, packed into a
context budget, and combined into one global answer (by Gemini, or as a
list of themes offline).

Every community with at least `min_size` entities is mapped (smaller ones
only get extractive summaries anyway), and the reduce context budget
(`max_chars`) matches the vector RAG baseline's context (~5 chunks, ~27k
characters), so both systems answer from the same amount of text.
"""
from __future__ import annotations

import re
from dataclasses import dataclass

from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

from .communities import Community, CommunityStore

_SENTENCE_SPLIT = re.compile(r"(?<=[.!?])\s+")

_MAP_PROMPT = """Using ONLY the community report and its supporting details below, answer the
question as well as they allow. Return JSON {{"answer": a detailed partial answer (2-4
paragraphs) that cites the specific names, dates, numbers and facts from the details,
"score": integer 0-100 for how helpful this community is for answering the question
(0 = irrelevant)}}.

Report: {title}
{summary}

Supporting details (entities and relationships in this community):
{details}

Question: {query}
"""


_REDUCE_PROMPT = """You are given partial answers to a question, each written from one part of
a document collection and ordered from most to least helpful. Combine them into one
comprehensive, well-organized answer that covers the distinct themes they raise. Do not add
information that is not in the partial answers.

{context}

Question: {query}
Answer:"""


@dataclass
class PartialAnswer:
    community_id: str
    title: str
    text: str
    score: int


def _relevance(query: str, communities: list[Community]) -> list[float]:
    docs = [f"{c.title}. {c.summary}" for c in communities]
    vectorizer = TfidfVectorizer(stop_words="english").fit(docs + [query])
    return list(cosine_similarity(vectorizer.transform([query]), vectorizer.transform(docs))[0])


def _extractive_partial(query: str, community: Community, relevance: float) -> PartialAnswer:
    sentences = [s for s in _SENTENCE_SPLIT.split(community.summary) if s.strip()]
    query_terms = set(re.findall(r"\w+", query.lower()))
    sentences.sort(key=lambda s: -len(query_terms & set(re.findall(r"\w+", s.lower()))))
    # Global questions ("main themes") often share few words with any summary,
    # so every non-empty community still contributes a small base score
    # weighted by its size (rank), as in "what is this whole corpus about".
    score = int(min(100, round(relevance * 100 + min(community.rank, 10))))
    return PartialAnswer(community.community_id, community.title, " ".join(sentences[:2]), score)


def community_details(community: Community, graph_store, max_chars: int = 6000) -> str:
    """Entity and relationship descriptions behind a community, most connected first.

    The paper's community reports are long and fact-dense; ours are short
    (4-6 sentences), so the map step also gets the extracted facts they were
    written from. No extra LLM calls: these come from the graph."""
    if graph_store is None:
        return "(none)"
    graph = graph_store.graph
    nodes = set(community.nodes)
    degree = {n: 0 for n in nodes}
    relations = []
    for u, v, d in graph.edges(data=True):
        if u in nodes and v in nodes:
            degree[u] += 1
            degree[v] += 1
            relations.append((u, v, d))
    lines = []
    for n in sorted(nodes, key=lambda n: -degree[n]):
        desc = graph.nodes[n].get("description", "")
        if desc:
            lines.append(f"- {graph.nodes[n]['name']}: {desc}")
    relations.sort(key=lambda e: -(degree[e[0]] + degree[e[1]]))
    lines += list(dict.fromkeys(f"- {d['sentence']}" for _, _, d in relations))
    text, used = [], 0
    for line in lines:
        if used + len(line) > max_chars:
            break
        text.append(line)
        used += len(line) + 1
    return "\n".join(text) or "(none)"


def _llm_partial(query: str, community: Community, llm, graph_store=None) -> PartialAnswer | None:
    data = llm.generate_json(_MAP_PROMPT.format(
        title=community.title, summary=community.summary,
        details=community_details(community, graph_store), query=query,
    ))
    if not isinstance(data, dict):
        return None
    try:
        score = max(0, min(100, int(data.get("score", 0))))
    except (TypeError, ValueError):
        score = 0
    return PartialAnswer(community.community_id, community.title, str(data.get("answer") or "").strip(), score)


def map_step(query: str, communities: list[Community], llm=None, graph_store=None) -> list[PartialAnswer]:
    if not communities:
        return []
    relevances = _relevance(query, communities)
    partials = []
    for community, relevance in zip(communities, relevances):
        partial = _llm_partial(query, community, llm, graph_store) if llm is not None else None
        if partial is None:
            partial = _extractive_partial(query, community, relevance)
        if partial.score > 0 and partial.text:
            partials.append(partial)
    return sorted(partials, key=lambda p: p.score, reverse=True)


def reduce_step(query: str, partials: list[PartialAnswer], llm=None, max_chars: int = 27000) -> dict:
    packed, used = [], 0
    for partial in partials:
        if used + len(partial.text) > max_chars:
            break
        packed.append(partial)
        used += len(partial.text)
    if not packed:
        return {"answer": "No community summaries were relevant to this question.", "backend": "extractive", "context": ""}

    context = "\n".join(f"[{p.title}] (score {p.score}): {p.text}" for p in packed)
    if llm is not None:
        answer = llm.generate(_REDUCE_PROMPT.format(context=context, query=query)).strip()
        if answer:
            return {"answer": answer, "backend": "gemini", "context": context}

    lines = [f"The corpus covers {len(packed)} relevant theme(s):"]
    lines += [f"- {p.title}: {p.text}" for p in packed]
    return {"answer": "\n".join(lines), "backend": "extractive", "context": context}


def default_level(store: CommunityStore) -> int:
    """Coarsest level that still has more than one community: fewest summaries
    to map over while still separating distinct themes."""
    multi = [lv for lv in store.levels if len(store.at_level(lv)) > 1]
    return multi[-1] if multi else store.levels[0]


def global_answer(
    query: str, store: CommunityStore, level: int | None = None, llm=None, max_communities: int | None = None,
    min_size: int = 3,
    graph_store=None,
) -> dict:
    if not store.levels:
        return {"answer": "No communities have been built.", "backend": "extractive", "level": None, "partials": []}
    if level is None:
        level = default_level(store)
    communities = sorted(
        (c for c in store.at_level(level) if len(c.nodes) >= min_size), key=lambda c: c.rank, reverse=True
    )[:max_communities]
    partials = map_step(query, communities, llm, graph_store)
    result = reduce_step(query, partials, llm)
    result.update(level=level, partials=partials)
    return result
