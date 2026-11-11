"""Build the final GraphRAG deck from the Geometric color block template.

python build_deck.py   (reads ../Graph-Rag/results/eval.json and figures)
"""
import copy
import json
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu, Inches, Pt

HERE = Path(__file__).resolve().parent
REPO = HERE.parent
RES = REPO / "results"
FIG = RES / "figures"
NAVY = RGBColor(0x2B, 0x3A, 0x8C)
INK = RGBColor(0x22, 0x22, 0x33)
PINK = RGBColor(0xF6, 0xD6, 0xDA)

ev = json.loads((RES / "eval.json").read_text(encoding="utf-8"))
v1 = json.loads((RES / "eval_v1_unbalanced.json").read_text(encoding="utf-8"))
stats = json.loads((FIG / "stats.json").read_text(encoding="utf-8"))
GK = next(k for k in ev["wins"])
W = ev["wins"][GK]
S = ev["summary"]
V1W = v1["wins"][next(iter(v1["wins"]))]
NQ = len(ev["questions"])

prs = Presentation(HERE / "template.pptx")
layouts = {l.name: l for l in prs.slide_layouts}

# drop the template's sample slides
sld_ids = prs.slides._sldIdLst
for sld in list(sld_ids):
    prs.part.drop_rel(sld.rId)
    sld_ids.remove(sld)


def ph(slide, idx):
    for p in slide.placeholders:
        if p.placeholder_format.idx == idx:
            return p
    raise KeyError(idx)


def bullets(shape, items, size=None):
    tf = shape.text_frame
    tf.clear()
    for i, item in enumerate(items):
        level = 0
        if isinstance(item, tuple):
            level, item = item
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.level = level
        if "**" in item:  # **bold** prefix
            head, rest = item.split("**")[1], item.split("**", 2)[2]
            r = para.add_run()
            r.text = head
            r.font.bold = True
            r2 = para.add_run()
            r2.text = rest
            runs = [r, r2]
        else:
            r = para.add_run()
            r.text = item
            runs = [r]
        if size:
            for r in runs:
                r.font.size = Pt(size)


def slide(layout, title):
    s = prs.slides.add_slide(layouts[layout])
    s.shapes.title.text = title
    return s


def picture(s, path, left, top, width=None, height=None):
    return s.shapes.add_picture(str(path), Inches(left), Inches(top),
                                Inches(width) if width else None, Inches(height) if height else None)


def remove_ph(s, idx):
    p = ph(s, idx)
    p._element.getparent().remove(p._element)


def table(s, rows, left, top, width, height, col_widths=None, size=14):
    shape = s.shapes.add_table(len(rows), len(rows[0]), Inches(left), Inches(top), Inches(width), Inches(height))
    t = shape.table
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            cell = t.cell(r, c)
            cell.text = str(val)
            for p in cell.text_frame.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(size)
                    run.font.bold = r == 0
                    if r == 0:
                        run.font.color.rgb = NAVY
    if col_widths:
        for i, w in enumerate(col_widths):
            t.columns[i].width = Inches(w)
    return t


def textbox(s, text, left, top, width, height, size=14, bold=False, color=INK, align=None):
    tb = s.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, line in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        r = p.add_run()
        r.text = line
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color
        if align:
            p.alignment = align
    return tb


def pct(w, key="graphrag"):
    return w[key] / (sum(w.values()) or 1)


# 1 title -------------------------------------------------------------------
s = slide("Title Slide", "GraphRAG")
textbox(s, "From local to global query-focused summarization\nFinal project presentation",
        3.4, 3.55, 6.5, 0.9, size=16, align=PP_ALIGN.CENTER)
textbox(s, "Sarmada Alla (22CS30067)\nKaranam Keerthi Chandrika (22CS30035)\nGanta Devi Abhijna (22CS30024)",
        4.2, 4.45, 4.9, 1.0, size=12, align=PP_ALIGN.CENTER)

# 2 agenda ------------------------------------------------------------------
s = slide("Agenda", "Agenda")
bullets(ph(s, 1), ["Background: RAG and why GraphRAG", "Our implementation", "Dataset and knowledge graph",
                   "Evaluation setup", "Results", "Analysis, limitations and next steps"])

# 3 background --------------------------------------------------------------
s = slide("Comparison 4", "Why GraphRAG?")
bullets(ph(s, 2), ["**Vector RAG**", "Retrieves the top-k chunks most similar to the query",
                   "Good for specific, local questions",
                   "Misses information scattered across many documents"], size=15)
bullets(ph(s, 4), ["**GraphRAG**", "Builds a knowledge graph of entities and relations",
                   "Groups it into hierarchical communities with LLM summaries",
                   "Answers global 'sensemaking' questions by map-reduce over all communities"], size=15)

# 4 pipeline ----------------------------------------------------------------
s = slide("Title Only", "Our pipeline")
steps = [("Documents", "32 Wikipedia articles"), ("Chunks", "61 × ~1k tokens"),
         ("Extraction", "Gemini: entities + relations"), ("Knowledge graph", "NetworkX"),
         ("Communities", "Louvain, 3 levels"), ("Summaries", "bottom-up Gemini reports")]
x, y, w, h, gap = 0.75, 2.55, 1.78, 1.25, 0.24
for i, (head, sub) in enumerate(steps):
    box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x + i * (w + gap)), Inches(y), Inches(w), Inches(h))
    box.fill.solid()
    box.fill.fore_color.rgb = NAVY
    box.line.fill.background()
    tf = box.text_frame
    tf.word_wrap = True
    tf.paragraphs[0].text = head
    tf.paragraphs[0].runs[0].font.size = Pt(15)
    tf.paragraphs[0].runs[0].font.bold = True
    p = tf.add_paragraph()
    p.text = sub
    p.runs[0].font.size = Pt(11)
    if i < len(steps) - 1:
        arrow = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(x + i * (w + gap) + w + 0.02), Inches(y + h / 2 - 0.1),
                                   Inches(gap - 0.04), Inches(0.2))
        arrow.fill.solid()
        arrow.fill.fore_color.rgb = INK
        arrow.line.fill.background()
textbox(s, "Indexing (once)", 0.75, 2.1, 4, 0.4, size=13, bold=True, color=NAVY)
textbox(s, "Query time", 0.75, 4.25, 4, 0.4, size=13, bold=True, color=NAVY)
q = [("Global question", ""), ("Map", "partial answer + 0–100 score per community"),
     ("Reduce", "merge best partials (≤27k chars)"), ("Global answer", "")]
for i, (head, sub) in enumerate(q):
    bx = x + i * 3.0
    box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(bx), Inches(4.7), Inches(2.5), Inches(1.15))
    box.fill.solid()
    box.fill.fore_color.rgb = PINK
    box.line.fill.background()
    tf = box.text_frame
    tf.word_wrap = True
    tf.paragraphs[0].text = head
    r = tf.paragraphs[0].runs[0]
    r.font.size, r.font.bold, r.font.color.rgb = Pt(15), True, INK
    if sub:
        p = tf.add_paragraph()
        p.text = sub
        p.runs[0].font.size, p.runs[0].font.color.rgb = Pt(11), INK
    if i < len(q) - 1:
        arrow = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(bx + 2.55), Inches(5.17), Inches(0.4), Inches(0.2))
        arrow.fill.solid()
        arrow.fill.fore_color.rgb = INK
        arrow.line.fill.background()
textbox(s, "Baseline (vector RAG): question → Gemini embedding → top-5 chunks (~27k chars) → Gemini answer",
        0.75, 6.1, 11.5, 0.4, size=13, color=INK)

# 5 dataset & setup ---------------------------------------------------------
s = slide("Summary 3", "Dataset and setup")
bullets(ph(s, 2), ["**Corpus: **32 English Wikipedia articles on space exploration",
                   (1, "Space Race, Apollo, ISS, ESA, ISRO, CNSA, SpaceX, telescopes, rovers…"),
                   (1, "~249k characters ≈ 65k tokens (sized for the Gemini free tier)"),
                   "**Questions: **12 global sensemaking questions",
                   (1, "e.g. “How has international cooperation and competition shaped space exploration?”"),
                   "**Same corpus and chunks for both systems**"], size=15)
bullets(ph(s, 15), ["**LLM: **Google Gemini (free tier)", (1, "Index: 3.5 Flash-Lite"),
                    (1, "Answers + judge: 3.1 Flash-Lite"), "**Embeddings: **gemini-embedding-001",
                    "**Graph: **NetworkX + Louvain", "**Language: **Python"], size=15)

# 6 index stats -------------------------------------------------------------
s = slide("Timeline 2", "What the index contains")
remove_ph(s, 4)
lv = stats["levels"]
rows = [["Stage", "Result"],
        ["Chunks", "61 chunks (~40 sentences each, 3-sentence overlap)"],
        ["Entities", f"{stats['entities']} (organizations, people, spacecraft, missions, locations…)"],
        ["Relations", f"{stats['relations']} described relationships"],
        ["Communities", f"Level 0: {lv['0']['communities']}   ·   Level 1: {lv['1']['communities']}   ·   Level 2: {lv['2']['communities']}"],
        ["LLM-written reports", f"{lv['0']['with_3plus']} / {lv['1']['with_3plus']} / {lv['2']['with_3plus']} communities with 3+ entities"],
        ["Indexing cost", "≈200 Gemini requests in total (extraction + summaries), cached on disk"]]
table(s, rows, 1.0, 2.5, 11.5, 3.9, col_widths=[2.8, 8.7], size=15)

# 7 graph -------------------------------------------------------------------
s = slide("Title Only", "Knowledge graph")
picture(s, FIG / "graph_communities.png", 2.4, 2.25, height=5.1)

# 8 hierarchy ---------------------------------------------------------------
s = slide("Content with Caption", "Community hierarchy")
bullets(ph(s, 2), ["Louvain partitions the graph into nested communities: level 0 is finest, level 2 coarsest.",
                   "Summaries are written bottom-up: leaves from entities and relations, parents from child reports.",
                   "Answers use level 2: its 21 communities with 3+ entities cover the whole corpus."], size=14)
remove_ph(s, 1)
picture(s, FIG / "levels.png", 5.6, 1.4, width=6.9)

# 9 free tier ---------------------------------------------------------------
s = slide("Quote", "Engineering for the free tier")
bullets(ph(s, 2), ["**Rate limiting: **requests spaced to 10 per minute; 429s retried with Gemini's own delay",
                   "**Disk cache: **every reply cached, so a run that hits the daily quota resumes for free",
                   "**Budget: **large chunks (1 extraction call each), extractive summaries for tiny communities",
                   "**Robust JSON: **malformed replies repaired, then retried once",
                   "**Embeddings: **token-bounded batches under the 30k tokens/minute limit"], size=15)

# 10 evaluation setup -------------------------------------------------------
s = slide("Timeline 3", "Evaluation setup")
bullets(ph(s, 13), ["Both systems answer each global question.", "An LLM judge compares the two answers, in both orders to cancel position bias.",
                    "Same context budget cap for both: ~27k chars.", "12 questions, all judged."], size=14)
remove_ph(s, 1)
table(s, [["Criterion", "Question the judge asks"],
          ["Comprehensiveness", "How much detail covers all aspects?"],
          ["Diversity", "How varied are the perspectives?"],
          ["Empowerment", "Does it help the reader judge the topic?"],
          ["Directness", "How clearly does it answer the question?"],
          ["+ cheap metrics", "entity coverage, communities touched, context size, time"]],
      5.56, 2.55, 6.94, 3.9, col_widths=[2.3, 4.64], size=14)

# 11 results win rate -------------------------------------------------------
s = slide("Title Only", "Results: LLM judge")
picture(s, FIG / f"winrate_{GK}.png", 0.9, 2.15, height=4.6)
textbox(s, "GraphRAG win rate\n" + "\n".join(f"{c.capitalize()}: {pct(W[c]):.0%}" for c in W),
        9.7, 3.0, 3.0, 2.5, size=15, color=INK)

# 12 metrics ---------------------------------------------------------------
s = slide("Title Only", "Coverage, context and cost")
picture(s, FIG / "metrics.png", 0.6, 2.3, width=12.1)
g, b = S[GK], S["vector_rag"]
textbox(s, f"GraphRAG answers mention {g['coverage'] / max(b['coverage'], 1e-9):.1f}× as many graph entities and touch "
           f"{g['diversity']:.0f} vs {b['diversity']:.0f} communities. It is slower: one map call per community (times include free-tier rate-limit waits).",
        0.9, 6.25, 11.5, 0.7, size=14)

# 13 before / after ---------------------------------------------------------
s = slide("Comparison 4", "Why a fair context budget matters")
c = list(V1W)
bullets(ph(s, 2), ["**First run (unbalanced)**",
                   "Partial answers capped at 5 sentences; only 15 communities",
                   f"GraphRAG context ≈ {v1['summary'][next(iter(v1['wins']))]['context'] / 1000:.1f}k chars vs 27k for vector RAG",
                   "GraphRAG win rate: " + ", ".join(f"{k[:5]}. {pct(V1W[k]):.0%}" for k in c)], size=14)
bullets(ph(s, 4), ["**Final run (same 27k budget cap)**",
                   "Detailed partial answers; all 21 meaningful communities",
                   f"GraphRAG context ≈ {g['context'] / 1000:.1f}k chars vs {b['context'] / 1000:.1f}k",
                   "GraphRAG win rate: " + ", ".join(f"{k[:5]}. {pct(W[k]):.0%}" for k in c)], size=14)

# 14 analysis -------------------------------------------------------------
s = slide("Summary 3", "Analysis")
ANALYSIS = json.loads((HERE / "analysis.json").read_text(encoding="utf-8"))
bullets(ph(s, 2), ANALYSIS["left"], size=14)
bullets(ph(s, 15), ANALYSIS["right"], size=14)

# 15 limitations -----------------------------------------------------------
s = slide("Summary 3", "Limitations and next steps")
bullets(ph(s, 2), ["**Louvain instead of Leiden: **ships with NetworkX; Leiden adds well-connected communities",
                   "**No claim extraction or gleaning passes: **kept the build within the free tier",
                   "**Small corpus: **65k tokens vs ~1M in the paper; 5 chunks already cover ~8% of it",
                   "**One LLM judge: **3.1 Flash-Lite, 12 questions; small sample"], size=14)
bullets(ph(s, 15), ["**Next**", "Larger corpus (≥500k tokens)", "Compare community levels 0–2",
                    "Leiden + claim extraction", "Stronger judge model, more questions"], size=14)

# 16 closing --------------------------------------------------------------
s = slide("Closing", "Thank\vyou")
bullets(ph(s, 1), ["Code: github.com/Kiki150804/sdm_proj", "Sarmada Alla · Karanam Keerthi Chandrika · Ganta Devi Abhijna"], size=14)

# slide numbers are layout fields; nothing else to do
out = HERE / "GraphRAG_final.pptx"
prs.save(out)
print("saved", out, len(prs.slides), "slides")
