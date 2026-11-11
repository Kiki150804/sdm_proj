"""Build the final GraphRAG deck from the Geometric color block template.

python build_deck.py   (reads ../Graph-Rag/results/*.json and figures)
"""
import json
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

HERE = Path(__file__).resolve().parent
REPO = HERE.parent
RES = REPO / "results"
FIG = RES / "figures"
NAVY = RGBColor(0x2B, 0x3A, 0x8C)
INK = RGBColor(0x22, 0x22, 0x33)
PINK = RGBColor(0xF6, 0xD6, 0xDA)
ROSE = RGBColor(0xE0, 0x7A, 0x8A)
SKY = RGBColor(0xAC, 0xC4, 0xEA)
OLIVE = RGBColor(0xD5, 0xD8, 0x95)
CREAM = RGBColor(0xFF, 0xFB, 0xF5)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)


def load(name):
    return json.loads((RES / name).read_text(encoding="utf-8"))


ev, v1, v2 = load("eval.json"), load("eval_v1_unbalanced.json"), load("eval_v2_summaries_only.json")
RUNS = [("Unbalanced\n(10 questions)", v1), ("Equal budget\n(12 questions)", v2), ("+ community details\n(final, 6 questions)", ev)]
stats = json.loads((FIG / "stats.json").read_text(encoding="utf-8"))
legend = json.loads((FIG / "graph_legend.json").read_text(encoding="utf-8"))
ANALYSIS = json.loads((HERE / "analysis.json").read_text(encoding="utf-8"))


def gk(run):
    return next(iter(run["wins"]))


W, S = ev["wins"][gk(ev)], ev["summary"]
G, B = S[gk(ev)], S["vector_rag"]
NQ = len(ev["questions"])

prs = Presentation(HERE / "template.pptx")
layouts = {l.name: l for l in prs.slide_layouts}
sld_ids = prs.slides._sldIdLst
for sld in list(sld_ids):  # drop the template's sample slides
    prs.part.drop_rel(sld.rId)
    sld_ids.remove(sld)


# -- helpers -----------------------------------------------------------------
def ph(slide, idx):
    return next(p for p in slide.placeholders if p.placeholder_format.idx == idx)


def remove_ph(s, idx):
    p = ph(s, idx)
    p._element.getparent().remove(p._element)


def place(p, left, top, width, height):
    p.left, p.top, p.width, p.height = Inches(left), Inches(top), Inches(width), Inches(height)


def runs_for(para, text, size, color=None, bold_head=True):
    """'**Head** rest' -> bold head run + normal run."""
    parts = text.split("**")
    pieces = [(parts[1], True), (parts[2], False)] if len(parts) == 3 and parts[0] == "" else [(text, False)]
    for chunk, bold in pieces:
        r = para.add_run()
        r.text = chunk
        r.font.size = Pt(size)
        r.font.bold = bold and bold_head
        if color is not None:
            r.font.color.rgb = color


def bullets(shape, items, size=18, color=None):
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    for i, item in enumerate(items):
        level = 0
        if isinstance(item, tuple):
            level, item = item
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.level = level
        para.space_after = Pt(6)
        runs_for(para, item, size - 2 * level, color)


def slide(layout, title):
    s = prs.slides.add_slide(layouts[layout])
    s.shapes.title.text = title
    return s


def picture(s, path, left, top, width=None, height=None):
    return s.shapes.add_picture(str(path), Inches(left), Inches(top),
                                Inches(width) if width else None, Inches(height) if height else None)


def textbox(s, text, left, top, width, height, size=16, bold=False, color=INK, align=None, anchor=None):
    tb = s.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame
    tf.word_wrap = True
    if anchor:
        tf.vertical_anchor = anchor
    for i, line in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        runs_for(p, line, size, color)
        if bold:
            for r in p.runs:
                r.font.bold = True
        if align:
            p.alignment = align
    return tb


def box(s, left, top, width, height, fill, shape=MSO_SHAPE.ROUNDED_RECTANGLE, line=None):
    b = s.shapes.add_shape(shape, Inches(left), Inches(top), Inches(width), Inches(height))
    b.fill.solid()
    b.fill.fore_color.rgb = fill
    if line is None:
        b.line.fill.background()
    else:
        b.line.color.rgb = line
        b.line.width = Pt(1.25)
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        b.adjustments[0] = 0.12
    b.shadow.inherit = False
    return b


def fill_text(shape, lines, color=INK, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE):
    """lines: [(text, size, bold)]"""
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for m in ("margin_left", "margin_right"):
        setattr(tf, m, Inches(0.12))
    for i, (text, size, bold) in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = text
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color


def stat_card(s, left, top, width, height, value, label, fill=NAVY, color=WHITE):
    b = box(s, left, top, width, height, fill)
    fill_text(b, [(value, 30, True), (label, 13, False)], color=color)


def card(s, left, top, width, height, head, body, head_fill=NAVY, head_color=WHITE, body_size=14, head_size=17):
    box(s, left, top, width, height, WHITE, line=head_fill)
    h = box(s, left, top, width, 0.62, head_fill)
    fill_text(h, [(head, head_size, True)], color=head_color)
    body_tb = textbox(s, "", left + 0.1, top + 0.72, width - 0.2, height - 0.8, size=body_size)
    bullets(body_tb, body, size=body_size, color=INK)


def table(s, rows, left, top, width, height, col_widths=None, size=16, first_col_bold=True):
    shape = s.shapes.add_table(len(rows), len(rows[0]), Inches(left), Inches(top), Inches(width), Inches(height))
    t = shape.table
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            cell = t.cell(r, c)
            cell.text = str(val)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            for p in cell.text_frame.paragraphs:
                if c > 0:
                    p.alignment = PP_ALIGN.CENTER
                for run in p.runs:
                    run.font.size = Pt(size)
                    run.font.bold = r == 0 or (c == 0 and first_col_bold)
                    if r == 0:
                        run.font.color.rgb = NAVY
    if col_widths:
        for i, w in enumerate(col_widths):
            t.columns[i].width = Inches(w)
    return t


def pct(w, key="graphrag"):
    return w[key] / (sum(w.values()) or 1)


def arrow(s, left, top, width, height=0.24):
    a = box(s, left, top, width, height, INK, shape=MSO_SHAPE.RIGHT_ARROW)
    return a


# 1 title ---------------------------------------------------------------------
s = slide("Title Slide", "GraphRAG")
textbox(s, "From local to global query-focused summarization", 3.4, 3.5, 6.5, 0.5, size=18, align=PP_ALIGN.CENTER)
textbox(s, "Sarmada Alla (22CS30067)\nKaranam Keerthi Chandrika (22CS30035)\nGanta Devi Abhijna (22CS30024)",
        4.0, 4.2, 5.3, 1.1, size=14, align=PP_ALIGN.CENTER)

# 2 agenda --------------------------------------------------------------------
s = slide("Agenda", "Agenda")
bullets(ph(s, 1), ["Why GraphRAG", "Our pipeline", "Dataset and knowledge graph",
                   "Evaluation setup", "Results and what improved them", "Limitations and next steps"], size=22)

# 3 why -----------------------------------------------------------------------
s = slide("Comparison 4", "Why GraphRAG?")
remove_ph(s, 2)
remove_ph(s, 4)
card(s, 1.0, 2.35, 4.05, 3.55, "Vector RAG",
     ["Embeds the question, retrieves the top-k most similar chunks",
      "Great for specific, local questions",
      "Misses facts scattered across many documents"], head_fill=ROSE, body_size=19)
card(s, 5.35, 2.35, 4.05, 3.55, "GraphRAG",
     ["LLM builds a knowledge graph of entities and relations",
      "Groups it into hierarchical communities with summaries",
      "Map-reduce over all communities answers corpus-wide questions"], body_size=19)
b = box(s, 1.0, 6.1, 8.4, 0.8, PINK)
fill_text(b, [("Question we test: does GraphRAG answer global “sensemaking” questions better?", 16, True)])

# 4 pipeline ------------------------------------------------------------------
s = slide("Title Only", "Our pipeline")
steps = [("Documents", "32 Wikipedia\narticles"), ("Chunks", "61 × ~1k\ntokens"),
         ("Extraction", "Gemini: entities\n+ relations"), ("Knowledge graph", "712 nodes\n885 edges"),
         ("Communities", "Louvain\n3 levels"), ("Summaries", "bottom-up\nGemini reports")]
x, y, w, h, gap = 0.7, 2.75, 1.8, 1.45, 0.24
textbox(s, "**Indexing (once)**", x, 2.25, 4, 0.45, size=17, color=NAVY)
for i, (head, sub) in enumerate(steps):
    b = box(s, x + i * (w + gap), y, w, h, NAVY)
    fill_text(b, [(head, 16, True)] + [(line, 12, False) for line in sub.split("\n")], color=WHITE)
    if i < len(steps) - 1:
        arrow(s, x + i * (w + gap) + w + 0.02, y + h / 2 - 0.12, gap - 0.04)
textbox(s, "**Query time**", x, 4.45, 4, 0.45, size=17, color=NAVY)
q = [("Global question", []), ("Map", ["each community:", "partial answer + score"]),
     ("Reduce", ["merge best partials", "(≤ 27k chars)"]), ("Global answer", [])]
for i, (head, sub) in enumerate(q):
    bx = x + i * 3.05
    b = box(s, bx, 4.95, 2.55, 1.2, PINK)
    fill_text(b, [(head, 16, True)] + [(line, 12, False) for line in sub])
    if i < len(q) - 1:
        arrow(s, bx + 2.58, 5.43, 0.44)
b = box(s, x, 6.4, 11.9, 0.6, SKY)
fill_text(b, [("Baseline (vector RAG): question → Gemini embedding → top-5 chunks (~27k chars) → Gemini answer", 14, True)])

# 5 dataset -------------------------------------------------------------------
s = slide("Summary 3", "Dataset and setup")
for i, (v, label, fill, color) in enumerate([
        ("32", "Wikipedia articles", NAVY, WHITE), ("65k", "tokens (~249k chars)", ROSE, WHITE),
        ("61", "chunks, shared by both systems", SKY, INK), ("12", "global questions", OLIVE, INK)]):
    stat_card(s, 1.7 + i * 2.75, 2.4, 2.5, 1.3, v, label, fill, color)
place(ph(s, 2), 1.7, 3.95, 6.4, 3.0)
place(ph(s, 15), 8.45, 3.95, 4.1, 3.0)
bullets(ph(s, 2), ["**Topic: **space exploration: Space Race, Apollo, ISS, ESA, ISRO, CNSA, SpaceX, telescopes, rovers",
                   "**Questions: **corpus-wide, e.g. “How has international cooperation and competition shaped space exploration?”",
                   "**Why this size: **fits a day of Gemini free-tier quota"], size=16)
bullets(ph(s, 15), ["**LLM: **Gemini (free tier)", (1, "index: 3.5 Flash-Lite"), (1, "answers + judge: 3.1 Flash-Lite"),
                    "**Embeddings: **gemini-embedding-001", "**Graph: **NetworkX + Louvain"], size=16)

# 6 index ---------------------------------------------------------------------
s = slide("Timeline 2", "What the index contains")
remove_ph(s, 4)
lv = stats["levels"]
total = sum(v["communities"] for v in lv.values())
for i, (v, label, fill, color) in enumerate([
        (str(stats["entities"]), "entities", NAVY, WHITE), (str(stats["relations"]), "relations", ROSE, WHITE),
        (str(total), "communities", SKY, INK), ("≈200", "Gemini calls to build", OLIVE, INK)]):
    stat_card(s, 1.0 + i * 2.9, 2.35, 2.65, 1.3, v, label, fill, color)
rows = [["Level", "Communities", "With 3+ entities (LLM report)", "Largest", "Mean size"]]
for k, v in lv.items():
    rows.append([f"Level {k}" + (" (finest)" if k == "0" else " (coarsest)" if k == str(len(lv) - 1) else ""),
                 v["communities"], v["with_3plus"], v["largest"], f"{v['mean_size']:.1f}"])
table(s, rows, 1.0, 3.95, 11.3, 2.6, col_widths=[2.6, 2.0, 3.5, 1.5, 1.7], size=17)
textbox(s, "Entity types: " + ", ".join(f"{k.lower()} {v}" for k, v in list(stats["entity_types"].items())[:6]),
        1.0, 6.65, 11.3, 0.45, size=14)

# 7 graph ---------------------------------------------------------------------
s = slide("Title Only", "Knowledge graph")
picture(s, FIG / "graph_only.png", 0.5, 2.1, height=5.3)
textbox(s, "**8 largest communities = the corpus's main themes**", 7.3, 2.15, 5.6, 0.45, size=15, color=NAVY)
for i, item in enumerate(legend):
    top = 2.7 + i * 0.56
    dot = box(s, 7.35, top + 0.1, 0.3, 0.3, RGBColor.from_string(item["color"][1:]), shape=MSO_SHAPE.OVAL)
    textbox(s, f"{item['title']}  ({item['size']})", 7.75, top, 5.3, 0.5, size=13)

# 8 hierarchy -----------------------------------------------------------------
s = slide("Content with Caption", "Community hierarchy")
bullets(ph(s, 2), ["Louvain nests communities: level 0 finest, level 2 coarsest.",
                   "Summaries are bottom-up: leaves from entities and relations, parents from child reports.",
                   "We answer from level 2: 21 rich communities cover the whole corpus."], size=16)
remove_ph(s, 1)
picture(s, FIG / "levels.png", 5.6, 1.5, width=7.0)

# 9 free tier -----------------------------------------------------------------
s = slide("Quote", "Engineering for the free tier")
remove_ph(s, 2)
items = [("Rate limiting", "Requests spaced to the per-minute limit; 429s retried with Gemini's own delay", NAVY, WHITE),
         ("Disk cache", "Every reply cached: a run that hits the daily quota resumes at no cost", ROSE, WHITE),
         ("Budget", "1 extraction call per ~1k-token chunk; extractive summaries for tiny communities", SKY, INK),
         ("Robust JSON", "Malformed model replies repaired, then retried once", OLIVE, INK),
         ("Embeddings", "Token-bounded batches under the 30k tokens/minute limit", PINK, INK)]
for i, (head, body, fill, color) in enumerate(items):
    col, row = i % 3, i // 3
    left = 3.78 + col * 2.95 + (1.47 if row else 0)
    b = box(s, left, 2.45 + row * 2.2, 2.75, 2.0, fill)
    fill_text(b, [(head, 18, True), (body, 13, False)], color=color)

# 10 evaluation setup ---------------------------------------------------------
s = slide("Timeline 3", "Evaluation setup")
place(ph(s, 13), 1.7, 2.5, 3.4, 4.2)
bullets(ph(s, 13), ["Both systems answer the same global questions (final run: 6 of 12 judged before the daily quota ran out).",
                    "An LLM judge compares each pair, in both orders, to cancel position bias.",
                    "Same 27k-character context cap for both."], size=15)
remove_ph(s, 1)
table(s, [["Criterion", "The judge asks"],
          ["Comprehensiveness", "How much detail covers all aspects?"],
          ["Diversity", "How varied are the perspectives?"],
          ["Empowerment", "Does it help the reader judge the topic?"],
          ["Directness", "How clearly does it answer?"],
          ["Cheap metrics", "entities mentioned, communities touched, context, time"]],
      5.4, 2.5, 7.1, 4.2, col_widths=[2.5, 4.6], size=16)

# 11 results ------------------------------------------------------------------
s = slide("Title Only", "Results: LLM judge")
picture(s, FIG / f"winrate_{gk(ev)}.png", 0.5, 2.4, width=8.8)
textbox(s, "**GraphRAG win rate**", 9.6, 2.2, 3.4, 0.45, size=15, color=NAVY)
fills = [(NAVY, WHITE), (ROSE, WHITE), (SKY, INK), (OLIVE, INK)]
for i, c in enumerate(W):
    b = box(s, 9.6, 2.7 + i * 1.08, 3.2, 0.95, fills[i][0])
    fill_text(b, [(f"{pct(W[c]):.0%}", 24, True), (c.capitalize(), 12, False)], color=fills[i][1])

# 12 metrics ------------------------------------------------------------------
s = slide("Title Only", "Coverage, context and cost")
picture(s, FIG / "metrics.png", 0.6, 2.2, width=12.1)
b = box(s, 0.9, 6.0, 11.5, 0.9, PINK)
fill_text(b, [(f"GraphRAG mentions {G['coverage'] / max(B['coverage'], 1e-9):.1f}× as many graph entities and touches "
               f"{G['diversity']:.0f} vs {B['diversity']:.0f} communities, at the cost of 22 LLM calls per question vs 1.", 15, False)])

# 13 progression --------------------------------------------------------------
s = slide("Timeline 2", "What made GraphRAG better")
remove_ph(s, 4)
crit = list(W)
rows = [["GraphRAG vs vector RAG"] + [name for name, _ in RUNS]]
for c in crit:
    rows.append([f"{c.capitalize()} win rate"] + [f"{pct(r['wins'][gk(r)][c]):.0%}" for _, r in RUNS])
rows.append(["GraphRAG context (chars)"] + [f"{r['summary'][gk(r)]['context'] / 1000:.1f}k" for _, r in RUNS])
rows.append(["Entities mentioned"] + [f"{r['summary'][gk(r)]['coverage']:.1%}" for _, r in RUNS])
table(s, rows, 1.0, 2.35, 11.3, 4.0, col_widths=[3.5, 2.6, 2.6, 2.6], size=16)
textbox(s, "Fix 1: partial answers no longer capped at 5 sentences; all 21 communities mapped.   "
           "Fix 2: map step also reads each community's extracted entity and relation descriptions.",
        1.0, 6.45, 11.3, 0.7, size=13)

# 14 analysis -----------------------------------------------------------------
s = slide("Summary 3", "Analysis")
bullets(ph(s, 2), ANALYSIS["left"], size=16)
bullets(ph(s, 15), ANALYSIS["right"], size=16)

# 15 limitations --------------------------------------------------------------
s = slide("Summary 3", "Limitations and next steps")
remove_ph(s, 2)
remove_ph(s, 15)
card(s, 1.7, 2.45, 5.3, 4.3, "Limitations",
     ["**Louvain, not Leiden: **Leiden guarantees well-connected communities",
      "**No claims or gleaning: **kept the build within the free tier",
      "**Small corpus: **65k tokens vs ~1M in the paper",
      "**Small sample: **one Lite judge; final run judged 6 questions"], head_fill=ROSE, body_size=18)
card(s, 7.3, 2.45, 5.2, 4.3, "Next steps",
     ["Larger corpus (≥ 500k tokens)", "Compare community levels 0-2", "Leiden + claim extraction",
      "Stronger judge model, more questions", "Hybrid: add top chunks to the reduce step"], body_size=18)

# 16 closing ------------------------------------------------------------------
s = slide("Closing", "Thank\vyou")
bullets(ph(s, 1), ["Questions?", "Code: github.com/Kiki150804/sdm_proj"], size=18)

out = HERE / "GraphRAG_final.pptx"
prs.save(out)
print("saved", out, len(prs.slides), "slides")
