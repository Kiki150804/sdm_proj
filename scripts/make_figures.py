#!/usr/bin/env python3
"""Make the report/slide figures from a built index and evaluation results.

Usage:
    python scripts/make_figures.py --index data/index --eval results/eval.json --out results/figures
"""
import argparse
import json
import sys
from collections import Counter
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import networkx as nx  # noqa: E402

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from graphrag.communities import CommunityStore  # noqa: E402
from graphrag.graph_store import GraphStore  # noqa: E402

GRAPH_COLOR = "#2B3A8C"   # GraphRAG
VECTOR_COLOR = "#E8A0A8"  # vector RAG
TIE_COLOR = "#C9CFA0"
PALETTE = ["#2B3A8C", "#E07A8A", "#9FB36B", "#6C8FD9", "#D9A441", "#8C5BA8", "#4FA3A5", "#C2603E"]

plt.rcParams.update({"font.size": 13, "axes.spines.top": False, "axes.spines.right": False})


def save(fig, out: Path, name: str) -> None:
    fig.tight_layout()
    fig.savefig(out / name, dpi=200, transparent=False, facecolor="white", bbox_inches="tight", pad_inches=0.1)
    plt.close(fig)
    print("wrote", out / name)


def win_rates(wins: dict, out: Path, name: str, title: str) -> None:
    criteria = list(wins)
    g = [wins[c]["graphrag"] for c in criteria]
    b = [wins[c]["baseline"] for c in criteria]
    t = [wins[c]["tie"] for c in criteria]
    fig, ax = plt.subplots(figsize=(9, 4.8))
    y = range(len(criteria))
    ax.barh(y, g, color=GRAPH_COLOR, label="GraphRAG wins")
    ax.barh(y, t, left=g, color=TIE_COLOR, label="Tie")
    ax.barh(y, b, left=[a + c for a, c in zip(g, t)], color=VECTOR_COLOR, label="Vector RAG wins")
    for i, (gi, ti, bi) in enumerate(zip(g, t, b)):
        total = gi + ti + bi or 1
        ax.text(gi / 2, i, f"{gi}", ha="center", va="center", color="white", fontweight="bold")
        if bi:
            ax.text(gi + ti + bi / 2, i, f"{bi}", ha="center", va="center", fontweight="bold")
        ax.text(total + 0.15, i, f"{gi / total:.0%}", va="center", color=GRAPH_COLOR, fontweight="bold")
    ax.set_yticks(list(y), [c.capitalize() for c in criteria])
    ax.invert_yaxis()
    ax.set_xlabel("Questions")
    ax.set_title(title)
    ax.legend(loc="lower center", bbox_to_anchor=(0.5, -0.38), ncol=3, frameon=False)
    save(fig, out, name)


def metric_bars(summary: dict, out: Path) -> None:
    names = list(summary)
    labels = ["GraphRAG" if n.startswith("graphrag") else "Vector RAG" for n in names]
    colors = [GRAPH_COLOR if n.startswith("graphrag") else VECTOR_COLOR for n in names]
    panels = [
        ("coverage", "Graph entities mentioned", lambda v: f"{v:.1%}", 100),
        ("diversity", "Communities touched", lambda v: f"{v:.1f}", 1),
        ("context", "Context given to answer (k chars)", lambda v: f"{v:.1f}k", 0.001),
        ("calls", "LLM calls per question", lambda v: f"{v:.0f}", 1),
    ]
    # Wall-clock latency is dominated by free-tier waits and the disk cache, so
    # cost is shown as LLM calls: GraphRAG maps every level-2 community with
    # 3+ entities (21) and reduces once; vector RAG makes one answer call.
    calls = {n: 22 if n.startswith("graphrag") else 1 for n in names}
    fig, axes = plt.subplots(1, 4, figsize=(14, 4.2))
    for ax, (key, title, fmt, scale) in zip(axes, panels):
        values = [(calls[n] if key == "calls" else summary[n][key]) * scale for n in names]
        bars = ax.bar(labels, values, color=colors, width=0.6)
        for bar, v in zip(bars, values):
            shown = v / 100 if key == "coverage" else v
            ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height(), fmt(shown), ha="center", va="bottom")
        ax.set_title(title, fontsize=12)
        ax.set_yticks([])
        ax.spines["left"].set_visible(False)
    save(fig, out, "metrics.png")


def graph_picture(graph_store: GraphStore, store: CommunityStore, out: Path, level: int) -> None:
    g = nx.Graph(graph_store.graph.to_undirected())
    g.remove_edges_from(nx.selfloop_edges(g))
    communities = sorted(store.at_level(level), key=lambda c: len(c.nodes), reverse=True)
    top = communities[:8]
    color_of = {n: PALETTE[i] for i, c in enumerate(top) for n in c.nodes}
    keep = [n for n in g if n in color_of]
    sub = g.subgraph(keep)
    pos = nx.spring_layout(sub, seed=7, k=0.35, iterations=80)
    fig, ax = plt.subplots(figsize=(13, 8))
    nx.draw_networkx_edges(sub, pos, ax=ax, alpha=0.15, width=0.6)
    degree = dict(sub.degree())
    nx.draw_networkx_nodes(
        sub, pos, ax=ax, node_color=[color_of[n] for n in sub],
        node_size=[15 + 18 * degree[n] for n in sub], alpha=0.9, linewidths=0,
    )
    hubs = sorted(sub, key=lambda n: -degree[n])[:18]
    nx.draw_networkx_labels(sub, pos, {n: graph_store.graph.nodes[n]["name"] for n in hubs}, ax=ax, font_size=9)
    handles = [
        plt.Line2D([], [], marker="o", ls="", color=PALETTE[i], label=f"{c.title[:45]} ({len(c.nodes)})")
        for i, c in enumerate(top)
    ]
    ax.legend(handles=handles, loc="center left", bbox_to_anchor=(1, 0.5), frameon=False, fontsize=10,
              title=f"Largest level-{level} communities")
    ax.axis("off")
    save(fig, out, "graph_communities.png")

    # legend-free, bigger-label version for slides (legend drawn as slide shapes)
    fig, ax = plt.subplots(figsize=(10, 8))
    nx.draw_networkx_edges(sub, pos, ax=ax, alpha=0.15, width=0.6)
    nx.draw_networkx_nodes(
        sub, pos, ax=ax, node_color=[color_of[n] for n in sub],
        node_size=[20 + 22 * degree[n] for n in sub], alpha=0.9, linewidths=0,
    )
    hubs = sorted(sub, key=lambda n: -degree[n])[:12]
    labels = {n: graph_store.graph.nodes[n]["name"] for n in hubs}
    nx.draw_networkx_labels(sub, pos, labels, ax=ax, font_size=12, font_weight="bold")
    ax.axis("off")
    save(fig, out, "graph_only.png")
    (out / "graph_legend.json").write_text(
        json.dumps([{"title": c.title, "size": len(c.nodes), "color": PALETTE[i]} for i, c in enumerate(top)]),
        encoding="utf-8",
    )


def level_stats(store: CommunityStore, out: Path) -> dict:
    stats = {}
    for lv in store.levels:
        cs = store.at_level(lv)
        sizes = [len(c.nodes) for c in cs]
        stats[lv] = {
            "communities": len(cs),
            "with_3plus": sum(s >= 3 for s in sizes),
            "largest": max(sizes),
            "mean_size": sum(sizes) / len(sizes),
        }
    fig, ax = plt.subplots(figsize=(8, 4.5))
    lv = list(stats)
    ax.bar([x - 0.2 for x in lv], [stats[x]["communities"] for x in lv], 0.4, color=VECTOR_COLOR, label="All communities")
    ax.bar([x + 0.2 for x in lv], [stats[x]["with_3plus"] for x in lv], 0.4, color=GRAPH_COLOR, label="3+ entities (LLM report)")
    for x in lv:
        ax.text(x - 0.2, stats[x]["communities"], stats[x]["communities"], ha="center", va="bottom")
        ax.text(x + 0.2, stats[x]["with_3plus"], stats[x]["with_3plus"], ha="center", va="bottom")
    ax.set_xticks(lv, [f"Level {x}\n({'finest' if x == 0 else 'coarser'})" for x in lv])
    ax.set_title("Community hierarchy (Louvain)")
    ax.legend(frameon=False)
    save(fig, out, "levels.png")
    return stats


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--index", default="data/index")
    parser.add_argument("--eval", default="results/eval.json")
    parser.add_argument("--out", default="results/figures")
    args = parser.parse_args()
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)

    index = Path(args.index)
    graph_store = GraphStore.load(index / "graph.pkl")
    store = CommunityStore.load(index / "communities.pkl")
    g = graph_store.graph
    types = Counter(d.get("type", "OTHER") for _, d in g.nodes(data=True))
    stats = {
        "entities": g.number_of_nodes(),
        "relations": g.number_of_edges(),
        "entity_types": dict(types.most_common(8)),
        "levels": level_stats(store, out),
        "top_communities": [
            {"id": c.community_id, "title": c.title, "size": len(c.nodes), "summary": c.summary}
            for c in sorted(store.at_level(max(store.levels)), key=lambda c: len(c.nodes), reverse=True)[:6]
        ],
    }
    graph_picture(graph_store, store, out, level=max(store.levels))

    fig, ax = plt.subplots(figsize=(8, 4.5))
    names, counts = zip(*types.most_common(8))
    ax.barh(names, counts, color=GRAPH_COLOR)
    ax.invert_yaxis()
    for i, c in enumerate(counts):
        ax.text(c, i, f" {c}", va="center")
    ax.set_title("Extracted entity types")
    save(fig, out, "entity_types.png")

    eval_path = Path(args.eval)
    if eval_path.exists():
        data = json.loads(eval_path.read_text(encoding="utf-8"))
        for name, wins in data["wins"].items():
            n = sum(wins[next(iter(wins))].values())
            win_rates(wins, out, f"winrate_{name}.png", f"LLM judge: {name.replace('graphrag_', 'GraphRAG ')} vs vector RAG ({n} questions)")
        metric_bars(data["summary"], out)
        stats["eval_questions"] = len(data["questions"])
    (out / "stats.json").write_text(json.dumps(stats, indent=2), encoding="utf-8")
    print(json.dumps({k: v for k, v in stats.items() if k != "top_communities"}, indent=2))


if __name__ == "__main__":
    main()
