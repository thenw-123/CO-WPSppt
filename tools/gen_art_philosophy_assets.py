# Generate thematic PNG assets for specs/art-philosophy-beauty.json (matplotlib, no network).
from __future__ import annotations

import math
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Circle, FancyArrowPatch, FancyBboxPatch

# Match deck theme
BG = "#f5f2ea"
INK = "#1a2f4a"
MUTED = "#4a5568"
SECTION = "#4a6670"
ACCENT = "#8b7355"


def _zh_font():
    plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "Noto Sans CJK SC", "DejaVu Sans"]
    plt.rcParams["axes.unicode_minus"] = False


def gist_flow(out: Path) -> None:
    _zh_font()
    fig, ax = plt.subplots(figsize=(10, 2.4), dpi=130)
    fig.patch.set_facecolor(BG)
    ax.set_facecolor(BG)
    ax.axis("off")
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 2.2)
    steps = ["追问", "古典", "近代", "当代", "对话"]
    colors = [INK, SECTION, "#5c6f7c", ACCENT, INK]
    n = len(steps)
    w = 1.28
    gap = (10 - n * w) / (n + 1)
    x = gap
    for i, (label, c) in enumerate(zip(steps, colors)):
        box = FancyBboxPatch(
            (x, 0.52),
            w,
            1.0,
            boxstyle="round,pad=0.02,rounding_size=0.06",
            facecolor=c,
            edgecolor="#e8e4d8",
            linewidth=1.2,
        )
        ax.add_patch(box)
        ax.text(
            x + w / 2,
            1.02,
            label,
            ha="center",
            va="center",
            fontsize=12,
            color="#f7f4ec",
            fontweight="bold",
        )
        if i < n - 1:
            arr = FancyArrowPatch(
                (x + w + 0.02, 1.02),
                (x + w + gap - 0.02, 1.02),
                arrowstyle="-|>",
                mutation_scale=14,
                color=MUTED,
                linewidth=1.8,
            )
            ax.add_patch(arr)
        x += w + gap
    fig.savefig(out, format="png", bbox_inches="tight", pad_inches=0.08, facecolor=BG)
    plt.close(fig)


def abstract_wash(out: Path, hue_shift: float) -> None:
    """Soft radial wash; hue_shift tweaks tone (0..3)."""
    fig, ax = plt.subplots(figsize=(5, 6), dpi=110)
    fig.patch.set_facecolor(BG)
    ax.set_facecolor(BG)
    ax.axis("off")
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    cx, cy = 0.35 + 0.1 * hue_shift, 0.4
    r = 0.85
    base = [(0.12, 0.35, 0.45), (0.18, 0.42, 0.48), (0.22, 0.38, 0.52), (0.35, 0.28, 0.22)]
    r0, g0, b0 = base[int(hue_shift) % len(base)]
    for k in range(8, 0, -1):
        t = k / 8.0
        a = 0.07 + 0.06 * t
        c = Circle((cx, cy), r * t * 0.95, facecolor=(r0, g0, b0, a), edgecolor="none")
        ax.add_patch(c)
    # faint arc — "classical" hint
    arc_x = [0.55 + 0.25 * math.sin(a) for a in [0.2 * i for i in range(20)]]
    arc_y = [0.15 + 0.2 * math.cos(a * 0.8) for a in [0.2 * i for i in range(20)]]
    ax.plot(arc_x, arc_y, color=ACCENT, alpha=0.25, lw=2)
    fig.savefig(out, format="png", bbox_inches="tight", pad_inches=0.02, facecolor=BG)
    plt.close(fig)


def triple_context(out: Path) -> None:
    _zh_font()
    fig, ax = plt.subplots(figsize=(7, 3.2), dpi=120)
    fig.patch.set_facecolor(BG)
    ax.set_facecolor(BG)
    ax.axis("off")
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 4)
    circles = [
        (3.2, 2.0, 1.35, INK, "文化传统"),
        (5.0, 2.35, 1.35, SECTION, "制度市场"),
        (6.8, 2.0, 1.35, ACCENT, "经典阐释"),
    ]
    for x, y, rad, color, label in circles:
        c = Circle((x, y), rad, facecolor=color, alpha=0.22, edgecolor=color, linewidth=2)
        ax.add_patch(c)
        ax.text(x, y, label, ha="center", va="center", fontsize=10, color=INK, fontweight="bold")
    ax.text(5, 0.45, "美在多重力量的交汇处被塑造", ha="center", va="center", fontsize=11, color=MUTED)
    fig.savefig(out, format="png", bbox_inches="tight", pad_inches=0.1, facecolor=BG)
    plt.close(fig)


def cover_band(out: Path) -> None:
    """Thin ornamental band for title slide (optional)."""
    fig, ax = plt.subplots(figsize=(12, 0.45), dpi=120)
    fig.patch.set_facecolor("#2c3e50")
    ax.set_facecolor("#2c3e50")
    ax.axis("off")
    ax.set_xlim(0, 12)
    ax.set_ylim(0, 1)
    for i in range(24):
        x0 = i * 0.5
        ax.plot([x0, x0 + 0.25], [0.5, 0.85], color="#c5d3e0", alpha=0.35, lw=1)
        ax.plot([x0 + 0.25, x0 + 0.5], [0.85, 0.5], color="#8b7355", alpha=0.4, lw=1)
    fig.savefig(out, format="png", bbox_inches="tight", pad_inches=0, facecolor="#2c3e50")
    plt.close(fig)


def main() -> int:
    root = Path(__file__).resolve().parents[1]
    d = root / "assets" / "art-philosophy"
    d.mkdir(parents=True, exist_ok=True)
    gist_flow(d / "gist-flow.png")
    for i in range(4):
        abstract_wash(d / f"abstract-{i + 1}.png", float(i))
    triple_context(d / "triple-context.png")
    cover_band(d / "cover-band.png")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
