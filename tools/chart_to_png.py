# Render a simple chart to PNG for WPS slides (matplotlib, headless).
# Usage: python chart_to_png.py <output.png> <config.json>
# config: { "chart_type": "bar"|"line"|"donut", "chart_data": { "labels": [...], "series": [{ "name"?: str, "values": [numbers] }] } }

from __future__ import annotations

import json
import sys
from pathlib import Path


def main() -> int:
    if len(sys.argv) != 3:
        print("Usage: chart_to_png.py <output.png> <config.json>", file=sys.stderr)
        return 2
    out_path = Path(sys.argv[1])
    cfg_path = Path(sys.argv[2])
    if not cfg_path.is_file():
        print(f"Config not found: {cfg_path}", file=sys.stderr)
        return 2
    raw = json.loads(cfg_path.read_text(encoding="utf-8-sig"))
    chart_type = str(raw.get("chart_type", "bar")).lower().strip()
    data = raw.get("chart_data") or {}
    labels = [str(x) for x in (data.get("labels") or [])]
    series_raw = data.get("series") or []
    if not labels:
        print("chart_data.labels required", file=sys.stderr)
        return 3
    series: list[tuple[str | None, list[float]]] = []
    for i, s in enumerate(series_raw):
        if isinstance(s, dict):
            name = s.get("name")
            vals = s.get("values") or []
        else:
            name = None
            vals = s if isinstance(s, list) else []
        try:
            nums = [float(x) for x in vals]
        except (TypeError, ValueError):
            print(f"series[{i}].values must be numbers", file=sys.stderr)
            return 3
        series.append((str(name) if name else None, nums))

    if not series:
        print("chart_data.series required", file=sys.stderr)
        return 3
    for i, (_, vals) in enumerate(series):
        if len(vals) != len(labels):
            print(
                f"series[{i}] length {len(vals)} != labels length {len(labels)}",
                file=sys.stderr,
            )
            return 3

    import matplotlib

    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    out_path.parent.mkdir(parents=True, exist_ok=True)

    fig, ax = plt.subplots(figsize=(8, 4.5), dpi=120)
    x = range(len(labels))

    if chart_type == "line":
        for name, vals in series:
            label = name or "Series"
            ax.plot(x, vals, marker="o", linewidth=2, label=label)
        ax.set_xticks(list(x))
        ax.set_xticklabels(labels)
        ax.legend()
        ax.grid(True, alpha=0.3)
    elif chart_type == "donut":
        # Single-series donut; if multiple, sum or use first
        vals = series[0][1]
        ax.pie(
            vals,
            labels=labels,
            autopct="%1.0f%%",
            wedgeprops=dict(width=0.5),
            startangle=90,
        )
        ax.axis("equal")
        if series[0][0]:
            ax.set_title(series[0][0])
    else:
        # bar (default)
        n = len(series)
        w = 0.8 / max(n, 1)
        for si, (name, vals) in enumerate(series):
            offset = (si - (n - 1) / 2) * w
            xs = [xi + offset for xi in x]
            label = name or f"S{si + 1}"
            ax.bar(xs, vals, width=w * 0.95, label=label)
        ax.set_xticks(list(x))
        ax.set_xticklabels(labels)
        ax.legend()
        ax.grid(True, axis="y", alpha=0.3)

    fig.tight_layout()
    fig.savefig(out_path, format="png", bbox_inches="tight")
    plt.close(fig)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
