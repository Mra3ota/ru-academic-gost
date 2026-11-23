#!/usr/bin/env python3
"""Generate print-ready GOST-styled academic charts (300 DPI PNG) with Times New Roman and Russian decimal commas."""

from __future__ import annotations

import argparse
import json
import sys
import textwrap
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _common import ensure_utf8_io

PALETTES: dict[str, list[str]] = {
    # Muted high-contrast academic steel/slate palette (readable in both PDF and B&W print)
    "academic": ["#243B53", "#627D98", "#9FB3C8", "#486581", "#BCCCDC", "#334E68"],
    # Strict monochrome grayscale for strict B&W printing
    "mono": ["#222222", "#666666", "#A6A6A6", "#D9D9D9", "#444444", "#888888"],
}

HATCHES = ["", "///", "...", "xxx", "\\\\\\", "+++"]
MARKERS = ["o", "s", "^", "D", "v", "P"]
LINESTYLES = ["-", "--", "-.", ":", (0, (5, 2, 1, 2)), (0, (3, 1, 1, 1))]


def fmt_ru_number(val: float, decimals: int | None = None, suffix: str = "") -> str:
    if decimals is None:
        if float(val).is_integer():
            s = f"{int(round(val)):,}".replace(",", " ")
        else:
            s = f"{val:,.1f}".replace(",", " ").replace(".", ",")
    else:
        s = f"{val:,.{decimals}f}".replace(",", " ").replace(".", ",")
    return f"{s}{suffix}"


def wrap_labels(labels: list[str], width: int = 18) -> list[str]:
    return ["\n".join(textwrap.wrap(str(lb), width=width)) for lb in labels]


def render_chart(spec: dict[str, Any], out_path: Path, dpi: int = 300) -> None:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.ticker import FuncFormatter

    plt.rcParams.update({
        "font.family": "serif",
        "font.serif": ["Times New Roman", "Liberation Serif", "DejaVu Serif"],
        "font.size": 11,
        "axes.labelsize": 11,
        "xtick.labelsize": 10,
        "ytick.labelsize": 10,
        "legend.fontsize": 10,
        "axes.edgecolor": "#333333",
        "axes.linewidth": 0.8,
        "figure.facecolor": "white",
        "axes.facecolor": "white",
    })

    chart_type = str(spec.get("type", "bar")).lower()
    palette_name = str(spec.get("palette", "academic")).lower()
    colors = PALETTES.get(palette_name, PALETTES["academic"])
    use_hatch = bool(spec.get("hatch", palette_name == "mono"))
    width_in = float(spec.get("width_in", 6.3))   # ~16 cm fits A4 GOST margins (16.5 cm)
    height_in = float(spec.get("height_in", 3.8))
    decimals = spec.get("decimals", None)
    val_suffix = str(spec.get("value_suffix", ""))
    y_label = str(spec.get("y_label", ""))
    x_label = str(spec.get("x_label", ""))
    categories = [str(c) for c in spec.get("categories", [])]
    series = spec.get("series", [])
    if not series and "values" in spec:
        series = [{"name": str(spec.get("series_name", "")), "values": spec["values"]}]

    fig, ax = plt.subplots(figsize=(width_in, height_in), dpi=dpi)

    def _ru_tick(x: float, _pos: Any) -> str:
        return fmt_ru_number(x, decimals=0 if float(x).is_integer() else 1)

    if chart_type == "bar":
        vals = [float(v) for v in series[0]["values"]]
        x_pos = list(range(len(categories)))
        bars = ax.bar(
            x_pos,
            vals,
            width=0.58,
            color=colors[0],
            edgecolor="#111111",
            linewidth=0.7,
            hatch=(HATCHES[0] if use_hatch else None),
            zorder=3,
        )
        ax.set_xticks(x_pos)
        ax.set_xticklabels(wrap_labels(categories, width=int(spec.get("label_wrap", 18))))
        ax.yaxis.set_major_formatter(FuncFormatter(_ru_tick))
        ax.grid(axis="y", linestyle="--", linewidth=0.5, color="#CCCCCC", zorder=0)
        ymax = max(vals) if vals else 1.0
        ax.set_ylim(0, ymax * 1.18)
        for bar, v in zip(bars, vals):
            ax.text(
                bar.get_x() + bar.get_width() / 2.0,
                v + ymax * 0.025,
                fmt_ru_number(v, decimals=decimals, suffix=val_suffix),
                ha="center",
                va="bottom",
                fontsize=10,
                color="#111111",
            )

    elif chart_type == "hbar":
        vals = [float(v) for v in series[0]["values"]]
        y_pos = list(range(len(categories)))
        bars = ax.barh(
            y_pos,
            vals,
            height=0.58,
            color=colors[0],
            edgecolor="#111111",
            linewidth=0.7,
            zorder=3,
        )
        ax.set_yticks(y_pos)
        ax.set_yticklabels(wrap_labels(categories, width=int(spec.get("label_wrap", 28))))
        ax.invert_yaxis()
        ax.xaxis.set_major_formatter(FuncFormatter(_ru_tick))
        ax.grid(axis="x", linestyle="--", linewidth=0.5, color="#CCCCCC", zorder=0)
        xmax = max(vals) if vals else 1.0
        ax.set_xlim(0, xmax * 1.22)
        for bar, v in zip(bars, vals):
            ax.text(
                v + xmax * 0.02,
                bar.get_y() + bar.get_height() / 2.0,
                fmt_ru_number(v, decimals=decimals, suffix=val_suffix),
                ha="left",
                va="center",
                fontsize=10,
                color="#111111",
            )

    elif chart_type == "grouped_bar":
        n_series = len(series)
        n_cats = len(categories)
        total_w = 0.74
        bar_w = total_w / max(1, n_series)
        all_vals = [float(v) for s in series for v in s["values"]]
        ymax = max(all_vals) if all_vals else 1.0

        for s_idx, s in enumerate(series):
            vals = [float(v) for v in s["values"]]
            offsets = [i - total_w / 2.0 + (s_idx + 0.5) * bar_w for i in range(n_cats)]
            bars = ax.bar(
                offsets,
                vals,
                width=bar_w * 0.92,
                label=str(s.get("name", f"Серия {s_idx + 1}")),
                color=colors[s_idx % len(colors)],
                edgecolor="#111111",
                linewidth=0.65,
                hatch=(HATCHES[s_idx % len(HATCHES)] if use_hatch else None),
                zorder=3,
            )
            for bar, v in zip(bars, vals):
                ax.text(
                    bar.get_x() + bar.get_width() / 2.0,
                    v + ymax * 0.022,
                    fmt_ru_number(v, decimals=decimals, suffix=val_suffix),
                    ha="center",
                    va="bottom",
                    fontsize=9,
                    color="#111111",
                )

        ax.set_xticks(list(range(n_cats)))
        ax.set_xticklabels(wrap_labels(categories, width=int(spec.get("label_wrap", 18))))
        ax.yaxis.set_major_formatter(FuncFormatter(_ru_tick))
        ax.grid(axis="y", linestyle="--", linewidth=0.5, color="#CCCCCC", zorder=0)
        ax.set_ylim(0, ymax * 1.22)
        ax.legend(frameon=False, loc="upper center", bbox_to_anchor=(0.5, 1.13), ncol=min(3, n_series))

    elif chart_type == "line":
        x_pos = list(range(len(categories)))
        all_vals = [float(v) for s in series for v in s["values"]]
        ymin, ymax = (min(all_vals), max(all_vals)) if all_vals else (0.0, 1.0)
        span = max(ymax - ymin, abs(ymax) * 0.2, 1.0)

        for s_idx, s in enumerate(series):
            vals = [float(v) for v in s["values"]]
            ax.plot(
                x_pos,
                vals,
                label=str(s.get("name", f"Серия {s_idx + 1}")),
                color=colors[s_idx % len(colors)],
                linestyle=LINESTYLES[s_idx % len(LINESTYLES)],
                marker=MARKERS[s_idx % len(MARKERS)],
                markersize=5.5,
                linewidth=1.8,
                zorder=4,
            )
            if spec.get("show_values", True):
                for xi, v in zip(x_pos, vals):
                    ax.text(
                        xi,
                        v + span * 0.045,
                        fmt_ru_number(v, decimals=decimals, suffix=val_suffix),
                        ha="center",
                        va="bottom",
                        fontsize=9,
                        color="#111111",
                    )

        ax.set_xticks(x_pos)
        ax.set_xticklabels(wrap_labels(categories, width=int(spec.get("label_wrap", 14))))
        ax.yaxis.set_major_formatter(FuncFormatter(_ru_tick))
        ax.grid(axis="y", linestyle="--", linewidth=0.5, color="#CCCCCC", zorder=0)
        ax.set_ylim(max(0.0, ymin - span * 0.15) if ymin >= 0 else ymin - span * 0.15, ymax + span * 0.22)
        if len(series) > 1 or series[0].get("name"):
            ax.legend(frameon=False, loc="upper center", bbox_to_anchor=(0.5, 1.13), ncol=min(3, len(series)))

    elif chart_type == "pie":
        vals = [float(v) for v in series[0]["values"]]
        total = sum(vals) or 1.0
        wedges, _, autotexts = ax.pie(
            vals,
            labels=None,
            colors=[colors[i % len(colors)] for i in range(len(vals))],
            autopct=lambda pct: fmt_ru_number(pct, decimals=1, suffix="%") if pct >= 4.0 else "",
            startangle=90,
            counterclock=False,
            wedgeprops={"edgecolor": "#FFFFFF", "linewidth": 1.2},
            pctdistance=0.72,
        )
        for at in autotexts:
            at.set_fontsize(9.5)
            at.set_color("#FFFFFF")
            at.set_fontweight("bold")
        legend_labels = [
            f"{cat} — {fmt_ru_number(v / total * 100.0, decimals=1, suffix='%')}"
            for cat, v in zip(categories, vals)
        ]
        ax.legend(wedges, legend_labels, loc="center left", bbox_to_anchor=(1.0, 0.5), frameon=False)
        ax.set_aspect("equal")
    else:
        raise ValueError(f"Unsupported chart type: {chart_type}")

    if chart_type != "pie":
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        if y_label:
            ax.set_ylabel(y_label)
        if x_label:
            ax.set_xlabel(x_label)

    # GOST puts caption below the figure in Word, so internal title is omitted unless explicitly forced
    if spec.get("internal_title"):
        ax.set_title(str(spec["internal_title"]), fontsize=12, fontweight="bold", pad=14)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    fig.tight_layout()
    fig.savefig(str(out_path), dpi=dpi, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def main() -> int:
    ensure_utf8_io()
    parser = argparse.ArgumentParser(
        description="Render a GOST-styled 300 DPI PNG chart from a JSON spec."
    )
    parser.add_argument("--spec", required=True, help="Path to JSON chart specification file")
    parser.add_argument("--out", required=True, help="Output PNG path (e.g., data/fig_1.png)")
    parser.add_argument("--dpi", type=int, default=300, help="Output resolution in DPI (default: 300)")
    args = parser.parse_args()

    spec_path = Path(args.spec)
    if not spec_path.exists():
        print(f"ERROR: spec file not found: {spec_path}", file=sys.stderr)
        return 2

    spec = json.loads(spec_path.read_text(encoding="utf-8"))
    out_path = Path(args.out)
    render_chart(spec, out_path, dpi=args.dpi)

    fig_num = spec.get("figure_number", 1)
    caption = spec.get("caption", "Название рисунка")
    source = spec.get("source", "")
    rel_img = args.out.replace("\\", "/")

    print(f"CHART OK: {out_path} ({args.dpi} DPI)")
    print("\nMarkdown snippet for draft:")
    print(f"![]({rel_img}){{width=15cm}}\n")
    print(f"Рисунок {fig_num} — {caption}")
    if source:
        print(f"\nИсточник: {source}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
