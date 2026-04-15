"""
Chart Generator Module
Creates professional, non-copyrighted charts from table data.
Supports: bar, grouped bar, line, pie, area charts.
Returns chart as bytes (PNG) for embedding in PPTX.
"""

import io
import math
from typing import Any

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.colors import to_rgba
import numpy as np


# Professional color palettes
CHART_PALETTES = {
    "dark": {
        "colors": ["#7EC8E3", "#F5A623", "#5CDB95", "#F96167", "#B8A9C9", "#FFD700"],
        "bg": "#1E2761",
        "text": "#FFFFFF",
        "grid": "#2D3E6B",
        "spine": "#3D4E7B",
    },
    "light": {
        "colors": ["#065A82", "#F96167", "#028090", "#F5A623", "#6D2E46", "#02C39A"],
        "bg": "#F4F6F9",
        "text": "#1A1A2E",
        "grid": "#E0E0E0",
        "spine": "#CCCCCC",
    },
}


class ChartGenerator:
    """Generates matplotlib charts and returns as PNG bytes."""

    def __init__(self, theme: str = "dark"):
        self.theme_name = theme
        self.palette = CHART_PALETTES.get(theme, CHART_PALETTES["dark"])

    def set_theme(self, theme: str):
        self.theme_name = theme
        self.palette = CHART_PALETTES.get(theme, CHART_PALETTES["dark"])

    def generate(self, table: dict, chart_type: str, width_in: float = 6, height_in: float = 4) -> bytes | None:
        """Generate a chart from table data. Returns PNG bytes."""
        try:
            headers = table.get("headers", [])
            rows = table.get("rows", [])
            numeric_cols = table.get("numeric_cols", [])

            if not rows or not headers:
                return None

            if chart_type == "table_only":
                return self._generate_table_chart(headers, rows, width_in, height_in)

            # Extract labels (first column) and numeric data
            labels = [row[0] if row else "" for row in rows]
            data_cols = {}
            for col_i in numeric_cols:
                if col_i < len(headers):
                    col_name = headers[col_i]
                    vals = []
                    for row in rows:
                        if col_i < len(row):
                            try:
                                v = float(str(row[col_i]).replace(",", "").replace("%", "").replace("$", "").replace("K", "e3").replace("M", "e6").replace("B", "e9"))
                                vals.append(v)
                            except (ValueError, AttributeError):
                                vals.append(0)
                        else:
                            vals.append(0)
                    data_cols[col_name] = vals

            if not data_cols:
                return None

            if chart_type == "bar":
                return self._bar_chart(labels, data_cols, width_in, height_in)
            elif chart_type == "grouped_bar":
                return self._grouped_bar_chart(labels, data_cols, width_in, height_in)
            elif chart_type == "line":
                return self._line_chart(labels, data_cols, width_in, height_in)
            elif chart_type == "pie":
                first_col = list(data_cols.values())[0]
                return self._pie_chart(labels, first_col, list(data_cols.keys())[0], width_in, height_in)
            elif chart_type == "area":
                return self._area_chart(labels, data_cols, width_in, height_in)
            else:
                return self._bar_chart(labels, data_cols, width_in, height_in)

        except Exception as e:
            print(f"   ⚠️  Chart generation failed: {e}")
            return None

    def _setup_fig(self, width_in, height_in):
        """Create figure with theme styling."""
        p = self.palette
        fig, ax = plt.subplots(figsize=(width_in, height_in))
        fig.patch.set_facecolor(p["bg"])
        ax.set_facecolor(p["bg"])
        ax.tick_params(colors=p["text"], labelsize=9)
        ax.xaxis.label.set_color(p["text"])
        ax.yaxis.label.set_color(p["text"])
        for spine in ax.spines.values():
            spine.set_edgecolor(p["spine"])
        ax.grid(True, color=p["grid"], alpha=0.4, linewidth=0.7, linestyle="--")
        ax.set_axisbelow(True)
        return fig, ax

    def _save_fig(self, fig) -> bytes:
        """Save figure to PNG bytes."""
        buf = io.BytesIO()
        fig.savefig(buf, format="png", bbox_inches="tight",
                    facecolor=fig.get_facecolor(), dpi=150)
        plt.close(fig)
        buf.seek(0)
        return buf.read()

    def _bar_chart(self, labels: list, data_cols: dict, w: float, h: float) -> bytes:
        p = self.palette
        col_name = list(data_cols.keys())[0]
        values = data_cols[col_name]
        fig, ax = self._setup_fig(w, h)

        bars = ax.bar(labels, values, color=p["colors"][0], edgecolor=p["bg"],
                      linewidth=0.5, width=0.6, zorder=3)

        # Value labels on bars
        for bar, val in zip(bars, values):
            ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() * 1.02,
                    f"{val:,.0f}" if val >= 100 else f"{val:.1f}",
                    ha="center", va="bottom", fontsize=8, color=p["text"], fontweight="bold")

        ax.set_ylabel(col_name, color=p["text"], fontsize=9)
        ax.tick_params(axis="x", rotation=30 if len(labels) > 5 else 0)
        plt.tight_layout()
        return self._save_fig(fig)

    def _grouped_bar_chart(self, labels: list, data_cols: dict, w: float, h: float) -> bytes:
        p = self.palette
        fig, ax = self._setup_fig(w, h)
        n_groups = len(labels)
        n_cols = len(data_cols)
        width = 0.8 / n_cols
        x = np.arange(n_groups)

        for i, (col_name, values) in enumerate(data_cols.items()):
            offset = (i - n_cols / 2 + 0.5) * width
            bars = ax.bar(x + offset, values, width * 0.9,
                          label=col_name, color=p["colors"][i % len(p["colors"])],
                          edgecolor=p["bg"], linewidth=0.3, zorder=3)

        ax.set_xticks(x)
        ax.set_xticklabels(labels, rotation=30 if n_groups > 5 else 0)
        legend = ax.legend(fontsize=8, labelcolor=p["text"],
                           facecolor=p["bg"], edgecolor=p["spine"])
        plt.tight_layout()
        return self._save_fig(fig)

    def _line_chart(self, labels: list, data_cols: dict, w: float, h: float) -> bytes:
        p = self.palette
        fig, ax = self._setup_fig(w, h)

        for i, (col_name, values) in enumerate(data_cols.items()):
            color = p["colors"][i % len(p["colors"])]
            ax.plot(labels, values, marker="o", linewidth=2, color=color,
                    markersize=5, label=col_name, zorder=3)
            ax.fill_between(labels, values, alpha=0.08, color=color)

        if len(data_cols) > 1:
            ax.legend(fontsize=8, labelcolor=p["text"],
                      facecolor=p["bg"], edgecolor=p["spine"])

        ax.tick_params(axis="x", rotation=30 if len(labels) > 6 else 0)
        plt.tight_layout()
        return self._save_fig(fig)

    def _pie_chart(self, labels: list, values: list, col_name: str, w: float, h: float) -> bytes:
        p = self.palette
        fig, ax = plt.subplots(figsize=(w, h))
        fig.patch.set_facecolor(p["bg"])
        ax.set_facecolor(p["bg"])

        colors = p["colors"][:len(values)]
        wedges, texts, autotexts = ax.pie(
            values, labels=labels, autopct="%1.1f%%",
            colors=colors, startangle=90,
            wedgeprops={"edgecolor": p["bg"], "linewidth": 1.5},
            pctdistance=0.8
        )
        for text in texts:
            text.set_color(p["text"])
            text.set_fontsize(8)
        for autotext in autotexts:
            autotext.set_color(p["bg"])
            autotext.set_fontsize(8)
            autotext.set_fontweight("bold")

        ax.set_title(col_name, color=p["text"], fontsize=10, pad=10)
        plt.tight_layout()
        return self._save_fig(fig)

    def _area_chart(self, labels: list, data_cols: dict, w: float, h: float) -> bytes:
        p = self.palette
        fig, ax = self._setup_fig(w, h)

        for i, (col_name, values) in enumerate(data_cols.items()):
            color = p["colors"][i % len(p["colors"])]
            ax.fill_between(labels, values, alpha=0.3, color=color, label=col_name)
            ax.plot(labels, values, color=color, linewidth=2)

        ax.legend(fontsize=8, labelcolor=p["text"], facecolor=p["bg"], edgecolor=p["spine"])
        ax.tick_params(axis="x", rotation=30 if len(labels) > 6 else 0)
        plt.tight_layout()
        return self._save_fig(fig)

    def _generate_table_chart(self, headers: list, rows: list, w: float, h: float) -> bytes:
        """Render a markdown table as a matplotlib table."""
        p = self.palette
        fig, ax = plt.subplots(figsize=(w, h))
        fig.patch.set_facecolor(p["bg"])
        ax.axis("off")

        # Truncate for display
        display_rows = rows[:10]
        display_headers = headers[:6]
        display_data = [[row[i] if i < len(row) else "" for i in range(len(display_headers))]
                        for row in display_rows]

        tbl = ax.table(
            cellText=display_data,
            colLabels=display_headers,
            cellLoc="center",
            loc="center",
        )
        tbl.auto_set_font_size(False)
        tbl.set_fontsize(9)
        tbl.scale(1, 1.5)

        # Style header
        for j in range(len(display_headers)):
            cell = tbl[(0, j)]
            cell.set_facecolor(p["colors"][0])
            cell.set_text_props(color=p["bg"] if self.theme_name == "dark" else "white", fontweight="bold")

        # Style data rows
        for i in range(1, len(display_rows) + 1):
            row_bg = p["bg"] if i % 2 == 0 else p["grid"]
            for j in range(len(display_headers)):
                cell = tbl[(i, j)]
                cell.set_facecolor(row_bg)
                cell.set_text_props(color=p["text"])
                cell.set_edgecolor(p["spine"])

        plt.tight_layout()
        return self._save_fig(fig)
