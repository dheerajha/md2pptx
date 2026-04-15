"""
Slide Planner / Formatter Module
Assigns visual design decisions to each slide outline:
- Layout type (two-column, card-grid, big-stat, timeline, etc.)
- Color palette (theme-aware)
- Visual element type (chart, icon, infographic, image placeholder)
- Typography decisions
"""

import random
from typing import Any


# ── Color Themes ──────────────────────────────────────────────────────────────

THEMES = {
    "dark": {
        "primary": "1E2761",       # Deep navy
        "secondary": "7EC8E3",     # Sky blue
        "accent": "F5A623",        # Golden amber
        "bg_dark": "0D1B2A",       # Almost-black navy
        "bg_light": "1E2761",      # Navy for content
        "bg_title": "0D1B2A",      # Dark for title
        "text_dark": "FFFFFF",
        "text_light": "E8EAF6",
        "text_muted": "90A4AE",
        "card_bg": "243447",
        "name": "Midnight Executive",
    },
    "light": {
        "primary": "065A82",       # Deep teal
        "secondary": "1C7293",     # Mid teal
        "accent": "F96167",        # Coral
        "bg_dark": "065A82",
        "bg_light": "F4F6F9",
        "bg_title": "021E2B",
        "text_dark": "FFFFFF",
        "text_light": "1A1A2E",
        "text_muted": "546E7A",
        "card_bg": "E8F4F8",
        "name": "Ocean Professional",
    },
}

# Layout templates
LAYOUT_TYPES = [
    "two_column",      # Text left, visual right
    "card_grid",       # 2x2 or 3-card grid
    "big_stat",        # Large number callout center
    "timeline",        # Horizontal steps
    "full_text",       # Large text + accent line
    "image_left",      # Image/chart left, bullets right
    "bullet_clean",    # Clean bullet list
    "comparison",      # Two-column comparison
]

# Icon set (emoji-based for universal compatibility)
ICONS = {
    "technology": "⚙️",
    "data": "📊",
    "growth": "📈",
    "people": "👥",
    "money": "💰",
    "time": "⏱️",
    "security": "🔒",
    "cloud": "☁️",
    "ai": "🤖",
    "global": "🌍",
    "innovation": "💡",
    "process": "🔄",
    "target": "🎯",
    "success": "✅",
    "warning": "⚠️",
    "speed": "⚡",
    "collaboration": "🤝",
    "analysis": "🔍",
    "default": "▶",
}


class SlidePlanner:
    """
    Enriches slide outlines with visual design decisions.
    """

    def __init__(self, theme: str = "auto", verbose: bool = False):
        self.verbose = verbose
        if theme == "auto":
            self.theme_name = "dark"
        else:
            self.theme_name = theme
        self.theme = THEMES.get(self.theme_name, THEMES["dark"])

    def plan(self, slide_outlines: list[dict], document: dict) -> list[dict]:
        """Add visual design metadata to each slide outline."""
        total = len(slide_outlines)
        designed = []

        for i, slide in enumerate(slide_outlines):
            position = i / max(total - 1, 1)  # 0.0 → 1.0
            designed_slide = self._design_slide(slide, position, document)
            designed.append(designed_slide)

        return designed

    def _design_slide(self, slide: dict, position: float, document: dict) -> dict:
        """Apply visual design to a single slide."""
        s = dict(slide)  # Copy
        s_type = s.get("type", "content")

        # ── Background / dark-light sandwich ─────────────────────────────────
        if s_type in ("title", "conclusion"):
            s["bg_color"] = self.theme["bg_title"]
            s["text_color"] = self.theme["text_dark"]
            s["is_dark_bg"] = True
        elif s_type == "executive_summary":
            s["bg_color"] = self.theme["primary"]
            s["text_color"] = self.theme["text_dark"]
            s["is_dark_bg"] = True
        elif s_type == "agenda":
            s["bg_color"] = self.theme["bg_dark"] if self.theme_name == "dark" else self.theme["primary"]
            s["text_color"] = self.theme["text_dark"]
            s["is_dark_bg"] = True
        elif s_type in ("stats", "data_chart"):
            s["bg_color"] = self.theme["secondary"] if not self.theme_name == "dark" else self.theme["card_bg"]
            s["text_color"] = self.theme["text_dark"] if self.theme_name == "dark" else self.theme["text_light"]
            s["is_dark_bg"] = self.theme_name == "dark"
        else:
            # Content slides: alternate to break monotony
            if int(position * 10) % 2 == 0:
                s["bg_color"] = self.theme["bg_light"]
                s["text_color"] = self.theme["text_light"]
                s["is_dark_bg"] = self.theme_name == "dark"
            else:
                s["bg_color"] = self.theme["card_bg"] if self.theme_name == "dark" else "EEF2F7"
                s["text_color"] = self.theme["text_light"]
                s["is_dark_bg"] = self.theme_name == "dark"

        # ── Layout selection ──────────────────────────────────────────────────
        s["layout"] = self._select_layout(s)

        # ── Accent color ──────────────────────────────────────────────────────
        s["accent_color"] = self.theme["accent"]
        s["primary_color"] = self.theme["primary"]
        s["secondary_color"] = self.theme["secondary"]

        # ── Icon ──────────────────────────────────────────────────────────────
        s["icon"] = self._select_icon(s.get("heading", ""))

        # ── Chart type for data slides ────────────────────────────────────────
        if s_type == "data_chart" and s.get("table"):
            s["chart_type"] = self._select_chart_type(s["table"])

        return s

    def _select_layout(self, slide: dict) -> str:
        """Choose the best layout for a slide based on its content type and data."""
        s_type = slide.get("type", "content")
        bullets = slide.get("bullets", [])
        paragraphs = slide.get("paragraphs", [])
        stats = slide.get("stats", [])

        if s_type == "title":
            return "title_hero"
        elif s_type == "agenda":
            return "agenda_numbered"
        elif s_type == "executive_summary":
            return "exec_summary"
        elif s_type == "stats":
            return "stat_callout" if len(stats) <= 3 else "stat_grid"
        elif s_type == "data_chart":
            return "chart_full"
        elif s_type == "conclusion":
            return "conclusion_dark"
        elif s_type == "infographic":
            return "process_flow" if len(bullets) >= 3 else "card_grid"
        elif len(bullets) >= 4:
            return "two_column"
        elif len(bullets) >= 2:
            return "bullet_clean"
        elif paragraphs:
            return "full_text"
        else:
            return "bullet_clean"

    def _select_icon(self, heading: str) -> str:
        """Pick a thematic icon for the slide."""
        heading_lower = heading.lower()
        keyword_map = {
            ("tech", "system", "platform", "software", "code", "api"): "technology",
            ("data", "analytics", "metric", "stat", "number"): "data",
            ("growth", "increase", "scale", "expand"): "growth",
            ("team", "people", "user", "customer", "stakeholder"): "people",
            ("revenue", "cost", "budget", "financial", "profit", "money"): "money",
            ("time", "timeline", "schedule", "deadline", "phase"): "time",
            ("security", "risk", "compliance", "protect"): "security",
            ("cloud", "infrastructure", "deploy", "server"): "cloud",
            ("ai", "ml", "machine learning", "intelligence", "model"): "ai",
            ("global", "market", "international", "world"): "global",
            ("innovat", "creat", "idea", "new", "future"): "innovation",
            ("process", "workflow", "pipeline", "flow"): "process",
            ("goal", "target", "objective", "kpi"): "target",
            ("result", "outcome", "success", "achiev"): "success",
            ("challeng", "problem", "issue", "concern"): "warning",
            ("fast", "speed", "quick", "agile", "rapid"): "speed",
            ("collaborat", "partner", "together"): "collaboration",
            ("analys", "research", "review", "assess"): "analysis",
        }
        for keywords, icon_key in keyword_map.items():
            if any(kw in heading_lower for kw in keywords):
                return ICONS[icon_key]
        return ICONS["default"]

    def _select_chart_type(self, table: dict) -> str:
        """Determine best chart type for a table."""
        headers = table.get("headers", [])
        rows = table.get("rows", [])
        numeric_cols = table.get("numeric_cols", [])

        if not numeric_cols:
            return "table_only"

        num_rows = len(rows)
        num_numeric = len(numeric_cols)

        # Time series: first col looks like years/dates
        if headers and any(kw in headers[0].lower() for kw in ["year", "month", "quarter", "date", "period"]):
            return "line"

        # Comparison categories: bar chart
        if num_rows <= 8 and num_numeric == 1:
            return "bar"

        # Multiple metrics: grouped bar
        if num_rows <= 6 and num_numeric >= 2:
            return "grouped_bar"

        # Proportions: pie (only if 1 numeric col, <8 rows)
        if num_rows <= 6 and num_numeric == 1:
            return "pie"

        return "bar"  # Default
