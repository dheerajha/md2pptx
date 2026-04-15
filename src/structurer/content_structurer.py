"""
Content Structurer Module
Transforms a parsed document into an ordered sequence of slide outlines,
ensuring every important section is covered and the slide count target is met.

Flow: Title → Executive Summary → Agenda → Section Content → Data Slides → Conclusion
"""

import math
from typing import Any


class ContentStructurer:
    """
    Distributes document content across a target number of slides with
    intelligent grouping, splitting, and fallback for irregular input.
    """

    def __init__(self, target_slides: int = 12, verbose: bool = False):
        self.target_slides = max(10, min(15, target_slides))
        self.verbose = verbose

    def structure(self, document: dict[str, Any]) -> list[dict]:
        """Convert parsed document into ordered slide outlines."""
        slides = []
        sections = document["sections"]
        tables = document["tables"]
        key_stats = document["key_stats"]

        # ── Reserved slides ──────────────────────────────────────────────────
        # 1 Title, 1 Agenda, 1 Executive Summary, 1 Conclusion = 4 fixed
        reserved = 4
        data_slides_needed = min(len(tables), 2)  # up to 2 data slides
        content_budget = self.target_slides - reserved - data_slides_needed

        # ── 1. Title Slide ───────────────────────────────────────────────────
        slides.append({
            "type": "title",
            "heading": document["title"],
            "subtitle": document["subtitle"],
            "order": 1,
        })

        # ── 2. Executive Summary ─────────────────────────────────────────────
        exec_summary = self._build_exec_summary(sections, key_stats)
        slides.append({
            "type": "executive_summary",
            "heading": "Executive Summary",
            "bullets": exec_summary,
            "order": 2,
        })

        # ── 3. Agenda ────────────────────────────────────────────────────────
        # Get top-level section names (exclude title section)
        h1_sections = [s for s in sections if s["level"] <= 2 and s["heading"] != document["title"]]
        agenda_items = [s["heading"] for s in h1_sections[:8]]
        if agenda_items:
            slides.append({
                "type": "agenda",
                "heading": "Agenda",
                "items": agenda_items,
                "order": 3,
            })

        # ── 4. Content Slides ────────────────────────────────────────────────
        content_sections = self._filter_content_sections(sections, document["title"])
        content_slides = self._distribute_content(content_sections, content_budget)
        for i, cs in enumerate(content_slides):
            cs["order"] = len(slides) + 1 + i
        slides.extend(content_slides)

        # ── 5. Data / Chart Slides ────────────────────────────────────────────
        for table in tables[:data_slides_needed]:
            slides.append({
                "type": "data_chart",
                "heading": self._infer_table_title(table, sections),
                "table": table,
                "order": len(slides) + 1,
            })

        # ── 6. Key Stats Slide (if we have stats and budget allows) ──────────
        remaining_budget = self.target_slides - len(slides) - 1  # -1 for conclusion
        if key_stats and remaining_budget > 0 and not any(s.get("type") == "stats" for s in slides):
            slides.append({
                "type": "stats",
                "heading": "Key Metrics",
                "stats": key_stats[:6],
                "order": len(slides) + 1,
            })

        # ── 7. Conclusion / Key Takeaways ────────────────────────────────────
        takeaways = self._build_takeaways(sections, key_stats)
        slides.append({
            "type": "conclusion",
            "heading": "Key Takeaways",
            "bullets": takeaways,
            "order": len(slides) + 1,
        })

        # ── Pad or trim to target ────────────────────────────────────────────
        slides = self._adjust_count(slides, document)

        if self.verbose:
            for s in slides:
                print(f"     [{s['order']:2d}] {s['type']:20s} {s['heading'][:40]}")

        return slides

    def _filter_content_sections(self, sections: list, doc_title: str) -> list:
        """Get meaningful content sections, excluding title/intro overhead."""
        skippable = {doc_title.lower(), "introduction", "overview", "abstract", "summary"}
        result = []
        for s in sections:
            if s["heading"].lower() in skippable:
                continue
            if not s["content"].strip() and not s["bullet_points"]:
                continue
            result.append(s)
        return result

    def _distribute_content(self, sections: list, budget: int) -> list[dict]:
        """Distribute content sections into slide-sized chunks."""
        if not sections:
            return []

        slides = []
        # Simple strategy: group small sections, split large ones
        groups = self._group_sections(sections, budget)

        for group in groups:
            if len(group) == 1:
                s = group[0]
                slide = self._section_to_slide(s)
            else:
                # Merged slide
                merged_heading = group[0]["heading"]
                merged_bullets = []
                for s in group:
                    if s["bullet_points"]:
                        merged_bullets.extend(s["bullet_points"][:3])
                    elif s["paragraphs"]:
                        merged_bullets.append(s["paragraphs"][0][:120])
                slide = {
                    "type": "bullets",
                    "heading": merged_heading,
                    "bullets": merged_bullets[:6],
                    "subheadings": [s["heading"] for s in group[1:]],
                }
            slides.append(slide)

        return slides[:budget]

    def _group_sections(self, sections: list, budget: int) -> list[list]:
        """Group sections into budget-sized chunks."""
        if len(sections) <= budget:
            return [[s] for s in sections]

        # Need to merge some sections
        groups = [[s] for s in sections]
        while len(groups) > budget:
            # Merge two smallest adjacent groups
            min_size = float("inf")
            min_i = 0
            for i in range(len(groups) - 1):
                size = sum(len(s["bullet_points"]) + len(s["paragraphs"]) for s in groups[i] + groups[i + 1])
                if size < min_size:
                    min_size = size
                    min_i = i
            groups[min_i] = groups[min_i] + groups[min_i + 1]
            groups.pop(min_i + 1)

        return groups

    def _section_to_slide(self, section: dict) -> dict:
        """Convert a single section to a slide outline."""
        # Determine best slide type based on content
        has_bullets = len(section["bullet_points"]) >= 2
        has_para = len(section["paragraphs"]) >= 1
        has_table_ref = "__TABLE_" in section["content"]

        if has_bullets:
            return {
                "type": "bullets",
                "heading": section["heading"],
                "bullets": section["bullet_points"][:6],
                "paragraphs": section["paragraphs"][:1],
            }
        elif has_para:
            return {
                "type": "content",
                "heading": section["heading"],
                "paragraphs": section["paragraphs"][:3],
                "bullets": [],
            }
        else:
            return {
                "type": "content",
                "heading": section["heading"],
                "paragraphs": section["paragraphs"][:2] if section["paragraphs"] else [],
                "bullets": section["bullet_points"][:4],
            }

    def _build_exec_summary(self, sections: list, key_stats: list) -> list[str]:
        """Build executive summary bullets from first paragraphs of each section."""
        bullets = []
        for s in sections[:6]:
            if s["paragraphs"]:
                p = s["paragraphs"][0]
                # Take first sentence
                sent = p.split(". ")[0]
                if len(sent) > 15:
                    bullets.append(sent[:120])
            elif s["bullet_points"]:
                bullets.append(s["bullet_points"][0][:120])

        # Add a stat if available
        if key_stats and len(bullets) < 5:
            bullets.append(key_stats[0]["context"][:120])

        return bullets[:5] if bullets else ["Comprehensive analysis and insights presented in this deck."]

    def _build_takeaways(self, sections: list, key_stats: list) -> list[str]:
        """Build conclusion takeaways."""
        # Look for conclusion/summary section
        for s in sections:
            if any(kw in s["heading"].lower() for kw in ["conclusion", "takeaway", "summary", "recommend"]):
                if s["bullet_points"]:
                    return s["bullet_points"][:5]
                if s["paragraphs"]:
                    return [p[:120] for p in s["paragraphs"][:4]]

        # Fallback: last bullet from each major section
        takeaways = []
        for s in sections[-4:]:
            if s["bullet_points"]:
                takeaways.append(s["bullet_points"][-1][:120])
            elif s["paragraphs"]:
                takeaways.append(s["paragraphs"][-1][:120])

        return takeaways[:5] if takeaways else ["Key insights and findings presented throughout this presentation."]

    def _infer_table_title(self, table: dict, sections: list) -> str:
        """Try to find a section heading near this table."""
        placeholder = table.get("placeholder", "")
        for s in sections:
            if placeholder in s.get("content", ""):
                return s["heading"]
        # Fallback: use column headers
        if table["headers"]:
            return f"Data: {table['headers'][0]}"
        return "Data Overview"

    def _adjust_count(self, slides: list, document: dict) -> list[dict]:
        """Pad or trim slides to hit target count."""
        current = len(slides)

        if current < self.target_slides:
            # Add infographic/process slides from sections with enough content
            gap = self.target_slides - current
            sections = document["sections"]
            for s in sections:
                if gap <= 0:
                    break
                # Sections with multiple bullets not yet in a slide
                heading_in_slides = any(sl.get("heading") == s["heading"] for sl in slides)
                if not heading_in_slides and len(s["bullet_points"]) >= 3:
                    slides.insert(-1, {  # Insert before conclusion
                        "type": "infographic",
                        "heading": s["heading"],
                        "bullets": s["bullet_points"][:5],
                        "order": 0,
                    })
                    gap -= 1

        elif current > self.target_slides:
            # Remove least important slides (keep first 3, last 1, trim middle)
            keep_start = 3
            keep_end = 1
            middle = slides[keep_start:-keep_end]
            excess = current - self.target_slides
            # Remove from middle, prioritizing shorter content slides
            middle.sort(key=lambda s: len(s.get("bullets", [])) + len(s.get("paragraphs", [])))
            middle = middle[excess:]
            middle.sort(key=lambda s: s.get("order", 99))
            slides = slides[:keep_start] + middle + slides[-keep_end:]

        # Re-number orders
        for i, s in enumerate(slides):
            s["order"] = i + 1

        return slides
