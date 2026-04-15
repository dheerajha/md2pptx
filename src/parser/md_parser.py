"""
Markdown Parser Module
Converts raw markdown text into a structured document representation.
Handles: headings, paragraphs, lists, tables, code blocks, emphasis, links.
"""

import re
import json
from typing import Any


class MarkdownParser:
    """
    Parses Markdown into a structured document dictionary.
    Robust to irregular formatting, missing sections, and large files.
    """

    def __init__(self, verbose: bool = False):
        self.verbose = verbose

    def parse(self, text: str) -> dict[str, Any]:
        """Parse markdown text into structured document."""
        # Normalize line endings
        text = text.replace("\r\n", "\n").replace("\r", "\n")

        doc = {
            "title": "",
            "subtitle": "",
            "sections": [],
            "tables": [],
            "code_blocks": [],
            "key_stats": [],
            "metadata": {},
            "stats": {},
        }

        # Extract frontmatter if present
        text, doc["metadata"] = self._extract_frontmatter(text)

        # Extract code blocks (protect from further parsing)
        text, code_blocks = self._extract_code_blocks(text)
        doc["code_blocks"] = code_blocks

        # Extract tables
        text, tables = self._extract_tables(text)
        doc["tables"] = tables

        # Parse headings and sections
        sections = self._parse_sections(text)
        doc["sections"] = sections

        # Determine title
        doc["title"] = self._find_title(sections, doc["metadata"])
        doc["subtitle"] = self._find_subtitle(sections)

        # Extract key statistics (numbers with context)
        doc["key_stats"] = self._extract_stats(text)

        # Compute stats
        doc["stats"] = {
            "word_count": len(text.split()),
            "section_count": len(sections),
            "table_count": len(tables),
            "code_block_count": len(code_blocks),
            "stat_count": len(doc["key_stats"]),
        }

        if self.verbose:
            print(f"     Title: {doc['title'][:60]}")
            print(f"     Sections: {[s['heading'] for s in sections[:5]]}")

        return doc

    def _extract_frontmatter(self, text: str) -> tuple[str, dict]:
        """Extract YAML frontmatter if present."""
        metadata = {}
        if text.startswith("---"):
            end = text.find("---", 3)
            if end != -1:
                fm = text[3:end].strip()
                for line in fm.split("\n"):
                    if ":" in line:
                        k, _, v = line.partition(":")
                        metadata[k.strip()] = v.strip()
                text = text[end + 3:].strip()
        return text, metadata

    def _extract_code_blocks(self, text: str) -> tuple[str, list]:
        """Extract and replace code blocks with placeholders."""
        blocks = []
        def replace_block(m):
            lang = m.group(1) or "text"
            code = m.group(2).strip()
            placeholder = f"__CODE_BLOCK_{len(blocks)}__"
            blocks.append({"lang": lang, "code": code, "placeholder": placeholder})
            return placeholder
        text = re.sub(r"```(\w*)\n(.*?)```", replace_block, text, flags=re.DOTALL)
        # Inline code
        text = re.sub(r"`([^`]+)`", r"[\1]", text)
        return text, blocks

    def _extract_tables(self, text: str) -> tuple[str, list]:
        """Extract markdown tables into structured data."""
        tables = []
        lines = text.split("\n")
        result_lines = []
        i = 0
        while i < len(lines):
            line = lines[i]
            # Detect table start: line with pipes
            if "|" in line and i + 1 < len(lines) and re.match(r"[\|\s\-:]+", lines[i + 1]):
                table_lines = [line]
                i += 1
                # Skip separator row
                sep_line = lines[i]
                i += 1
                # Collect rows
                while i < len(lines) and "|" in lines[i]:
                    table_lines.append(lines[i])
                    i += 1
                table = self._parse_table(table_lines)
                if table:
                    placeholder = f"__TABLE_{len(tables)}__"
                    tables.append({**table, "placeholder": placeholder})
                    result_lines.append(placeholder)
                continue
            result_lines.append(line)
            i += 1
        return "\n".join(result_lines), tables

    def _parse_table(self, lines: list[str]) -> dict | None:
        """Parse table lines into header + rows."""
        if not lines:
            return None
        def split_row(line):
            cells = [c.strip() for c in line.strip("|").split("|")]
            return [c for c in cells if c or True]  # keep empty cells

        headers = split_row(lines[0])
        rows = [split_row(l) for l in lines[1:]]

        # Detect if numeric data (for chart generation)
        numeric_cols = []
        for col_i in range(1, len(headers)):
            vals = []
            for row in rows:
                if col_i < len(row):
                    try:
                        vals.append(float(row[col_i].replace(",", "").replace("%", "").replace("$", "")))
                    except ValueError:
                        break
            if len(vals) == len(rows):
                numeric_cols.append(col_i)

        return {
            "headers": headers,
            "rows": rows,
            "numeric_cols": numeric_cols,
            "has_numeric": len(numeric_cols) > 0,
        }

    def _parse_sections(self, text: str) -> list[dict]:
        """Parse markdown into hierarchical sections."""
        sections = []
        lines = text.split("\n")
        current_section = None
        current_content = []

        def flush():
            if current_section is not None:
                current_section["content"] = "\n".join(current_content).strip()
                current_section["bullet_points"] = self._extract_bullets(current_section["content"])
                current_section["paragraphs"] = self._extract_paragraphs(current_section["content"])
                sections.append(current_section)

        for line in lines:
            # Heading detection
            m = re.match(r"^(#{1,6})\s+(.+)$", line)
            if m:
                flush()
                level = len(m.group(1))
                heading = m.group(2).strip()
                # Clean markdown from heading
                heading = re.sub(r"[*_`]", "", heading)
                current_section = {
                    "heading": heading,
                    "level": level,
                    "content": "",
                    "bullet_points": [],
                    "paragraphs": [],
                }
                current_content = []
            else:
                if current_section is None and line.strip():
                    # Content before any heading — create implicit intro section
                    current_section = {
                        "heading": "Introduction",
                        "level": 1,
                        "content": "",
                        "bullet_points": [],
                        "paragraphs": [],
                    }
                    current_content = []
                if current_section is not None:
                    current_content.append(line)

        flush()
        return sections

    def _extract_bullets(self, text: str) -> list[str]:
        """Extract bullet/numbered list items from text."""
        bullets = []
        for line in text.split("\n"):
            m = re.match(r"^\s*[-*+•]\s+(.+)$", line)
            if m:
                bullets.append(re.sub(r"[*_](.+?)[*_]", r"\1", m.group(1).strip()))
                continue
            m = re.match(r"^\s*\d+[.)]\s+(.+)$", line)
            if m:
                bullets.append(re.sub(r"[*_](.+?)[*_]", r"\1", m.group(1).strip()))
        return bullets

    def _extract_paragraphs(self, text: str) -> list[str]:
        """Extract clean paragraphs (non-list, non-heading)."""
        paras = []
        for block in re.split(r"\n{2,}", text):
            block = block.strip()
            if not block:
                continue
            # Skip if it's all bullets
            if all(re.match(r"^\s*[-*+•\d]", l) for l in block.split("\n") if l.strip()):
                continue
            # Skip placeholders
            if re.match(r"^__[A-Z_]+_\d+__$", block):
                continue
            # Clean markdown emphasis
            clean = re.sub(r"\*\*(.+?)\*\*", r"\1", block)
            clean = re.sub(r"\*(.+?)\*", r"\1", clean)
            clean = re.sub(r"__(.+?)__", r"\1", clean)
            clean = re.sub(r"_(.+?)_", r"\1", clean)
            clean = re.sub(r"\[([^\]]+)\]\([^\)]+\)", r"\1", clean)
            clean = re.sub(r"\n+", " ", clean).strip()
            if clean and len(clean) > 20:
                paras.append(clean)
        return paras

    def _find_title(self, sections: list[dict], metadata: dict) -> str:
        """Determine document title."""
        if "title" in metadata:
            return metadata["title"]
        for s in sections:
            if s["level"] == 1:
                return s["heading"]
        if sections:
            return sections[0]["heading"]
        return "Presentation"

    def _find_subtitle(self, sections: list[dict]) -> str:
        """Find subtitle from first section's first paragraph."""
        for s in sections:
            if s["level"] == 1 and s["paragraphs"]:
                p = s["paragraphs"][0]
                # Truncate to ~100 chars
                return (p[:100] + "...") if len(p) > 100 else p
        return ""

    def _extract_stats(self, text: str) -> list[dict]:
        """Extract key statistics: numbers with surrounding context."""
        stats = []
        # Pattern: number + unit + context
        patterns = [
            r"(\d+(?:\.\d+)?)\s*(%)\s+([^\n.]{5,50})",
            r"(\$\d+(?:\.\d+)?[KMBkm]?)\s+([^\n.]{5,50})",
            r"(\d+(?:\.\d+)?[KMBkm]?)\s+(users|customers|revenue|growth|increase|decrease|reduction|improvement)\s*([^\n.]{0,40})",
        ]
        seen = set()
        for pat in patterns:
            for m in re.finditer(pat, text, re.IGNORECASE):
                val = m.group(1)
                if val not in seen:
                    seen.add(val)
                    stats.append({
                        "value": val,
                        "context": m.group(0).strip()[:80],
                    })
        return stats[:8]  # Cap at 8 key stats
