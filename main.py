#!/usr/bin/env python3
"""
md2pptx — Markdown to PowerPoint Converter
A modular, AI-powered system for generating professional PPTX presentations from Markdown.

Usage:
    python main.py input.md [output.pptx] [--slides 10-15] [--theme dark|light]
"""

import argparse
import sys
import os
import time
from pathlib import Path

from src.parser.md_parser import MarkdownParser
from src.structurer.content_structurer import ContentStructurer
from src.formatter.slide_planner import SlidePlanner
from src.renderer.pptx_renderer import PPTXRenderer
from src.charts.chart_generator import ChartGenerator


def main():
    parser = argparse.ArgumentParser(
        description="Convert Markdown to professional PowerPoint presentations",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python main.py report.md
  python main.py report.md output/deck.pptx --slides 12
  python main.py report.md --theme dark
        """
    )
    parser.add_argument("input", help="Input Markdown file (.md)")
    parser.add_argument("output", nargs="?", help="Output PPTX file (default: <input>.pptx)")
    parser.add_argument("--slides", type=int, default=12,
                        help="Target slide count (10–15, default: 12)")
    parser.add_argument("--theme", choices=["dark", "light", "auto"], default="auto",
                        help="Presentation theme (default: auto)")
    parser.add_argument("--verbose", "-v", action="store_true",
                        help="Verbose output")

    args = parser.parse_args()

    # Validate input
    input_path = Path(args.input)
    if not input_path.exists():
        print(f"❌ Error: Input file not found: {args.input}")
        sys.exit(1)

    if input_path.stat().st_size > 5 * 1024 * 1024:
        print("❌ Error: Input file exceeds 5 MB limit")
        sys.exit(1)

    # Validate slide count
    slide_count = max(10, min(15, args.slides))
    if slide_count != args.slides:
        print(f"⚠️  Slide count clamped to {slide_count} (valid range: 10–15)")

    # Determine output path
    if args.output:
        output_path = Path(args.output)
    else:
        output_path = input_path.with_suffix(".pptx")

    output_path.parent.mkdir(parents=True, exist_ok=True)

    print(f"\n🚀 md2pptx — Markdown to PowerPoint Converter")
    print(f"{'─' * 50}")
    print(f"  Input:  {input_path}")
    print(f"  Output: {output_path}")
    print(f"  Slides: {slide_count}")
    print(f"  Theme:  {args.theme}")
    print(f"{'─' * 50}\n")

    t0 = time.time()

    try:
        # ── Step 1: Parse Markdown ──────────────────────────────────────────
        print("📖 [1/4] Parsing markdown content...")
        md_parser = MarkdownParser(verbose=args.verbose)
        document = md_parser.parse(input_path.read_text(encoding="utf-8"))
        print(f"   ✓ Parsed {len(document['sections'])} sections, "
              f"{document['stats']['word_count']} words, "
              f"{len(document['tables'])} tables, "
              f"{len(document['code_blocks'])} code blocks")

        # ── Step 2: Structure content ───────────────────────────────────────
        print("🧠 [2/4] Structuring content into slide narrative...")
        structurer = ContentStructurer(target_slides=slide_count, verbose=args.verbose)
        slide_plan = structurer.structure(document)
        print(f"   ✓ Generated {len(slide_plan)} slide outlines")

        # ── Step 3: Plan slide layouts ──────────────────────────────────────
        print("🎨 [3/4] Planning slide layouts and visual elements...")
        planner = SlidePlanner(theme=args.theme, verbose=args.verbose)
        designed_slides = planner.plan(slide_plan, document)
        print(f"   ✓ Designed {len(designed_slides)} slides with visual layouts")

        # ── Step 4: Render PPTX ─────────────────────────────────────────────
        print("✨ [4/4] Rendering PowerPoint presentation...")
        chart_gen = ChartGenerator()
        renderer = PPTXRenderer(chart_generator=chart_gen, verbose=args.verbose)
        renderer.render(designed_slides, str(output_path))

        elapsed = time.time() - t0
        file_size = output_path.stat().st_size / 1024

        print(f"\n{'─' * 50}")
        print(f"✅ Success! Presentation generated in {elapsed:.1f}s")
        print(f"   📊 {len(designed_slides)} slides  |  {file_size:.0f} KB")
        print(f"   📁 {output_path.resolve()}")
        print(f"{'─' * 50}\n")

    except KeyboardInterrupt:
        print("\n⚠️  Interrupted by user")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ Error: {e}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
