#!/usr/bin/env python3
"""
Draftify PPT Generator v3.0

Generates planning documents (기획서) in PowerPoint format,
preserving exact template styles including colors, positions, and layouts.

Usage:
    python generate_ppt.py <project_output_dir> [--template <template_path>]
"""

import argparse
import json
import re
import sys
from datetime import datetime
from pathlib import Path
from typing import Any
from copy import deepcopy

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
    from pptx.oxml.ns import qn
    from pptx.oxml import parse_xml
except ImportError:
    print("Error: python-pptx is required. Install with: pip install python-pptx")
    sys.exit(1)

try:
    from PIL import Image
except ImportError:
    Image = None


# Constants
SCRIPT_DIR = Path(__file__).parent
SKILL_DIR = SCRIPT_DIR.parent
ASSETS_DIR = SKILL_DIR / "assets"
DEFAULT_TEMPLATE = ASSETS_DIR / "ppt_template.pptx"

# Template colors (extracted from template)
COLOR_PRIMARY = RGBColor(0x5E, 0x2B, 0xB8)    # Purple - main accent
COLOR_SECONDARY = RGBColor(0x00, 0x20, 0x60)  # Dark blue - badges
COLOR_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
COLOR_BLACK = RGBColor(0x00, 0x00, 0x00)
COLOR_GRAY = RGBColor(0x66, 0x66, 0x66)

# Template layout names
LAYOUT_COVER = "사용자 지정 레이아웃"
LAYOUT_CONTENT = "4_사용자 지정 레이아웃"
LAYOUT_SECTION = "3_빈화면"
LAYOUT_PROCESS = "6_사용자 지정 레이아웃"
LAYOUT_SCREEN = "7_사용자 지정 레이아웃"
LAYOUT_BLANK = "빈화면"


class MarkdownParser:
    """Parse markdown content into structured data."""

    @staticmethod
    def parse_tables(content: str) -> list[dict]:
        """Extract markdown tables as list of dicts."""
        tables = []
        table_pattern = r'\|(.+)\|\n\|[-:\s|]+\|\n((?:\|.+\|\n?)+)'

        for match in re.finditer(table_pattern, content):
            headers = [h.strip() for h in match.group(1).split('|') if h.strip()]
            rows = []
            for row_line in match.group(2).strip().split('\n'):
                cells = [c.strip() for c in row_line.split('|') if c.strip()]
                if cells:
                    rows.append(cells)
            tables.append({'headers': headers, 'rows': rows})

        return tables

    @staticmethod
    def parse_sections(content: str) -> list[dict]:
        """Parse markdown into sections by headers."""
        sections = []
        current_section = {'level': 0, 'title': '', 'content': ''}

        for line in content.split('\n'):
            if line.startswith('###'):
                if current_section['title']:
                    sections.append(current_section)
                current_section = {'level': 3, 'title': line[3:].strip(), 'content': ''}
            elif line.startswith('##'):
                if current_section['title']:
                    sections.append(current_section)
                current_section = {'level': 2, 'title': line[2:].strip(), 'content': ''}
            elif line.startswith('#'):
                if current_section['title']:
                    sections.append(current_section)
                current_section = {'level': 1, 'title': line[1:].strip(), 'content': ''}
            else:
                current_section['content'] += line + '\n'

        if current_section['title']:
            sections.append(current_section)

        return sections

    @staticmethod
    def strip_markdown(content: str) -> str:
        """Remove markdown formatting for plain text."""
        content = re.sub(r'```[\s\S]*?```', '', content)
        content = re.sub(r'`[^`]+`', '', content)
        content = re.sub(r'\*\*([^*]+)\*\*', r'\1', content)
        content = re.sub(r'\*([^*]+)\*', r'\1', content)
        content = re.sub(r'^#{1,6}\s+', '', content, flags=re.MULTILINE)
        content = re.sub(r'\|[-:\s|]+\|', '', content)
        return content.strip()


class DraftifyPPTGenerator:
    """Generates planning document PPT preserving template styles."""

    def __init__(self, project_dir: Path, template_path: Path | None = None):
        self.project_dir = Path(project_dir)
        self.template_path = template_path or DEFAULT_TEMPLATE
        self.screenshots_dir = self.project_dir / "screenshots"
        self.analysis_dir = self.project_dir / "analysis"
        self.sections_dir = self.project_dir / "sections"
        self.output_path = self.project_dir / "final-draft.pptx"

        self.analyzed_data: dict[str, Any] = {}
        self.sections: dict[str, str] = {}
        self.prs: Presentation | None = None
        self.layouts: dict[str, Any] = {}
        self.parser = MarkdownParser()

    def load_data(self) -> bool:
        """Load analyzed structure and section markdown files."""
        structure_file = self.analysis_dir / "analyzed-structure.json"
        if structure_file.exists():
            with open(structure_file, "r", encoding="utf-8") as f:
                self.analyzed_data = json.load(f)
            print(f"✓ Loaded: {structure_file.name}")
        else:
            print(f"⚠ Warning: {structure_file} not found")
            self.analyzed_data = {"project": {"name": "Unknown Project"}}

        section_files = {
            "glossary": "05-glossary.md",
            "policy": "06-policy-definition.md",
            "process": "07-process-flow.md",
            "screen": "08-screen-definition.md",
        }

        for key, filename in section_files.items():
            filepath = self.sections_dir / filename
            if filepath.exists():
                with open(filepath, "r", encoding="utf-8") as f:
                    self.sections[key] = f.read()
                print(f"✓ Loaded: {filename}")
            else:
                print(f"⚠ Warning: {filename} not found")
                self.sections[key] = ""

        return True

    def create_presentation(self) -> None:
        """Initialize presentation from template."""
        if self.template_path.exists():
            self.prs = Presentation(str(self.template_path))
            print(f"✓ Using template: {self.template_path.name}")

            # Cache layouts by name
            for layout in self.prs.slide_layouts:
                self.layouts[layout.name] = layout

            # Keep first 2 slides (cover, history), delete the rest
            while len(self.prs.slides) > 2:
                rId = self.prs.slides._sldIdLst[2].rId
                self.prs.part.drop_rel(rId)
                del self.prs.slides._sldIdLst[2]
            print(f"✓ Kept cover and history slides from template")

            # Remove all PowerPoint sections (slide grouping)
            self._remove_ppt_sections()
            print(f"✓ Removed PowerPoint sections")
        else:
            self.prs = Presentation()
            self.prs.slide_width = Inches(13.333)
            self.prs.slide_height = Inches(7.5)
            print("⚠ Template not found, using blank presentation")

    def _remove_ppt_sections(self) -> None:
        """Remove all PowerPoint sections (slide grouping) from the presentation."""
        pres_elm = self.prs._element
        # Find and remove sectionLst element (PowerPoint 2010+ sections)
        ns_p14 = '{http://schemas.microsoft.com/office/powerpoint/2010/main}'
        for section_lst in pres_elm.findall(f'.//{ns_p14}sectionLst'):
            section_lst.getparent().remove(section_lst)

    def _get_layout(self, name: str):
        """Get layout by name with fallback."""
        if name in self.layouts:
            return self.layouts[name]
        # Fallback to first available
        return list(self.layouts.values())[0] if self.layouts else self.prs.slide_layouts[0]

    def _add_text_box(self, slide, left: float, top: float, width: float, height: float,
                      text: str, font_size: int = 12, bold: bool = False,
                      font_color: RGBColor = None, fill_color: RGBColor = None,
                      align: PP_ALIGN = PP_ALIGN.LEFT):
        """Add a styled text box."""
        txBox = slide.shapes.add_textbox(
            Inches(left), Inches(top), Inches(width), Inches(height)
        )
        tf = txBox.text_frame
        tf.word_wrap = True

        # Remove auto-fit to allow proper sizing
        tf.auto_size = None

        p = tf.paragraphs[0]
        p.alignment = align

        # Create run explicitly to ensure font settings apply
        run = p.add_run()
        run.text = text
        run.font.size = Pt(font_size)
        run.font.bold = bold

        if font_color:
            run.font.color.rgb = font_color

        if fill_color:
            txBox.fill.solid()
            txBox.fill.fore_color.rgb = fill_color

        return txBox

    def _add_rounded_rect(self, slide, left: float, top: float, width: float, height: float,
                          text: str = "", fill_color: RGBColor = COLOR_PRIMARY,
                          font_size: int = 18, font_color: RGBColor = COLOR_WHITE):
        """Add a rounded rectangle with text (template style section bar)."""
        shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(left), Inches(top), Inches(width), Inches(height)
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
        shape.line.fill.background()  # No border

        if text and shape.has_text_frame:
            tf = shape.text_frame
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            tf.anchor = MSO_ANCHOR.MIDDLE

            # Create run explicitly for proper font settings
            run = p.add_run()
            run.text = text
            run.font.size = Pt(font_size)
            run.font.color.rgb = font_color
            run.font.bold = True

        return shape

    def _add_number_badge(self, slide, left: float, top: float, number: str,
                          size: float = 0.39, font_size: int = 14,
                          fill_color: RGBColor = None, font_color: RGBColor = None):
        """Add a small numbered badge (template style - rectangle with gray fill, no border)."""
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(left), Inches(top), Inches(size), Inches(size)
        )
        # Fill color (default: light gray like template)
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color if fill_color else RGBColor(0xF2, 0xF2, 0xF2)
        # No border
        shape.line.fill.background()

        if shape.has_text_frame:
            tf = shape.text_frame
            # Match template text frame settings
            tf.word_wrap = False  # Disable text wrapping (template setting)

            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            tf.anchor = MSO_ANCHOR.MIDDLE

            # Create run - NOT bold for default style (template uses regular weight)
            run = p.add_run()
            run.text = number
            run.font.size = Pt(font_size)
            run.font.color.rgb = font_color if font_color else COLOR_BLACK
            # Only bold if custom fill_color is provided (for section badges)
            run.font.bold = True if fill_color else False

        return shape

    def _add_line(self, slide, start_x: float, start_y: float, end_x: float, end_y: float,
                  color: RGBColor = COLOR_GRAY, width: float = 1.0):
        """Add a line shape."""
        connector = slide.shapes.add_connector(
            1,  # Straight connector
            Inches(start_x), Inches(start_y),
            Inches(end_x), Inches(end_y)
        )
        connector.line.color.rgb = color
        connector.line.width = Pt(width)
        return connector

    def _create_table(self, slide, left: float, top: float, width: float, height: float,
                      headers: list, rows: list):
        """Create a styled table matching template."""
        cols = len(headers)
        row_count = len(rows) + 1

        table_shape = slide.shapes.add_table(
            row_count, cols,
            Inches(left), Inches(top),
            Inches(width), Inches(height)
        )
        table = table_shape.table

        col_width = Inches(width / cols)
        for i in range(cols):
            table.columns[i].width = col_width

        # Header row styling
        for i, header in enumerate(headers):
            cell = table.cell(0, i)
            cell.text = header
            # Header background (purple)
            cell.fill.solid()
            cell.fill.fore_color.rgb = COLOR_PRIMARY
            for p in cell.text_frame.paragraphs:
                p.font.bold = True
                p.font.size = Pt(10)
                p.font.color.rgb = COLOR_WHITE

        # Data rows
        for row_idx, row_data in enumerate(rows):
            for col_idx, cell_text in enumerate(row_data):
                if col_idx < cols:
                    cell = table.cell(row_idx + 1, col_idx)
                    cell.text = str(cell_text)
                    for p in cell.text_frame.paragraphs:
                        p.font.size = Pt(9)

        return table

    # === Slide Generation Methods ===

    def add_cover_slide(self) -> None:
        """Update existing cover slide from template - only change title and date."""
        # Use the first slide from template (already preserved in create_presentation)
        slide = self.prs.slides[0]

        project_name = self.analyzed_data.get("project", {}).get("name", "기획서")
        version = self.analyzed_data.get("project", {}).get("version", "1.0")
        date_str = datetime.now().strftime("%Y-%m-%d")

        # Find and update existing shapes by name
        for shape in slide.shapes:
            if shape.name == "제목 3" and shape.has_text_frame:
                # Update main title - clear all runs, set text in first run only
                tf = shape.text_frame
                if tf.paragraphs:
                    p = tf.paragraphs[0]
                    # Template has multiple runs: "Draftify" + " " + "기획서"
                    # Clear all runs first, then set first run
                    for i, run in enumerate(p.runs):
                        if i == 0:
                            run.text = f"{project_name} 기획서"
                        else:
                            run.text = ""

            elif shape.name == "부제목 4" and shape.has_text_frame:
                # Update date in the subtitle text box (3 lines format)
                tf = shape.text_frame
                for p in tf.paragraphs:
                    if "Last update" in p.text:
                        # Template: "Last update " + ": YYYY-MM-DD"
                        for i, run in enumerate(p.runs):
                            if i == 0:
                                run.text = f"Last update : {date_str}"
                            else:
                                run.text = ""
                    elif "Version" in p.text:
                        if p.runs:
                            p.runs[0].text = f"Version: {version}"

        print("✓ Updated: Cover slide (template preserved)")

    def add_history_slide(self) -> None:
        """Update existing history slide from template - only update table row 1."""
        # Use the second slide from template (already preserved in create_presentation)
        slide = self.prs.slides[1]

        version = self.analyzed_data.get("project", {}).get("version", "1.0")
        author = self.analyzed_data.get("project", {}).get("author", "Draftify")
        date_str = datetime.now().strftime("%Y.%m.%d")

        # Find the table and update first data row
        for shape in slide.shapes:
            if shape.has_table:
                table = shape.table
                # Row 0 is header, Row 1 is first data row
                if len(table.rows) > 1:
                    row = table.rows[1]
                    # Columns: ver, 변경일, 변경내용, 작성자, 비고
                    new_values = [version, date_str, "기획서 자동 생성", author, ""]
                    for ci, cell in enumerate(row.cells):
                        if ci < len(new_values):
                            # Preserve formatting by updating run text, not cell.text
                            self._update_cell_text(cell, new_values[ci])
                break

        print("✓ Updated: History slide (template preserved)")

    def _update_cell_text(self, cell, new_text: str) -> None:
        """Update cell text while preserving formatting."""
        tf = cell.text_frame
        if tf.paragraphs:
            p = tf.paragraphs[0]
            if p.runs:
                # Clear all runs except first, set new text in first run
                for i, run in enumerate(p.runs):
                    if i == 0:
                        run.text = new_text
                    else:
                        run.text = ""
            else:
                # No runs exist, set paragraph text (will lose formatting)
                p.text = new_text
        else:
            cell.text = new_text

    def _extract_section_headers(self, section_key: str) -> list[str]:
        """Extract ## level headers from section markdown."""
        content = self.sections.get(section_key, "")
        headers = []
        for line in content.split('\n'):
            # Match ## headers (not ### or deeper)
            if line.startswith('## ') and not line.startswith('### '):
                # Extract just the title part, remove numbering like "1." or "## 1."
                header = line[3:].strip()
                # Remove leading numbers like "1. " or "1) "
                header = re.sub(r'^[\d]+[\.\)]\s*', '', header)
                if header:
                    headers.append(header)
        return headers

    def add_contents_slide(self) -> None:
        """Add contents/TOC slide matching template style with sub-items."""
        slide = self.prs.slides.add_slide(self._get_layout(LAYOUT_CONTENT))

        project_name = self.analyzed_data.get("project", {}).get("name", "서비스")

        # "CONTENTS" title - 32pt Bold, Purple (from template)
        self._add_text_box(slide, 1.27, 1.24, 2.44, 0.64, "CONTENTS",
                          font_size=32, bold=True, font_color=COLOR_PRIMARY)

        # Project subtitle - 16pt (from template)
        self._add_text_box(slide, 1.33, 1.86, 4.0, 0.37,
                          f"{project_name} 정책 및 프로세스",
                          font_size=16, font_color=COLOR_WHITE,
                          fill_color=COLOR_PRIMARY)

        # Extract sub-items from sections
        policy_subs = self._extract_section_headers("policy")
        process_subs = self._extract_section_headers("process")
        screen_subs = self._extract_section_headers("screen")

        # Build TOC items with sub-items
        toc_items = [
            ("01", "용어 정의", []),
            ("02", "정책 정의", policy_subs),
            ("03", "프로세스 흐름도", process_subs),
            ("04", "화면상세", screen_subs),
        ]

        # Calculate total height needed
        total_items = len(toc_items)
        for _, _, subs in toc_items:
            total_items += min(len(subs), 6)  # Cap at 6 sub-items per section

        # Layout: 2 columns
        col1_x = 1.5   # Left column x position
        col2_x = 7.0   # Right column x position
        y_start = 2.6
        line_height_main = 0.5
        line_height_sub = 0.32
        badge_size = 0.39

        current_col = 1
        y = y_start

        for num, title, subs in toc_items:
            # Check if we need to switch to column 2
            items_height = line_height_main + (min(len(subs), 6) * line_height_sub)
            if y + items_height > 6.5 and current_col == 1:
                current_col = 2
                y = y_start

            x = col1_x if current_col == 1 else col2_x

            # Number badge - 14pt
            self._add_number_badge(slide, x, y, num, size=badge_size)
            # Main title - 14pt
            self._add_text_box(slide, x + 0.5, y, 3.0, 0.4, title, font_size=14)
            y += line_height_main

            # Sub-items - 12pt, indented
            for i, sub in enumerate(subs[:6]):  # Max 6 sub-items
                # Truncate long names
                display_sub = sub[:25] + "..." if len(sub) > 25 else sub
                self._add_text_box(slide, x + 0.6, y, 3.5, 0.3,
                                  f"· {display_sub}", font_size=12, font_color=COLOR_GRAY)
                y += line_height_sub

            # Add spacing between sections
            y += 0.15

        print("✓ Added: Contents slide")

    def add_section_divider(self, section_num: str, title: str, subtitles: list = None) -> None:
        """Add section divider slide with purple bar (template style)."""
        slide = self.prs.slides.add_slide(self._get_layout(LAYOUT_SECTION))

        # Main section bar (purple rounded rectangle)
        # Position from template: 2.74, 2.46, size: 7.86x1.07
        # Template: 32pt, number in cyan (#08D2D9), title in white
        self._add_section_title_bar(
            slide, 2.74, 2.46, 7.86, 1.07,
            section_num, title
        )

        # Subtitles with number badges (template: 16pt)
        if subtitles:
            y_start = 4.16  # From template position
            for i, subtitle in enumerate(subtitles[:6], 1):
                y = y_start + (i - 1) * 0.52
                self._add_number_badge(slide, 5.93, y, f"{i:02d}", size=0.31)
                self._add_text_box(slide, 6.30, y - 0.03, 2.95, 0.37,
                                  subtitle, font_size=16)

        print(f"✓ Added: Section divider - {title}")

    def _add_section_title_bar(self, slide, left: float, top: float,
                                width: float, height: float,
                                section_num: str, title: str) -> None:
        """Add section title bar with colored number (template style)."""
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,  # Template uses rectangle, not rounded
            Inches(left), Inches(top), Inches(width), Inches(height)
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = COLOR_PRIMARY
        shape.line.fill.background()

        if shape.has_text_frame:
            tf = shape.text_frame
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            tf.anchor = MSO_ANCHOR.MIDDLE

            # Number in cyan (#08D2D9)
            run_num = p.add_run()
            run_num.text = f"{section_num} "
            run_num.font.size = Pt(32)
            run_num.font.color.rgb = RGBColor(0x08, 0xD2, 0xD9)  # Cyan
            run_num.font.bold = True

            # Title in white
            run_title = p.add_run()
            run_title.text = title
            run_title.font.size = Pt(32)
            run_title.font.color.rgb = COLOR_WHITE
            run_title.font.bold = True

    def add_glossary_slides(self) -> None:
        """Add glossary slides with proper tables."""
        self.add_section_divider("01", "용어 정의")

        content = self.sections.get("glossary", "")
        if not content:
            return

        tables = self.parser.parse_tables(content)

        slide = self.prs.slides.add_slide(self._get_layout(LAYOUT_CONTENT))

        # Title with section number badge
        self._add_number_badge(slide, 0.40, 0.56, "1", size=0.28,
                              fill_color=COLOR_PRIMARY, font_color=COLOR_WHITE)
        self._add_text_box(slide, 0.75, 0.50, 4.0, 0.40,
                          "용어 정의", font_size=18, bold=True)

        y_pos = 1.25
        for table_data in tables[:2]:
            if table_data['headers'] and table_data['rows']:
                rows_to_show = table_data['rows'][:12]
                row_height = 0.32 * (len(rows_to_show) + 1)
                self._create_table(
                    slide, 0.40, y_pos, 12.80, row_height,
                    table_data['headers'], rows_to_show
                )
                y_pos += row_height + 0.3

        print("✓ Added: Glossary slides")

    def add_policy_slides(self) -> None:
        """Add policy definition slides."""
        content = self.sections.get("policy", "")
        sections = self.parser.parse_sections(content)
        policy_groups = [s for s in sections if s['level'] == 2]

        # Extract subtitles, removing leading numbers like "1. "
        subtitles = []
        for pg in policy_groups[:6]:
            title = pg['title'].split('(')[0].strip()
            title = re.sub(r'^[\d]+[\.\)]\s*', '', title)  # Remove leading number
            subtitles.append(title[:20])
        self.add_section_divider("02", "정책 정의", subtitles)

        for idx, pg in enumerate(policy_groups[:6], 1):
            slide = self.prs.slides.add_slide(self._get_layout(LAYOUT_CONTENT))

            # Title with number badge
            self._add_number_badge(slide, 0.40, 0.56, str(idx), size=0.28,
                                  fill_color=COLOR_PRIMARY, font_color=COLOR_WHITE)

            title_text = pg['title'][:40]
            self._add_text_box(slide, 0.75, 0.50, 10.0, 0.40,
                              f"{idx}. {title_text}", font_size=16, bold=True)

            tables = self.parser.parse_tables(pg['content'])
            y_pos = 1.25

            if tables:
                for table_data in tables[:2]:
                    if table_data['headers'] and table_data['rows']:
                        rows_to_show = table_data['rows'][:10]
                        row_height = 0.30 * (len(rows_to_show) + 1)
                        self._create_table(
                            slide, 0.40, y_pos, 12.80, row_height,
                            table_data['headers'], rows_to_show
                        )
                        y_pos += row_height + 0.2
            else:
                clean_content = self.parser.strip_markdown(pg['content'])[:1800]
                self._add_text_box(slide, 0.40, y_pos, 12.80, 5.0,
                                  clean_content, font_size=10)

        print("✓ Added: Policy slides")

    def add_process_slides(self) -> None:
        """Add process flow slides."""
        content = self.sections.get("process", "")
        sections = self.parser.parse_sections(content)
        process_sections = [s for s in sections if s['level'] == 2]

        # Extract subtitles, removing leading numbers like "1. "
        subtitles = []
        for ps in process_sections[:6]:
            title = ps['title'].split('(')[0].strip()
            title = re.sub(r'^[\d]+[\.\)]\s*', '', title)
            subtitles.append(title[:20])
        self.add_section_divider("03", "프로세스 흐름도", subtitles)

        for idx, ps in enumerate(process_sections[:4], 1):
            slide = self.prs.slides.add_slide(self._get_layout(LAYOUT_PROCESS))

            # Title
            self._add_number_badge(slide, 0.40, 0.56, str(idx), size=0.28,
                                  fill_color=COLOR_PRIMARY, font_color=COLOR_WHITE)
            self._add_text_box(slide, 0.75, 0.50, 10.0, 0.40,
                              f"3-{idx}. {ps['title'][:35]}", font_size=16, bold=True)

            # Process content (ASCII/text flow)
            clean_content = ps['content'][:2500]
            self._add_text_box(slide, 0.40, 1.20, 12.80, 5.8,
                              clean_content, font_size=9)

        print("✓ Added: Process flow slides")

    def add_screen_slides(self) -> None:
        """Add screen definition slides matching template layout."""
        screens = self.analyzed_data.get("screens", [])

        subtitles = [f"{s.get('id')}: {s.get('name')[:15]}" for s in screens[:6]]
        self.add_section_divider("04", "화면상세", subtitles)

        # Screen list summary slide
        if screens:
            slide = self.prs.slides.add_slide(self._get_layout(LAYOUT_CONTENT))
            self._add_number_badge(slide, 0.40, 0.56, "4", size=0.28,
                                  fill_color=COLOR_PRIMARY, font_color=COLOR_WHITE)
            self._add_text_box(slide, 0.75, 0.50, 4.0, 0.40,
                              "화면 목록", font_size=18, bold=True)

            headers = ["화면 ID", "화면명", "목적", "URL"]
            rows = [[
                s.get("id", ""),
                s.get("name", ""),
                s.get("purpose", "")[:25],
                s.get("url", "")
            ] for s in screens[:12]]

            self._create_table(slide, 0.40, 1.25, 12.80, 0.35 * (len(rows) + 1),
                              headers, rows)

        # Individual screen slides
        screen_content = self.sections.get("screen", "")
        screen_sections = self._parse_screen_sections(screen_content)

        for idx, screen in enumerate(screens, 1):
            screen_id = screen.get("id", "")
            screen_name = screen.get("name", "")

            slide = self.prs.slides.add_slide(self._get_layout(LAYOUT_SCREEN))

            # Section title at top
            self._add_text_box(slide, 0.0, 0.10, 11.50, 0.40,
                              "4. 화면상세", font_size=14, bold=True)

            # Screen name with badge
            self._add_number_badge(slide, 0.40, 0.56, str(idx), size=0.28,
                                  fill_color=COLOR_PRIMARY, font_color=COLOR_WHITE)
            self._add_text_box(slide, 0.75, 0.50, 8.0, 0.40,
                              f"{screen_id}: {screen_name}", font_size=14, bold=True)

            # Layout: Left ~24cm for images, Right ~7.8cm for description
            # Template: Description at x=25.8cm (10.16in), width=7.8cm (3.07in)
            desc_left = 10.2   # 25.9cm / 2.54
            desc_width = 3.07  # 7.8cm / 2.54

            # "Description" header bar (template style - light gray RGB(240,240,240), 9pt, not bold)
            desc_header = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                Inches(desc_left), Inches(0.15), Inches(desc_width), Inches(0.28)
            )
            desc_header.fill.solid()
            desc_header.fill.fore_color.rgb = RGBColor(0xF0, 0xF0, 0xF0)  # RGB(240,240,240)
            desc_header.line.fill.background()
            if desc_header.has_text_frame:
                tf = desc_header.text_frame
                p = tf.paragraphs[0]
                p.alignment = PP_ALIGN.CENTER
                run = p.add_run()
                run.text = "Description"
                run.font.size = Pt(9)  # Template: 9pt
                run.font.color.rgb = COLOR_BLACK  # Black text on light gray
                run.font.bold = False  # Template: not bold

            # Screenshot on left side (within ~24cm area)
            screenshot_path = self.screenshots_dir / f"{screen_id}.png"
            if screenshot_path.exists():
                try:
                    slide.shapes.add_picture(
                        str(screenshot_path),
                        Inches(0.7), Inches(1.6),
                        width=Inches(4.5)
                    )
                except Exception as e:
                    print(f"  ⚠ Screenshot error for {screen_id}: {e}")

            # Description table (template style: 2 cols - number | description)
            # Template has narrow first column for numbers
            self._create_description_table(
                slide, desc_left, 0.45, desc_width,
                screen_id, screen_name, screen, screen_sections
            )

        print("✓ Added: Screen definition slides")

    def _parse_screen_sections(self, content: str) -> dict[str, str]:
        """Parse screen sections from markdown."""
        sections = {}
        pattern = r"##\s*(?:화면\s*(?:정의)?:?\s*)?(SCR-\d{3})[^\n]*\n(.*?)(?=##\s*(?:화면|SCR-)|$)"
        matches = re.findall(pattern, content, re.DOTALL | re.IGNORECASE)
        for match in matches:
            sections[match[0]] = match[1].strip()
        return sections

    def _create_description_table(self, slide, left: float, top: float, width: float,
                                   screen_id: str, screen_name: str, screen: dict,
                                   screen_sections: dict) -> None:
        """Create description table in template style (number | description format)."""
        # Template style: 2 columns - narrow number col, wide description col
        # Row 0: Reference links
        # Row 1+: Numbered descriptions

        # Build description content
        purpose = screen.get("purpose", "-")
        url = screen.get("url", "-")
        process_step = screen.get("processStep", "-")

        # Get additional content from markdown
        extra_content = ""
        if screen_id in screen_sections:
            extra_content = self.parser.strip_markdown(screen_sections[screen_id][:300])

        # Create table with template structure
        rows_data = [
            ("📑참고", f"화면 ID: {screen_id}\n화면명: {screen_name}"),
            ("1", f"목적: {purpose}\nURL: {url}"),
            ("2", f"프로세스 단계: {process_step}"),
        ]

        if extra_content:
            rows_data.append(("3", extra_content[:150]))

        # Create the table
        table_height = 0.6 * len(rows_data)
        table = slide.shapes.add_table(
            len(rows_data), 2,
            Inches(left), Inches(top), Inches(width), Inches(table_height)
        ).table

        # Set column widths (narrow first col for numbers)
        table.columns[0].width = Inches(0.35)
        table.columns[1].width = Inches(width - 0.35)

        # Fill table data with proper text colors
        for ri, (num, desc) in enumerate(rows_data):
            # Number cell
            cell0 = table.cell(ri, 0)
            cell0.text = num
            for para in cell0.text_frame.paragraphs:
                para.alignment = PP_ALIGN.CENTER
                for run in para.runs:
                    run.font.size = Pt(9)
                    run.font.bold = True
                    run.font.color.rgb = COLOR_BLACK  # Ensure black text

            # Description cell
            cell1 = table.cell(ri, 1)
            cell1.text = desc
            for para in cell1.text_frame.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(8)
                    run.font.color.rgb = COLOR_BLACK  # Ensure black text

        # Style table (template: transparent fill, borders only top/bottom/inside at 1/4pt)
        num_rows = len(rows_data)
        for ri in range(num_rows):
            for ci in range(2):
                cell = table.cell(ri, ci)
                # Transparent fill
                cell.fill.background()

                # Borders: top, bottom, inside only (no left/right outer edges)
                # 1/4pt = 3175 EMU
                self._set_cell_borders_selective(cell, 3175, "808080", ri, ci, num_rows, 2)

    def _set_cell_borders_selective(self, cell, width_emu: int, color_hex: str,
                                      row: int, col: int, total_rows: int, total_cols: int) -> None:
        """Set cell borders selectively (top/bottom/inside only, no outer left/right).

        Matches template XML structure exactly:
        <a:lnT w="6350" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill><a:srgbClr val="808080"/></a:solidFill>
          <a:prstDash val="solid"/>
          <a:round/>
          <a:headEnd type="none" w="med" len="med"/>
          <a:tailEnd type="none" w="med" len="med"/>
        </a:lnT>
        """
        from pptx.oxml.ns import qn
        from lxml import etree

        tc = cell._tc
        tcPr = tc.find(qn('a:tcPr'))
        if tcPr is None:
            tcPr = etree.SubElement(tc, qn('a:tcPr'))

        # Determine which borders to show
        borders_to_set = []
        borders_to_set.append('lnT')  # Top border always
        borders_to_set.append('lnB')  # Bottom border always

        # Inside vertical borders (between columns)
        if col < total_cols - 1:
            borders_to_set.append('lnR')  # Right border if not last column
        if col > 0:
            borders_to_set.append('lnL')  # Left border if not first column

        # Set borders with complete XML structure matching template
        for border_name in ['lnL', 'lnR', 'lnT', 'lnB']:
            border = tcPr.find(qn(f'a:{border_name}'))

            if border_name in borders_to_set:
                # Remove existing border if present and recreate
                if border is not None:
                    tcPr.remove(border)

                # Create new border element with all attributes
                border = etree.SubElement(tcPr, qn(f'a:{border_name}'))
                border.set('w', str(width_emu))
                border.set('cap', 'flat')
                border.set('cmpd', 'sng')
                border.set('algn', 'ctr')

                # solidFill with srgbClr
                solidFill = etree.SubElement(border, qn('a:solidFill'))
                srgbClr = etree.SubElement(solidFill, qn('a:srgbClr'))
                srgbClr.set('val', color_hex)

                # prstDash - solid line
                prstDash = etree.SubElement(border, qn('a:prstDash'))
                prstDash.set('val', 'solid')

                # round - rounded line join
                etree.SubElement(border, qn('a:round'))

                # headEnd - line start cap
                headEnd = etree.SubElement(border, qn('a:headEnd'))
                headEnd.set('type', 'none')
                headEnd.set('w', 'med')
                headEnd.set('len', 'med')

                # tailEnd - line end cap
                tailEnd = etree.SubElement(border, qn('a:tailEnd'))
                tailEnd.set('type', 'none')
                tailEnd.set('w', 'med')
                tailEnd.set('len', 'med')
            else:
                # Remove border completely or set to noFill
                if border is not None:
                    tcPr.remove(border)
                # Create empty border with noFill
                border = etree.SubElement(tcPr, qn(f'a:{border_name}'))
                border.set('w', '0')
                etree.SubElement(border, qn('a:noFill'))

    def add_reference_slide(self) -> None:
        """Add reference slide."""
        slide = self.prs.slides.add_slide(self._get_layout(LAYOUT_CONTENT))

        self._add_text_box(slide, 0.35, 0.55, 3.0, 0.64,
                          "참고 문헌", font_size=24, bold=True)

        refs = self.analyzed_data.get("references", [])
        ref_text = "\n".join([f"• {ref}" for ref in refs]) if refs else "• 자동 생성된 기획서\n• Draftify 크롤링 데이터 기반"

        self._add_text_box(slide, 0.51, 1.5, 12.42, 4.0, ref_text, font_size=12)

        print("✓ Added: Reference slide")

    def add_eod_slide(self) -> None:
        """Add end of document slide."""
        slide = self.prs.slides.add_slide(self._get_layout(LAYOUT_BLANK))

        # Centered "End of document" text (from template: 3.34, 3.14)
        self._add_text_box(slide, 3.34, 3.14, 6.82, 0.61,
                          "End of document",
                          font_size=28, bold=True, font_color=COLOR_GRAY,
                          align=PP_ALIGN.CENTER)

        print("✓ Added: EOD slide")

    def generate(self) -> Path:
        """Generate the complete PPT document."""
        print(f"\n{'='*50}")
        print("Draftify PPT Generator v3.0")
        print(f"{'='*50}\n")

        print("Loading data...")
        self.load_data()

        print("\nCreating presentation...")
        self.create_presentation()

        print("\nGenerating slides...")
        self.add_cover_slide()
        self.add_history_slide()
        self.add_contents_slide()
        self.add_glossary_slides()
        self.add_policy_slides()
        self.add_process_slides()
        self.add_screen_slides()
        self.add_reference_slide()
        self.add_eod_slide()

        print(f"\nSaving to: {self.output_path}")
        self.prs.save(str(self.output_path))

        print(f"\n{'='*50}")
        print(f"✓ SUCCESS: Generated {self.output_path}")
        print(f"  Total slides: {len(self.prs.slides)}")
        print(f"{'='*50}\n")

        return self.output_path


def main():
    parser = argparse.ArgumentParser(
        description="Generate planning document PPT from analyzed project data"
    )
    parser.add_argument(
        "project_dir",
        help="Path to project output directory (e.g., outputs/my-project/)"
    )
    parser.add_argument(
        "--template",
        help="Path to PPT template file",
        default=None
    )

    args = parser.parse_args()

    project_dir = Path(args.project_dir)
    if not project_dir.exists():
        print(f"Error: Project directory not found: {project_dir}")
        sys.exit(1)

    template_path = Path(args.template) if args.template else None

    generator = DraftifyPPTGenerator(project_dir, template_path)
    generator.generate()


if __name__ == "__main__":
    main()
