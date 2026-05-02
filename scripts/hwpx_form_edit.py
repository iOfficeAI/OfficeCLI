#!/usr/bin/env python3
"""HWPX Korean Document Pattern Matching and Editing Prototype.

Classifies HWPX documents (exam, regulation, form, report, mixed) and
provides extraction/editing utilities for Korean government forms, exam
papers, regulations, and application documents.

Usage:
    python hwpx_form_edit.py classify doc.hwpx
    python hwpx_form_edit.py hierarchy doc.hwpx
    python hwpx_form_edit.py appendix doc.hwpx
    python hwpx_form_edit.py strip-lineseg doc.hwpx output.hwpx
    python hwpx_form_edit.py extract doc.hwpx
    python hwpx_form_edit.py digit-headings doc.hwpx
    python hwpx_form_edit.py pages doc.hwpx
    python hwpx_form_edit.py problems doc.hwpx
    python hwpx_form_edit.py incell doc.hwpx
    python hwpx_form_edit.py markers doc.hwpx
    python hwpx_form_edit.py headers-footers doc.hwpx
    python hwpx_form_edit.py fill doc.hwpx output.hwpx '성명=홍길동,주소=서울'
"""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import sys
import tempfile
import zipfile
from typing import Any

# Security: Use defusedxml if available (XXE defense)
try:
    import defusedxml.ElementTree as ET
except ImportError:
    import xml.etree.ElementTree as ET
    import warnings
    warnings.warn(
        "defusedxml not available - using stdlib ElementTree. "
        "Install defusedxml for enhanced security: pip install defusedxml",
        ImportWarning,
        stacklevel=2,
    )

# ---------------------------------------------------------------------------
# HWPX Namespaces
# ---------------------------------------------------------------------------

NS = {
    "hp": "urn:hancom:hwpml:2011:paragraph",
    "hs": "urn:hancom:hwpml:2011:section",
    "hh": "urn:hancom:hwpml:2011:head",
}

# Fallback namespace variants (some documents use http:// or 2016 URIs)
NS_ALT = {
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hs": "http://www.hancom.co.kr/hwpml/2011/section",
    "hh": "http://www.hancom.co.kr/hwpml/2011/head",
}

# ---------------------------------------------------------------------------
# Compiled Regex Patterns (R1-R25)
# ---------------------------------------------------------------------------

# -- Tier 1: Structure Detection --

# R1: Chapter/section heading  (제1장 총칙, 제2절 ...)
R1_CHAPTER_HEADING = re.compile(r"^제\s*(\d+)\s*[장절편관]\s*(.+)")

# R2: Article  (제1조(목적), 제3조의2(특례))
R2_ARTICLE = re.compile(
    r"^제\s*(\d+)\s*조(?:\s*의\s*(\d+))?\s*[((]\s*(.+?)\s*[))]"
)

# R3: Circled number item  (① 항목 ...)
R3_CIRCLED_NUMBER = re.compile(
    r"^\s*[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]\s*(.+)"
)

# R4: Numbered list  (1. 항목)
R4_NUMBERED_LIST = re.compile(r"^\s*(\d{1,2})\.\s+(.+)")

# R5: Korean letter list  (가. 항목)
R5_KOREAN_LETTER = re.compile(
    r"^\s*[가나다라마바사아자차카타파하]\.\s*(.+)"
)

# -- Tier 2: Form Patterns --

# R6: Checkbox flat  (□ 항목, ■ 항목)
R6_CHECKBOX_FLAT = re.compile(r"^\s*[□■☐☑]\s*(.+)")

# R7: Inline checkbox group  (구분: □ A □ B □ C)
R7_CHECKBOX_GROUP = re.compile(
    r"^(.+?)[\s:：]\s*[□■]\s*(.+?)(?:\s*[□■]\s*(.+?))*"
)

# R8: Appendix reference  ([별첨 제1호], [별지], [별표 2])
R8_APPENDIX_REF = re.compile(r"\[별[첨지표]\s*(?:제?\s*(\d+)\s*호?)?\]")

# R9: Digit-concatenated heading  (3지원금 집행기준)
R9_DIGIT_HEADING = re.compile(r"^(\d{1,2})([가-힣])")

# R10: Label-colon-value  (성명: 홍길동)
R10_LABEL_COLON_VALUE = re.compile(r"([가-힣]{2,6})\s*[:：]\s*(.+)")

# -- Tier 3: Content Patterns --

# R11: Date  (2024.03.15, 2024-3-15, 2024년 3월 15일)
R11_DATE = re.compile(
    r"\d{4}[.\-/년]\s*\d{1,2}[.\-/월]\s*\d{1,2}[일]?"
)

# R12: Currency amount  (1,000,000 원)
R12_CURRENCY = re.compile(r"[\d,]+\s*원")

# R13: Phone number  (02-1234-5678, 010-1234-5678)
R13_PHONE = re.compile(r"\d{2,3}-\d{3,4}-\d{4}")

# R14: Resident registration number  (880101-1234567)
R14_RRN = re.compile(r"\d{6}-[1-4]\d{6}")

# R15: Checkbox hierarchy markers (□=0, ○=1, -=2, *=3)
R15_CHECKBOX_HIERARCHY = re.compile(r"^([□○●◎\-\*])\s*(.+)")

# R16: Appendix ref (same as R8, kept as alias for Tier 3 grouping)
R16_APPENDIX_REF = R8_APPENDIX_REF

# R17: Digit heading (same as R9, kept as alias for Tier 3 grouping)
R17_DIGIT_HEADING = R9_DIGIT_HEADING

# -- Tier 3: Shared Utilities --

# R18: Whitespace collapse
R18_WHITESPACE = re.compile(r"\s+")

# R19: Trailing colon strip
R19_TRAILING_COLON = re.compile(r"[:：]\s*$")

# R20: Short Korean label heuristic  (2-8 chars, Korean+spaces+parens)
R20_SHORT_KOREAN_LABEL = re.compile(r"^[\uAC00-\uD7A3\s()·]{2,8}$")

# R21: Checkbox prefix strip
R21_CHECKBOX_PREFIX = re.compile(r"^[□○●◎\-\*]\s*")

# R22: Chapter/section number extract
R22_CHAPTER_NUM = re.compile(r"제(\d+)[장절편]\s")

# R23: Article number extract
R23_ARTICLE_NUM = re.compile(r"제(\d+)조")

# R24: Parenthesized text extract
R24_PAREN_TEXT = re.compile(r"\((.+?)\)")

# R25: Leading number strip
R25_LEADING_NUMBER = re.compile(r"^\d{1,2}[.)]?\s*")

# -- Phase C: New Patterns (R26-R41) --

# R26: Problem/question number (KICE exam style)
R26_PROBLEM_NUMBER = re.compile(
    r"^\s*(?:"
    r"(\d{1,2})\s*[..]"            # "1." or "1."
    r"|(\d{1,2})\s*번"              # "1번"
    r"|\[(\d{1,2})\]"              # "[1]"
    r"|\((\d{1,2})\)"              # "(1)"
    r")"
    r"\s*(.*)",
    re.DOTALL,
)

# R27-R30: Table context patterns (in-cell detection)
R27_CELL_LABEL = re.compile(r"^([가-힣]{2,6})\s*$")  # Short Korean label only
R28_CELL_VALUE = re.compile(r"^[^\uAC00-\uD7A3\s]+$")  # Non-Korean value only
R29_CELL_MIXED = re.compile(r"([가-힣]{2,6})\s*[:：]\s*(.+)")  # Label: value
R30_CELL_CHECKBOX = re.compile(r"^[□■☐☑]")  # Cell starts with checkbox

# R31-R33: Page number patterns (footer detection)
R31_PAGE_NUM_DASH = re.compile(r"^\s*-\s*\d+\s*-\s*$")  # "- 5 -"
R32_PAGE_NUM_PLAIN = re.compile(r"^\s*\d+\s*$")  # "5"
R33_PAGE_NUM_OF = re.compile(r"^\s*\d+\s*/\s*\d+\s*$")  # "5 / 10"

# R34-R36: Header/footer spatial markers
R34_HEADER_MARKER = re.compile(r"^(?:제목|머리말|Header)")  # Header keywords
R35_FOOTER_MARKER = re.compile(r"^(?:바닥글|Footer|페이지)")  # Footer keywords
R36_SHORT_LINE = re.compile(r"^.{1,15}$")  # Short line (header/footer candidate)

# R37-R40: Korean punctuation markers (kordoc P13)
R37_KR_COMMA = re.compile(r"[,，、]")  # Korean/CJK comma variants
R38_KR_PERIOD = re.compile(r"[.。．]")  # Korean/CJK period variants
R39_KR_SPACE_COMMA = re.compile(r"\s+,")  # Space before comma (error)
R40_KR_DOUBLE_SPACE = re.compile(r"  +")  # Multiple spaces

# R41: Merge-line heuristic (cross-script boundary)
R41_CROSS_SCRIPT = re.compile(
    r"([\uAC00-\uD7A3])\s*$"  # Korean char at end
    r"|\s*([\uAC00-\uD7A3])"  # or Korean char at start
)

# ---------------------------------------------------------------------------
# Label keywords for form field detection
# ---------------------------------------------------------------------------

LABEL_KEYWORDS: set[str] = {
    # Personal info
    "성명", "이름", "주소", "전화", "전화번호", "휴대폰", "연락처", "핸드폰",
    "생년월일", "주민등록번호", "소속", "직위", "직급", "부서",
    "이메일", "학교", "학년", "반", "번호", "학번", "학적", "학과",
    "캠퍼스", "대학", "단과대학",
    # Application-related
    "신청인", "대표자", "담당자", "작성자", "확인자", "승인자",
    "일시", "날짜", "기간", "장소", "목적", "사유", "비고",
    # Amount/quantity
    "금액", "수량", "단가", "합계", "계", "소계",
    # Form-specific
    "동아리명", "사업분야", "참가구분", "접수", "인원수", "아이템",
    "사업명", "기관명", "단체명", "프로젝트명",
    # Regulation-specific
    "비목", "항목해설", "증빙", "집행", "비용항목", "지출",
    "결제일", "결제금액", "카드번호", "승인번호", "사용처",
    "구분", "내용", "지도교수", "검수자", "검수일",
}

KR_CHAR_RE = re.compile(r"^[\uAC00-\uD7AF\u3131-\u318E]$")

# ---------------------------------------------------------------------------
# XML / ZIP Helpers
# ---------------------------------------------------------------------------


def local_tag(el: ET.Element) -> str:
    """Return the local name of an element, ignoring namespace."""
    tag = el.tag
    return tag.split("}")[-1] if "}" in tag else tag


def has_tag(parent: ET.Element, tag_name: str) -> bool:
    """Check if any descendant has the given local tag name."""
    return any(local_tag(child) == tag_name for child in parent.iter())


def collect_text(el: ET.Element) -> str:
    """Concatenate all <t> text nodes under an element."""
    parts: list[str] = []
    for child in el.iter():
        if local_tag(child) == "t" and child.text:
            parts.append(child.text)
    return "".join(parts)


def find_all_paragraphs(root: ET.Element) -> list[ET.Element]:
    """Return all paragraph <p> elements regardless of namespace."""
    return [
        el
        for el in root.iter()
        if el.tag.endswith("}p") and "paragraph" in el.tag
    ]


def _list_section_files(zf: zipfile.ZipFile) -> list[str]:
    """List section XML files inside a HWPX zip (Contents/section0.xml, etc.)."""
    sections: list[str] = []
    for name in sorted(zf.namelist()):
        if name.startswith("Contents/section") and name.endswith(".xml"):
            sections.append(name)
    return sections


def _parse_section(zf: zipfile.ZipFile, section_path: str) -> ET.Element:
    """Parse a section XML file from a HWPX zip into an ElementTree root.
    
    Args:
        zf: Open ZipFile object.
        section_path: Path to section XML within the ZIP.
        
    Returns:
        Parsed XML root element.
        
    Raises:
        ValueError: If section_path contains suspicious path components (ZipSlip defense).
    """
    # ZipSlip defense
    _validate_zip_path(section_path)
    
    with zf.open(section_path) as f:
        return ET.fromstring(f.read())


def _validate_zip_path(path: str) -> None:
    """Validate ZIP entry path for ZipSlip attacks.
    
    Args:
        path: ZIP entry filename to validate.
        
    Raises:
        ValueError: If path contains '..', absolute paths, or null bytes.
    """
    if ".." in path.split("/"):
        raise ValueError(f"ZipSlip attack detected: path contains '..': {path}")
    if os.path.isabs(path):
        raise ValueError(f"ZipSlip attack detected: absolute path: {path}")
    if "\x00" in path:
        raise ValueError(f"ZipSlip attack detected: null byte in path: {path}")


def normalize_uniform_spaces(text: str) -> str:
    """Normalize uniformly-distributed single-character Korean tokens.

    Korean form software often inserts spaces between every character
    for visual alignment (e.g. "학 번" for "학번"). This collapses those
    back when 70%+ of space-separated tokens are single Korean characters
    and total length <= 30.
    """
    if len(text) > 30 or " " not in text:
        return text
    tokens = text.split(" ")
    if len(tokens) < 2:
        return text
    kr_single = sum(1 for t in tokens if len(t) == 1 and KR_CHAR_RE.match(t))
    # For 2-token case: both must be single Korean chars (e.g. "학 번")
    # For 3+ tokens: 70% threshold applies (e.g. "소 속 대 학")
    if len(tokens) == 2:
        if kr_single == 2:
            return "".join(tokens)
    elif kr_single / len(tokens) >= 0.7:
        return "".join(tokens)
    return text


# F7: Phone spacing normalization
R_PHONE_SPACED = re.compile(
    r"(?<!\d)"
    r"(\d(?:[ ]?\d){1,3})"        # group 1: 2-4 digits with optional spaces
    r"\s*-\s*"
    r"(\d(?:[ ]?\d){2,3})"        # group 2: 3-4 digits with optional spaces
    r"\s*-\s*"
    r"(\d(?:[ ]?\d){2,3})"        # group 3: 3-4 digits (mandatory for phone)
    r"(?!\d)"
)


def normalize_phone_spacing(text: str) -> str:
    """Collapse uniform-spaced phone numbers.

    Examples:
        "45 0 -7 3 40" -> "450-7340"
        "0 1 0 -1 2 3 4 -5 6 7 8" -> "010-1234-5678"
    """
    def _collapse(m: re.Match) -> str:
        return "-".join(
            m.group(i).replace(" ", "") for i in (1, 2, 3)
        )

    return R_PHONE_SPACED.sub(_collapse, text)


# ---------------------------------------------------------------------------
# Core: extract_paragraphs
# ---------------------------------------------------------------------------


def extract_paragraphs(hwpx_path: str) -> list[str]:
    """Extract all paragraph texts from all HWPX sections.

    Opens the HWPX file as a ZIP, iterates over all section XML files,
    finds paragraph elements, and collects their text content.

    Args:
        hwpx_path: Path to the .hwpx file.

    Returns:
        List of paragraph text strings (may include empty strings for
        blank paragraphs).
    """
    texts: list[str] = []
    with zipfile.ZipFile(hwpx_path, "r") as zf:
        for section_path in _list_section_files(zf):
            root = _parse_section(zf, section_path)
            for p in find_all_paragraphs(root):
                texts.append(normalize_phone_spacing(collect_text(p)))
    return texts


# ---------------------------------------------------------------------------
# Core: classify_document
# ---------------------------------------------------------------------------


def classify_document(hwpx_path: str) -> tuple[str, dict[str, Any]]:
    """Classify an HWPX document into one of 5 types based on content analysis.

    Types:
        exam       - equations > 3 AND rect shapes > 5 (KICE-style exam papers)
        regulation - circle_bullets > 10 AND (appendix_refs > 0 OR
                     article_refs > 3) AND tables > 10
        form       - tables > 0 AND (checkboxes > 0 OR label_keywords > 3)
        report     - paragraphs > 50 AND tables < 3
        mixed      - default fallback

    Args:
        hwpx_path: Path to the .hwpx file.

    Returns:
        Tuple of (document_type, stats_dict).
    """
    stats: dict[str, int] = {
        "equations": 0,
        "tables": 0,
        "checkboxes": 0,
        "circle_bullets": 0,
        "rects": 0,
        "appendix_refs": 0,
        "article_refs": 0,
        "total_paragraphs": 0,
        "empty_paragraphs": 0,
        "label_keywords_found": 0,
    }

    with zipfile.ZipFile(hwpx_path, "r") as zf:
        for section_path in _list_section_files(zf):
            root = _parse_section(zf, section_path)
            _accumulate_stats(root, stats)

    # Classification logic
    if stats["equations"] > 3 and stats["rects"] > 5:
        return "exam", stats

    is_regulation = (
        stats["circle_bullets"] > 10
        and (stats["appendix_refs"] > 0 or stats["article_refs"] > 3)
        and stats["tables"] > 10
    )
    if is_regulation:
        return "regulation", stats

    if stats["tables"] > 0 and (
        stats["checkboxes"] > 0 or stats["label_keywords_found"] > 3
    ):
        return "form", stats

    non_empty = stats["total_paragraphs"] - stats["empty_paragraphs"]
    if non_empty > 50 and stats["tables"] < 3:
        return "report", stats

    return "mixed", stats


def _accumulate_stats(root: ET.Element, stats: dict[str, int]) -> None:
    """Walk an XML root and accumulate document statistics."""
    all_paragraphs = find_all_paragraphs(root)
    stats["total_paragraphs"] += len(all_paragraphs)

    for p in all_paragraphs:
        text = collect_text(p).strip()

        if not text:
            stats["empty_paragraphs"] += 1
            continue

        # Checkbox markers
        if re.search(r"[□■☑☐]", text):
            stats["checkboxes"] += 1

        # Circle bullets
        if text.startswith("○"):
            stats["circle_bullets"] += 1

        # Appendix references
        if R8_APPENDIX_REF.search(text):
            stats["appendix_refs"] += 1

        # Article references (제N조/항/호)
        if re.search(r"제\d+[조호항]", text):
            stats["article_refs"] += 1

        # Label keyword detection
        normalized = normalize_uniform_spaces(text)
        for kw in LABEL_KEYWORDS:
            if kw in normalized:
                stats["label_keywords_found"] += 1
                break  # count at most once per paragraph

    # Count structural elements across the entire tree
    for el in root.iter():
        tag = local_tag(el)
        if tag == "equation" or tag == "script":
            # Count only substantive equations (script text > 3 chars)
            if tag == "script":
                if el.text and len(el.text.strip()) > 3:
                    stats["equations"] += 1
            else:
                stats["equations"] += 1
        elif tag == "tbl":
            stats["tables"] += 1
        elif tag == "rect":
            stats["rects"] += 1


# ---------------------------------------------------------------------------
# Core: extract_checkbox_hierarchy
# ---------------------------------------------------------------------------

DEPTH_MAP: dict[str, int] = {
    "□": 0,
    "○": 1,
    "●": 1,
    "◎": 1,
    "-": 2,
    "*": 3,
}


def extract_checkbox_hierarchy(
    paragraphs: list[str],
) -> list[dict[str, Any]]:
    """Extract 4-level checkbox hierarchy from paragraph texts.

    Hierarchy levels:
        □ = heading (depth 0)
        ○/●/◎ = item (depth 1)
        - = detail (depth 2)
        * = note (depth 3)

    Args:
        paragraphs: List of paragraph text strings.

    Returns:
        List of dicts with keys: depth, marker, text, paragraph_index,
        children (always an empty list; caller may build tree from depth).
    """
    items: list[dict[str, Any]] = []

    for i, text in enumerate(paragraphs):
        stripped = text.strip()
        m = R15_CHECKBOX_HIERARCHY.match(stripped)
        if m:
            marker = m.group(1)
            content = m.group(2).strip()
            items.append(
                {
                    "depth": DEPTH_MAP.get(marker, 0),
                    "marker": marker,
                    "text": content,
                    "paragraph_index": i,
                    "children": [],
                }
            )

    return items


# ---------------------------------------------------------------------------
# Core: extract_appendix_refs
# ---------------------------------------------------------------------------


def extract_appendix_refs(
    paragraphs: list[str],
) -> list[dict[str, Any]]:
    """Extract appendix references ([별첨 제N호], [별지], [별표]) from paragraphs.

    Args:
        paragraphs: List of paragraph text strings.

    Returns:
        List of dicts with keys: ref, number (int or None), title, paragraph_index.
    """
    refs: list[dict[str, Any]] = []

    for i, text in enumerate(paragraphs):
        stripped = text.strip()
        m = R8_APPENDIX_REF.search(stripped)
        if m:
            number = int(m.group(1)) if m.group(1) else None
            title = stripped[m.end() :].strip() if m.end() < len(stripped) else ""
            refs.append(
                {
                    "ref": m.group(0),
                    "number": number,
                    "title": title[:60],
                    "paragraph_index": i,
                }
            )

    return refs


# ---------------------------------------------------------------------------
# Core: detect_digit_headings
# ---------------------------------------------------------------------------


def detect_digit_headings(
    paragraphs: list[str],
) -> list[dict[str, Any]]:
    """Detect digit-concatenated headings (e.g. '3지원금 집행기준').

    These are non-standard section numbering patterns found in Korean
    regulations where a digit is directly concatenated to the title
    without any space or punctuation.

    Args:
        paragraphs: List of paragraph text strings.

    Returns:
        List of dicts with keys: number, title, paragraph_index.
    """
    headings: list[dict[str, Any]] = []

    for i, text in enumerate(paragraphs):
        stripped = text.strip()
        m = R9_DIGIT_HEADING.match(stripped)
        if m:
            num = int(m.group(1))
            title = stripped[len(m.group(1)) :].strip()
            if len(title) >= 3:  # require at least 3 chars in title
                headings.append(
                    {
                        "number": num,
                        "title": title,
                        "paragraph_index": i,
                    }
                )

    return headings


# ---------------------------------------------------------------------------
# Phase C: New Functions (C1-C9)
# ---------------------------------------------------------------------------


def parse_field_string(fields_str: str) -> dict[str, str]:
    """Parse a comma-separated key=value field string with comma protection (C9).

    Splits on commas only when followed by a Korean/English key and '='.
    This allows values to contain commas safely.

    Examples:
        "성명=홍길동,주소=서울시 강남구" -> {"성명": "홍길동", "주소": "서울시 강남구"}
        "목적=연구, 개발,기간=1년" -> {"목적": "연구, 개발", "기간": "1년"}

    Args:
        fields_str: Comma-separated "key=value" string.

    Returns:
        Dict mapping field labels to values.

    Raises:
        ValueError: If a pair lacks '=' separator.
    """
    # Split only on comma followed by a key pattern (Korean or Latin + '=')
    pairs = re.split(r",(?=[가-힣A-Za-z][가-힣A-Za-z\s]*=)", fields_str)

    result: dict[str, str] = {}
    for pair in pairs:
        pair = pair.strip()
        if not pair:
            continue
        if "=" not in pair:
            raise ValueError(f"Invalid field pair (missing '='): {pair!r}")
        key, value = pair.split("=", 1)
        result[key.strip()] = value.strip()

    return result


def find_page_boundaries(hwpx_path: str) -> list[dict[str, Any]]:
    """Detect page/column boundaries from paragraph attributes (C1).

    HWPX paragraphs can carry pageBreak="1" or columnBreak="1" attributes
    on their <p> element or nested <paraShape>/<lineseg> nodes. This scans
    all section files and returns an ordered list of boundaries.

    Args:
        hwpx_path: Path to the .hwpx file.

    Returns:
        List of dicts with keys:
          - paragraph_index (int): global paragraph index
          - break_type (str): "page" | "column" | "section"
          - section (str): section filename (e.g. "Contents/section0.xml")
    """
    boundaries: list[dict[str, Any]] = []
    global_idx = 0

    with zipfile.ZipFile(hwpx_path, "r") as zf:
        section_files = _list_section_files(zf)

        for sec_idx, section_path in enumerate(section_files):
            # Each new section file is implicitly a section break
            if sec_idx > 0:
                boundaries.append({
                    "paragraph_index": global_idx,
                    "break_type": "section",
                    "section": section_path,
                })

            root = _parse_section(zf, section_path)
            paragraphs = find_all_paragraphs(root)

            for p in paragraphs:
                # Check direct attributes on <p>
                page_break = p.get("pageBreak", "0")
                col_break = p.get("columnBreak", "0")

                # Also check namespace-prefixed attributes
                for attr_name in list(p.attrib.keys()):
                    local = attr_name.split("}")[-1] if "}" in attr_name else attr_name
                    if local == "pageBreak" and p.get(attr_name) == "1":
                        page_break = "1"
                    elif local == "columnBreak" and p.get(attr_name) == "1":
                        col_break = "1"

                # Also scan child elements for break attributes
                for child in p.iter():
                    ltag = local_tag(child)
                    if ltag in ("paraShape", "paraPr", "lineseg"):
                        if child.get("pageBreak") == "1":
                            page_break = "1"
                        if child.get("columnBreak") == "1":
                            col_break = "1"

                if page_break == "1":
                    boundaries.append({
                        "paragraph_index": global_idx,
                        "break_type": "page",
                        "section": section_path,
                    })
                elif col_break == "1":
                    boundaries.append({
                        "paragraph_index": global_idx,
                        "break_type": "column",
                        "section": section_path,
                    })

                global_idx += 1

    return boundaries


def find_problem_starts(hwpx_path: str) -> list[dict[str, Any]]:
    """Map problem/question numbers to paragraph indices (exam documents, C2).

    Scans all paragraphs for patterns like "1.", "1번", "[1]", "(1)" at the
    start of text. Returns ordered list of problem starts with their paragraph
    positions.

    Args:
        hwpx_path: Path to the .hwpx file.

    Returns:
        List of dicts with keys:
          - problem_number (int): detected problem number
          - paragraph_index (int): global paragraph index
          - preview (str): first 60 chars of problem text
          - pattern (str): which pattern matched ("dot", "번", "bracket", "paren")
    """
    problems: list[dict[str, Any]] = []
    paragraphs = extract_paragraphs(hwpx_path)

    for i, text in enumerate(paragraphs):
        stripped = text.strip()
        if not stripped:
            continue

        m = R26_PROBLEM_NUMBER.match(stripped)
        if not m:
            continue

        # Determine which group matched
        if m.group(1) is not None:
            num, pattern = int(m.group(1)), "dot"
        elif m.group(2) is not None:
            num, pattern = int(m.group(2)), "번"
        elif m.group(3) is not None:
            num, pattern = int(m.group(3)), "bracket"
        elif m.group(4) is not None:
            num, pattern = int(m.group(4)), "paren"
        else:
            continue

        # Filter: problem numbers should be reasonable (1-50)
        if num < 1 or num > 50:
            continue

        # Extract preview text
        rest = m.group(5) or ""
        preview = rest[:60].strip() if rest else "(no text)"

        problems.append({
            "problem_number": num,
            "paragraph_index": i,
            "preview": preview,
            "pattern": pattern,
        })

    return problems


def detect_incell_patterns(hwpx_path: str) -> dict[str, Any]:
    """Detect in-cell form patterns (labels, values, checkboxes) in tables (C3).

    Analyzes table cell contents to identify common patterns:
    - Short Korean labels (e.g., "성명", "주소")
    - Value-only cells (numbers, dates)
    - Label: value cells (e.g., "성명: 홍길동")
    - Checkbox cells

    Args:
        hwpx_path: Path to the .hwpx file.

    Returns:
        Dict with keys:
          - label_cells (int): count of label-only cells
          - value_cells (int): count of value-only cells
          - mixed_cells (int): count of label:value cells
          - checkbox_cells (int): count of checkbox cells
          - total_cells (int): total cells analyzed
    """
    stats = {
        "label_cells": 0,
        "value_cells": 0,
        "mixed_cells": 0,
        "checkbox_cells": 0,
        "total_cells": 0,
    }

    with zipfile.ZipFile(hwpx_path, "r") as zf:
        section_files = _list_section_files(zf)

        for section_path in section_files:
            root = _parse_section(zf, section_path)

            # Find all table cells (tbl > tc or similar)
            for el in root.iter():
                ltag = local_tag(el)
                if ltag == "tc":  # table cell
                    stats["total_cells"] += 1
                    cell_text = collect_text(el).strip()

                    if not cell_text:
                        continue

                    # Check patterns in order of specificity
                    if R30_CELL_CHECKBOX.match(cell_text):
                        stats["checkbox_cells"] += 1
                    elif R29_CELL_MIXED.match(cell_text):
                        stats["mixed_cells"] += 1
                    elif R27_CELL_LABEL.match(cell_text):
                        stats["label_cells"] += 1
                    elif R28_CELL_VALUE.match(cell_text):
                        stats["value_cells"] += 1

    return stats


def strip_page_numbers(
    hwpx_path: str,
    output_path: str,
    patterns: list[str] | None = None,
) -> dict[str, Any]:
    """Remove page number paragraphs from HWPX (C4).

    Detects and removes paragraphs that contain only page numbers in common
    formats: "- 5 -", "5", "5 / 10".

    Args:
        hwpx_path: Path to input .hwpx file.
        output_path: Path for output .hwpx file.
        patterns: List of pattern names to use ("dash", "plain", "of").
                  If None, uses all patterns.

    Returns:
        Dict with keys:
          - removed_count (int): number of paragraphs removed
          - patterns_used (list[str]): patterns that matched
    """
    if patterns is None:
        patterns = ["dash", "plain", "of"]

    pattern_map = {
        "dash": R31_PAGE_NUM_DASH,
        "plain": R32_PAGE_NUM_PLAIN,
        "of": R33_PAGE_NUM_OF,
    }

    active_patterns = [pattern_map[p] for p in patterns if p in pattern_map]
    removed_count = 0
    matched_patterns: set[str] = set()

    tmp_fd, tmp_path = tempfile.mkstemp(suffix=".hwpx")
    os.close(tmp_fd)

    try:
        with zipfile.ZipFile(hwpx_path, "r") as zf_in:
            with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zf_out:
                for item in zf_in.infolist():
                    _validate_zip_path(item.filename)
                    data = zf_in.read(item.filename)

                    is_section = (
                        item.filename.endswith(".xml")
                        and re.search(r"[Ss]ection\d+\.xml$", item.filename)
                    )

                    if is_section:
                        root = ET.fromstring(data.decode("utf-8"))
                        paragraphs = find_all_paragraphs(root)

                        # Mark paragraphs for removal
                        to_remove: list[ET.Element] = []
                        for p in paragraphs:
                            p_text = collect_text(p).strip()
                            for pat_name, regex in zip(patterns, active_patterns):
                                if regex.match(p_text):
                                    to_remove.append(p)
                                    matched_patterns.add(pat_name)
                                    break

                        # Remove marked paragraphs
                        for p in to_remove:
                            parent = None
                            for candidate in root.iter():
                                if p in list(candidate):
                                    parent = candidate
                                    break
                            if parent is not None:
                                parent.remove(p)
                                removed_count += 1

                        # Serialize back
                        xml_text = ET.tostring(root, encoding="unicode", xml_declaration=False)
                        if data.decode("utf-8").startswith("<?xml"):
                            xml_text = '<?xml version="1.0" encoding="UTF-8"?>\n' + xml_text
                        data = xml_text.encode("utf-8")

                    # Write entry
                    if item.filename == "mimetype":
                        zf_out.writestr(item, data, compress_type=zipfile.ZIP_STORED)
                    else:
                        zf_out.writestr(item, data)

        shutil.move(tmp_path, output_path)
    except Exception:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)
        raise

    return {
        "removed_count": removed_count,
        "patterns_used": sorted(matched_patterns),
    }


def detect_headers_footers(hwpx_path: str) -> dict[str, Any]:
    """Detect header/footer paragraphs based on spatial markers (C5).

    Identifies paragraphs that likely represent headers or footers based on:
    - Short length (<= 15 chars)
    - Header/footer keywords
    - Position in section (first/last paragraphs)

    Args:
        hwpx_path: Path to the .hwpx file.

    Returns:
        Dict with keys:
          - headers (list[dict]): detected headers with paragraph_index, text
          - footers (list[dict]): detected footers with paragraph_index, text
          - total_candidates (int): total header/footer candidates found
    """
    headers: list[dict[str, Any]] = []
    footers: list[dict[str, Any]] = []
    global_idx = 0

    with zipfile.ZipFile(hwpx_path, "r") as zf:
        section_files = _list_section_files(zf)

        for section_path in section_files:
            root = _parse_section(zf, section_path)
            paragraphs = find_all_paragraphs(root)

            for local_idx, p in enumerate(paragraphs):
                p_text = collect_text(p).strip()

                # Short line heuristic
                if not R36_SHORT_LINE.match(p_text):
                    global_idx += 1
                    continue

                # Check for header markers
                if R34_HEADER_MARKER.search(p_text) or local_idx < 2:
                    headers.append({
                        "paragraph_index": global_idx,
                        "text": p_text,
                        "section": section_path,
                    })
                # Check for footer markers
                elif R35_FOOTER_MARKER.search(p_text) or local_idx >= len(paragraphs) - 2:
                    footers.append({
                        "paragraph_index": global_idx,
                        "text": p_text,
                        "section": section_path,
                    })

                global_idx += 1

    return {
        "headers": headers,
        "footers": footers,
        "total_candidates": len(headers) + len(footers),
    }


def detect_korean_markers(paragraphs: list[str]) -> dict[str, Any]:
    """Detect Korean punctuation markers and spacing errors (C6).

    Analyzes text for:
    - Korean comma variants (,，、)
    - Korean period variants (.。．)
    - Space before comma errors
    - Multiple consecutive spaces

    Args:
        paragraphs: List of paragraph text strings.

    Returns:
        Dict with keys:
          - kr_commas (int): paragraphs with Korean commas
          - kr_periods (int): paragraphs with Korean periods
          - space_comma_errors (int): paragraphs with space before comma
          - double_space_errors (int): paragraphs with multiple spaces
    """
    stats = {
        "kr_commas": 0,
        "kr_periods": 0,
        "space_comma_errors": 0,
        "double_space_errors": 0,
    }

    for text in paragraphs:
        if R37_KR_COMMA.search(text):
            stats["kr_commas"] += 1
        if R38_KR_PERIOD.search(text):
            stats["kr_periods"] += 1
        if R39_KR_SPACE_COMMA.search(text):
            stats["space_comma_errors"] += 1
        if R40_KR_DOUBLE_SPACE.search(text):
            stats["double_space_errors"] += 1

    return stats


def _is_marker_line(text: str) -> bool:
    """Helper: Check if a line is likely a structural marker (C7 helper).

    Args:
        text: Paragraph text.

    Returns:
        True if the line appears to be a heading, label, or marker.
    """
    stripped = text.strip()
    if not stripped:
        return False

    # Check for structural patterns
    if R1_CHAPTER_HEADING.match(stripped):
        return True
    if R2_ARTICLE.match(stripped):
        return True
    if R3_CIRCLED_NUMBER.match(stripped):
        return True
    if R4_NUMBERED_LIST.match(stripped):
        return True
    if R5_KOREAN_LETTER.match(stripped):
        return True
    if R6_CHECKBOX_FLAT.match(stripped):
        return True

    # Short all-caps or all-Korean lines
    if len(stripped) <= 20 and (stripped.isupper() or R20_SHORT_KOREAN_LABEL.match(stripped)):
        return True

    return False


def should_merge_lines(line1: str, line2: str) -> bool:
    """Determine if two lines should be merged based on cross-script boundaries (C7).

    Korean text often gets split mid-sentence when mixed with English/numbers.
    This detects cases where:
    - Line 1 ends with Korean but no sentence-ending punctuation
    - Line 2 starts with Korean or continues the sentence
    - Neither line is a structural marker (heading, list item, etc.)

    Args:
        line1: First line text.
        line2: Second line text.

    Returns:
        True if lines should be merged.
    """
    if not line1 or not line2:
        return False

    # Don't merge if either is a marker line
    if _is_marker_line(line1) or _is_marker_line(line2):
        return False

    # Check for sentence-ending punctuation
    if line1.rstrip().endswith((".", "。", "!", "?", ":", "：")):
        return False

    # Cross-script merge heuristic
    # If line1 ends with Korean char and no punctuation, likely continuation
    if R41_CROSS_SCRIPT.search(line1):
        return True

    # If line2 starts with lowercase or number, likely continuation
    if line2 and line2[0].islower():
        return True

    return False


def merge_lines(paragraphs: list[str]) -> list[str]:
    """Merge paragraphs that were incorrectly split mid-sentence (C7).

    Uses cross-script boundary heuristics to identify and merge split lines.

    Args:
        paragraphs: List of paragraph text strings.

    Returns:
        New list of paragraphs with splits merged.
    """
    if not paragraphs:
        return []

    merged: list[str] = [paragraphs[0]]

    for current in paragraphs[1:]:
        if merged and should_merge_lines(merged[-1], current):
            merged[-1] = merged[-1].rstrip() + " " + current.lstrip()
        else:
            merged.append(current)

    return merged


def fill_hwpx_preserve(
    hwpx_path: str,
    output_path: str,
    fields: dict[str, str],
) -> dict[str, Any]:
    """Fill form fields in HWPX by direct XML surgery, preserving styles (C8).

    Strategy:
    1. Open HWPX as ZIP
    2. Find section XML files matching /[Ss]ection\\d+\\.xml$/
    3. For each field, find the paragraph containing the label
    4. Replace text in the first <t> element of the first <run>
       (preserving charPrIDRef for style continuity)
    5. Clear remaining runs in that paragraph
    6. Strip all linesegarray elements (force recalc)
    7. Rewrite ZIP preserving original structure

    Args:
        hwpx_path: Path to input .hwpx file.
        output_path: Path for output .hwpx file.
        fields: Dict of {label: value} to fill.

    Returns:
        Dict with keys:
          - filled (list[str]): labels that were successfully filled
          - not_found (list[str]): labels not found in document
          - lineseg_stripped (int): count of lineseg elements removed
    """
    filled: list[str] = []
    not_found: list[str] = list(fields.keys())
    total_lineseg = 0

    tmp_fd, tmp_path = tempfile.mkstemp(suffix=".hwpx")
    os.close(tmp_fd)

    try:
        with zipfile.ZipFile(hwpx_path, "r") as zf_in:
            with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zf_out:
                for item in zf_in.infolist():
                    _validate_zip_path(item.filename)
                    data = zf_in.read(item.filename)

                    is_section = (
                        item.filename.endswith(".xml")
                        and re.search(r"[Ss]ection\d+\.xml$", item.filename)
                    )

                    if is_section:
                        xml_text = data.decode("utf-8")

                        # Parse and process
                        root = ET.fromstring(xml_text)
                        paragraphs = find_all_paragraphs(root)

                        for p in paragraphs:
                            p_text = collect_text(p).strip()
                            normalized_p = normalize_uniform_spaces(p_text)

                            for label, value in list(fields.items()):
                                if label in not_found and label in normalized_p:
                                    t_elements = [
                                        el for el in p.iter()
                                        if local_tag(el) == "t"
                                    ]

                                    if t_elements:
                                        # Strategy: find "label: ___" pattern and replace the value part
                                        # NOT the label itself
                                        import re as _re
                                        colon_pat = _re.compile(
                                            _re.escape(label) + r"\s*[:：]\s*(.*)",
                                            _re.DOTALL,
                                        )

                                        replaced = False
                                        for t_el in t_elements:
                                            if t_el.text is None:
                                                continue
                                            m = colon_pat.search(normalize_uniform_spaces(t_el.text))
                                            if m:
                                                # Replace only the value portion after label:
                                                t_el.text = colon_pat.sub(
                                                    label + ": " + value,
                                                    t_el.text,
                                                    count=1,
                                                )
                                                replaced = True
                                                break

                                        if not replaced:
                                            # Fallback: look for adjacent <t> after the label <t>
                                            label_idx = -1
                                            for i, t_el in enumerate(t_elements):
                                                if t_el.text and label in normalize_uniform_spaces(t_el.text):
                                                    label_idx = i
                                                    break

                                            if label_idx >= 0 and label_idx + 1 < len(t_elements):
                                                # Set the NEXT <t> element (value cell)
                                                t_elements[label_idx + 1].text = value
                                                replaced = True

                                        if replaced:
                                            filled.append(label)
                                            not_found.remove(label)

                        # Serialize back
                        xml_text = ET.tostring(
                            root, encoding="unicode", xml_declaration=False
                        )

                        # Add XML declaration if it was present
                        if data.decode("utf-8").startswith("<?xml"):
                            xml_text = '<?xml version="1.0" encoding="UTF-8"?>\n' + xml_text

                        # Strip lineseg
                        count_open = len(_LINESEG_OPEN.findall(xml_text))
                        count_self = len(_LINESEG_SELF.findall(xml_text))
                        xml_text = _LINESEG_OPEN.sub("", xml_text)
                        xml_text = _LINESEG_SELF.sub("", xml_text)
                        total_lineseg += count_open + count_self

                        data = xml_text.encode("utf-8")

                    # Write entry
                    if item.filename == "mimetype":
                        zf_out.writestr(item, data, compress_type=zipfile.ZIP_STORED)
                    else:
                        zf_out.writestr(item, data)

        shutil.move(tmp_path, output_path)
    except Exception:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)
        raise

    return {
        "filled": filled,
        "not_found": not_found,
        "lineseg_stripped": total_lineseg,
    }


# ---------------------------------------------------------------------------
# Core: strip_lineseg
# ---------------------------------------------------------------------------

_LINESEG_OPEN = re.compile(
    r"<(?:\w+:)?linesegarray[^>]*>.*?</(?:\w+:)?linesegarray>",
    re.DOTALL,
)
_LINESEG_SELF = re.compile(r"<(?:\w+:)?linesegarray[^/]*/>")


def strip_lineseg(hwpx_path: str, output_path: str) -> int:
    """Strip all linesegarray elements from an HWPX file.

    Linesegarray elements store line-break position caches. Removing
    them forces Hancom Office to recalculate line breaks on open, which
    is required after content edits to avoid layout corruption.

    Args:
        hwpx_path: Path to the input .hwpx file.
        output_path: Path for the output .hwpx file.

    Returns:
        Total count of linesegarray elements stripped.
    """
    total_stripped = 0

    # Work in a temporary file for atomic write
    tmp_fd, tmp_path = tempfile.mkstemp(suffix=".hwpx")
    os.close(tmp_fd)

    try:
        with zipfile.ZipFile(hwpx_path, "r") as zf_in:
            with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zf_out:
                for item in zf_in.infolist():
                    data = zf_in.read(item.filename)

                    if item.filename.endswith(".xml") and item.filename.startswith(
                        "Contents/section"
                    ):
                        xml_text = data.decode("utf-8")

                        # Count before stripping
                        count_open = len(_LINESEG_OPEN.findall(xml_text))
                        count_self = len(_LINESEG_SELF.findall(xml_text))

                        # Strip
                        xml_text = _LINESEG_OPEN.sub("", xml_text)
                        xml_text = _LINESEG_SELF.sub("", xml_text)

                        total_stripped += count_open + count_self
                        data = xml_text.encode("utf-8")

                    # Preserve mimetype as STORED (first entry convention)
                    if item.filename == "mimetype":
                        zf_out.writestr(item, data, compress_type=zipfile.ZIP_STORED)
                    else:
                        zf_out.writestr(item, data)

        # Atomic move to final destination
        shutil.move(tmp_path, output_path)
    except Exception:
        # Clean up temp file on failure
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)
        raise

    return total_stripped


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def _cmd_classify(args: argparse.Namespace) -> None:
    """Handle the 'classify' subcommand."""
    doc_type, stats = classify_document(args.hwpx_path)
    print(f"Document type: {doc_type}")
    print(f"Statistics:")
    for key, value in sorted(stats.items()):
        print(f"  {key}: {value}")


def _cmd_hierarchy(args: argparse.Namespace) -> None:
    """Handle the 'hierarchy' subcommand."""
    paragraphs = extract_paragraphs(args.hwpx_path)
    items = extract_checkbox_hierarchy(paragraphs)

    if not items:
        print("No checkbox hierarchy found.")
        return

    print(f"Checkbox hierarchy ({len(items)} items):")
    for item in items:
        indent = "  " * item["depth"]
        print(
            f"  {indent}{item['marker']} [{item['depth']}] "
            f"(p{item['paragraph_index']}): {item['text'][:70]}"
        )


def _cmd_appendix(args: argparse.Namespace) -> None:
    """Handle the 'appendix' subcommand."""
    paragraphs = extract_paragraphs(args.hwpx_path)
    refs = extract_appendix_refs(paragraphs)

    if not refs:
        print("No appendix references found.")
        return

    print(f"Appendix references ({len(refs)} found):")
    for ref in refs:
        num_str = f"#{ref['number']}" if ref["number"] is not None else "(unnumbered)"
        print(
            f"  {ref['ref']} {num_str} "
            f"(p{ref['paragraph_index']}): {ref['title']}"
        )


def _cmd_strip_lineseg(args: argparse.Namespace) -> None:
    """Handle the 'strip-lineseg' subcommand."""
    count = strip_lineseg(args.hwpx_path, args.output_path)
    print(f"Stripped {count} linesegarray element(s).")
    print(f"Output: {args.output_path}")


def _cmd_extract(args: argparse.Namespace) -> None:
    """Handle the 'extract' subcommand."""
    paragraphs = extract_paragraphs(args.hwpx_path)

    non_empty = [p for p in paragraphs if p.strip()]
    print(f"Total paragraphs: {len(paragraphs)} ({len(non_empty)} non-empty)")
    print("---")
    for i, text in enumerate(paragraphs):
        stripped = text.strip()
        if stripped:
            print(f"[{i:04d}] {stripped}")


def _cmd_digit_headings(args: argparse.Namespace) -> None:
    """Handle the 'digit-headings' subcommand."""
    paragraphs = extract_paragraphs(args.hwpx_path)
    headings = detect_digit_headings(paragraphs)

    if not headings:
        print("No digit-concatenated headings found.")
        return

    print(f"Digit headings ({len(headings)} found):")
    for h in headings:
        print(f"  {h['number']}. {h['title']} (p{h['paragraph_index']})")


def _cmd_pages(args: argparse.Namespace) -> None:
    """Handle the 'pages' subcommand (C1)."""
    boundaries = find_page_boundaries(args.hwpx_path)

    if not boundaries:
        print("No page/column/section boundaries found.")
        return

    print(f"Boundaries ({len(boundaries)} found):")
    for b in boundaries:
        print(
            f"  [{b['break_type']:7s}] paragraph {b['paragraph_index']:4d} "
            f"({b['section']})"
        )


def _cmd_problems(args: argparse.Namespace) -> None:
    """Handle the 'problems' subcommand (C2)."""
    problems = find_problem_starts(args.hwpx_path)

    if not problems:
        print("No problem numbers found.")
        return

    print(f"Problems ({len(problems)} found):")
    for prob in problems:
        print(
            f"  Problem {prob['problem_number']:2d} [{prob['pattern']:7s}] "
            f"p{prob['paragraph_index']:4d}: {prob['preview']}"
        )


def _cmd_incell(args: argparse.Namespace) -> None:
    """Handle the 'incell' subcommand (C3)."""
    stats = detect_incell_patterns(args.hwpx_path)

    print(f"In-cell pattern analysis:")
    print(f"  Total cells:     {stats['total_cells']}")
    print(f"  Label cells:     {stats['label_cells']}")
    print(f"  Value cells:     {stats['value_cells']}")
    print(f"  Mixed cells:     {stats['mixed_cells']}")
    print(f"  Checkbox cells:  {stats['checkbox_cells']}")


def _cmd_markers(args: argparse.Namespace) -> None:
    """Handle the 'markers' subcommand (C6)."""
    paragraphs = extract_paragraphs(args.hwpx_path)
    stats = detect_korean_markers(paragraphs)

    print(f"Korean marker analysis:")
    print(f"  Paragraphs with Korean commas:     {stats['kr_commas']}")
    print(f"  Paragraphs with Korean periods:    {stats['kr_periods']}")
    print(f"  Space before comma errors:         {stats['space_comma_errors']}")
    print(f"  Multiple space errors:             {stats['double_space_errors']}")


def _cmd_headers_footers(args: argparse.Namespace) -> None:
    """Handle the 'headers-footers' subcommand (C5)."""
    result = detect_headers_footers(args.hwpx_path)

    print(f"Header/footer detection ({result['total_candidates']} found):")
    
    if result["headers"]:
        print(f"\nHeaders ({len(result['headers'])}):")
        for h in result["headers"]:
            print(f"  p{h['paragraph_index']:4d}: {h['text']}")
    
    if result["footers"]:
        print(f"\nFooters ({len(result['footers'])}):")
        for f in result["footers"]:
            print(f"  p{f['paragraph_index']:4d}: {f['text']}")


def _cmd_fill(args: argparse.Namespace) -> None:
    """Handle the 'fill' subcommand (C8)."""
    fields = parse_field_string(args.fields)
    result = fill_hwpx_preserve(args.hwpx_path, args.output_path, fields)

    print(f"Filled {len(result['filled'])} field(s):")
    for label in result["filled"]:
        print(f"  + {label}")

    if result["not_found"]:
        print(f"Not found ({len(result['not_found'])}):")
        for label in result["not_found"]:
            print(f"  - {label}")

    print(f"Lineseg stripped: {result['lineseg_stripped']}")
    print(f"Output: {args.output_path}")


def main() -> None:
    """Entry point with argparse-based CLI."""
    parser = argparse.ArgumentParser(
        description="HWPX Korean Document Pattern Matching and Editing",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  %(prog)s classify doc.hwpx\n"
            "  %(prog)s hierarchy doc.hwpx\n"
            "  %(prog)s appendix doc.hwpx\n"
            "  %(prog)s strip-lineseg doc.hwpx output.hwpx\n"
            "  %(prog)s extract doc.hwpx\n"
            "  %(prog)s digit-headings doc.hwpx\n"
            "  %(prog)s pages doc.hwpx\n"
            "  %(prog)s problems doc.hwpx\n"
            "  %(prog)s incell doc.hwpx\n"
            "  %(prog)s markers doc.hwpx\n"
            "  %(prog)s headers-footers doc.hwpx\n"
            "  %(prog)s fill doc.hwpx output.hwpx '성명=홍길동,주소=서울'\n"
        ),
    )

    subparsers = parser.add_subparsers(dest="command", required=True)

    # classify
    p_classify = subparsers.add_parser(
        "classify", help="Classify document type (exam/regulation/form/report/mixed)"
    )
    p_classify.add_argument("hwpx_path", help="Path to .hwpx file")
    p_classify.set_defaults(func=_cmd_classify)

    # hierarchy
    p_hierarchy = subparsers.add_parser(
        "hierarchy", help="Extract checkbox hierarchy (4-level depth)"
    )
    p_hierarchy.add_argument("hwpx_path", help="Path to .hwpx file")
    p_hierarchy.set_defaults(func=_cmd_hierarchy)

    # appendix
    p_appendix = subparsers.add_parser(
        "appendix", help="Extract appendix references"
    )
    p_appendix.add_argument("hwpx_path", help="Path to .hwpx file")
    p_appendix.set_defaults(func=_cmd_appendix)

    # strip-lineseg
    p_strip = subparsers.add_parser(
        "strip-lineseg", help="Strip linesegarray elements from HWPX"
    )
    p_strip.add_argument("hwpx_path", help="Path to input .hwpx file")
    p_strip.add_argument("output_path", help="Path for output .hwpx file")
    p_strip.set_defaults(func=_cmd_strip_lineseg)

    # extract
    p_extract = subparsers.add_parser(
        "extract", help="Extract all paragraph texts"
    )
    p_extract.add_argument("hwpx_path", help="Path to .hwpx file")
    p_extract.set_defaults(func=_cmd_extract)

    # digit-headings (bonus command for detect_digit_headings)
    p_digit = subparsers.add_parser(
        "digit-headings", help="Detect digit-concatenated headings"
    )
    p_digit.add_argument("hwpx_path", help="Path to .hwpx file")
    p_digit.set_defaults(func=_cmd_digit_headings)

    # pages (C1)
    p_pages = subparsers.add_parser(
        "pages", help="Detect page/column/section boundaries"
    )
    p_pages.add_argument("hwpx_path", help="Path to .hwpx file")
    p_pages.set_defaults(func=_cmd_pages)

    # problems (C2)
    p_problems = subparsers.add_parser(
        "problems", help="Map problem/question numbers to paragraphs"
    )
    p_problems.add_argument("hwpx_path", help="Path to .hwpx file")
    p_problems.set_defaults(func=_cmd_problems)

    # incell (C3)
    p_incell = subparsers.add_parser(
        "incell", help="Detect in-cell form patterns (tables)"
    )
    p_incell.add_argument("hwpx_path", help="Path to .hwpx file")
    p_incell.set_defaults(func=_cmd_incell)

    # markers (C6)
    p_markers = subparsers.add_parser(
        "markers", help="Detect Korean punctuation markers and spacing errors"
    )
    p_markers.add_argument("hwpx_path", help="Path to .hwpx file")
    p_markers.set_defaults(func=_cmd_markers)

    # headers-footers (C5)
    p_hf = subparsers.add_parser(
        "headers-footers", help="Detect header/footer paragraphs"
    )
    p_hf.add_argument("hwpx_path", help="Path to .hwpx file")
    p_hf.set_defaults(func=_cmd_headers_footers)

    # fill (C8)
    p_fill = subparsers.add_parser(
        "fill", help="Fill form fields via XML surgery (style-preserving)"
    )
    p_fill.add_argument("hwpx_path", help="Path to input .hwpx file")
    p_fill.add_argument("output_path", help="Path for output .hwpx file")
    p_fill.add_argument(
        "fields",
        help="Comma-separated key=value pairs (e.g. '성명=홍길동,주소=서울')"
    )
    p_fill.set_defaults(func=_cmd_fill)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
