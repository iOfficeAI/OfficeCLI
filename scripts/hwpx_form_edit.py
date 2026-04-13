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
"""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from typing import Any

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
    """Parse a section XML file from a HWPX zip into an ElementTree root."""
    with zf.open(section_path) as f:
        return ET.fromstring(f.read())


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
                texts.append(collect_text(p))
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
# Core: strip_lineseg
# ---------------------------------------------------------------------------

_LINESEG_OPEN = re.compile(
    r"<(?:hp:)?linesegarray[^>]*>.*?</(?:hp:)?linesegarray>",
    re.DOTALL,
)
_LINESEG_SELF = re.compile(r"<(?:hp:)?linesegarray[^/]*/>")


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

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
