#!/usr/bin/env python3
"""
Translate README.md into multiple languages using the OpenAI API.

Usage:
    python translate_readme.py --source README.md \
        --targets README_zh.md:Chinese README_ja.md:Japanese README_ko.md:Korean

Environment variables:
    OPENAI_API_KEY  – required OpenAI API key
    OPENAI_BASE_URL – optional base URL override (default: https://api.openai.com/v1)
    OPENAI_MODEL    – optional model name (default: gpt-4.1-mini)
"""

import argparse
import os
import sys
import time

try:
    from openai import OpenAI
except ImportError:
    print("openai package not found. Install with: pip install openai", file=sys.stderr)
    sys.exit(1)

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

DEFAULT_MODEL = "gpt-4.1-mini"

# Language names shown in the system prompt.
LANGUAGE_LABELS = {
    "zh": "Simplified Chinese",
    "ko": "Korean",
    "ja": "Japanese",
}

# Lines that contain only a language-switcher bar (e.g. "**English** | [中文](...) | ...")
# should be replaced with the equivalent bar for the target language.
LANGUAGE_SWITCHER_PATTERNS = {
    "zh": "**中文** | [English](README.md) | [日本語](README_ja.md) | [한국어](README_ko.md)",
    "ja": "[English](README.md) | [中文](README_zh.md) | **日本語** | [한국어](README_ko.md)",
    "ko": "[English](README.md) | [中文](README_zh.md) | [日本語](README_ja.md) | **한국어**",
}

# The English switcher line we look for (must match exactly as it appears in README.md).
ENGLISH_SWITCHER = "**English** | [中文](README_zh.md) | [日本語](README_ja.md) | [한국어](README_ko.md)"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def derive_lang_code(filename: str) -> str | None:
    """Extract the language code from a filename like README_zh.md -> zh."""
    base = os.path.basename(filename)
    # Expect pattern: README_<lang>.md
    name, _ = os.path.splitext(base)
    parts = name.split("_", 1)
    return parts[1].lower() if len(parts) == 2 else None


def replace_switcher(content: str, lang_code: str) -> str:
    """Swap out the English language-switcher line for the target language."""
    switcher = LANGUAGE_SWITCHER_PATTERNS.get(lang_code)
    if switcher and ENGLISH_SWITCHER in content:
        content = content.replace(ENGLISH_SWITCHER, switcher)
    return content


def translate(client: OpenAI, model: str, source_text: str, language: str) -> str:
    """Call the OpenAI chat API to translate *source_text* into *language*."""
    system_prompt = (
        f"You are a professional technical translator. "
        f"Translate the following Markdown document from English into {language}. "
        "Rules:\n"
        "- Preserve ALL Markdown formatting, HTML tags, links, image tags, badges, "
        "  code blocks, and inline code exactly as they appear.\n"
        "- Do NOT translate code snippets, command-line examples, file names, "
        "  proper nouns (brand names, product names), or URLs.\n"
        "- Translate only the natural-language prose and headings.\n"
        "- Return ONLY the translated Markdown with no extra commentary."
    )

    for attempt in range(3):
        try:
            response = client.chat.completions.create(
                model=model,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": source_text},
                ],
                temperature=0.2,
            )
            return response.choices[0].message.content
        except Exception as exc:  # noqa: BLE001
            if attempt < 2:
                wait = 2 ** attempt * 5
                print(f"  Attempt {attempt + 1} failed ({exc}). Retrying in {wait}s…", file=sys.stderr)
                time.sleep(wait)
            else:
                raise


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(description="Translate README.md into multiple languages.")
    parser.add_argument("--source", default="README.md", help="Path to the English source README.")
    parser.add_argument(
        "--targets",
        nargs="+",
        metavar="FILE:LANGUAGE",
        default=["README_zh.md:Chinese", "README_ja.md:Japanese", "README_ko.md:Korean"],
        help="Target file/language pairs in FILE:LANGUAGE format.",
    )
    parser.add_argument("--model", default=os.environ.get("OPENAI_MODEL", DEFAULT_MODEL))
    args = parser.parse_args()

    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        print("ERROR: OPENAI_API_KEY environment variable is not set.", file=sys.stderr)
        sys.exit(1)

    base_url = os.environ.get("OPENAI_BASE_URL")
    client = OpenAI(api_key=api_key, **({"base_url": base_url} if base_url else {}))

    source_path = args.source
    if not os.path.isfile(source_path):
        print(f"ERROR: Source file not found: {source_path}", file=sys.stderr)
        sys.exit(1)

    with open(source_path, encoding="utf-8") as f:
        source_content = f.read()

    any_error = False
    for target_spec in args.targets:
        if ":" not in target_spec:
            print(f"WARNING: Skipping malformed target spec '{target_spec}' (expected FILE:LANGUAGE).", file=sys.stderr)
            continue

        target_file, language = target_spec.split(":", 1)
        lang_code = derive_lang_code(target_file)
        print(f"Translating → {target_file} ({language})…")

        try:
            translated = translate(client, args.model, source_content, language)
            # Fix the language-switcher bar so it highlights the correct language.
            if lang_code:
                translated = replace_switcher(translated, lang_code)

            with open(target_file, "w", encoding="utf-8") as f:
                f.write(translated)
            print(f"  ✓ Written {target_file}")
        except Exception as exc:  # noqa: BLE001
            print(f"  ✗ Failed to translate {target_file}: {exc}", file=sys.stderr)
            any_error = True

    if any_error:
        sys.exit(1)


if __name__ == "__main__":
    main()
