#!/usr/bin/env python3
"""Search and fetch web sources over MCP, then save excerpts in a Word report."""

import argparse
import asyncio
import json
import os
from pathlib import Path
import shutil
import subprocess
from datetime import timedelta
from uuid import uuid4

import httpx
from mcp import ClientSession
from mcp.client.streamable_http import streamable_http_client


MCP_URL = "https://search.parallel.ai/mcp"
USER_AGENT = "OfficeCLI-research-report/1.0 (https://github.com/iOfficeAI/OfficeCLI)"


async def call(session, name, arguments):
    result = await session.call_tool(name, arguments)
    if result.isError:
        raise RuntimeError(f"{name}: {result.content}")
    if result.structuredContent is not None:
        return result.structuredContent
    # Some servers return their JSON only as a text content block.
    for block in result.content:
        if block.type == "text":
            return json.loads(block.text)
    raise RuntimeError(f"{name} returned no JSON content")


async def research(objective, queries, limit):
    session_id = str(uuid4())
    # No auth header, key lookup, or saved MCP client credentials are used.
    async with httpx.AsyncClient(
        headers={"User-Agent": USER_AGENT}, timeout=60,
    ) as client:
        async with streamable_http_client(MCP_URL, http_client=client) as (read, write, _):
            async with ClientSession(
                read, write, read_timeout_seconds=timedelta(seconds=120),
            ) as session:
                await session.initialize()
                tools = {tool.name for tool in (await session.list_tools()).tools}
                if not {"web_search", "web_fetch"} <= tools:
                    raise RuntimeError("The server must provide web_search and web_fetch")
                search = await call(session, "web_search", {
                    "objective": objective, "search_queries": queries,
                    "session_id": session_id,
                })
                sources = search["results"][:limit]
                if not sources:
                    raise RuntimeError("No search results; no report was created")
                fetched = await call(session, "web_fetch", {
                    "urls": [source["url"] for source in sources],
                    "search_queries": queries, "session_id": session_id,
                    "full_content": False,
                })
                return sources, fetched


def paragraph(text, style="Normal"):
    return {"command": "add", "parent": "/body", "type": "paragraph",
            "props": {"text": text, "style": style}}


def write_report(binary, output, objective, sources, fetched):
    pages = {page["url"]: page for page in fetched.get("results", [])}
    errors = {error["url"]: error for error in fetched.get("errors", [])}
    items = [paragraph("Web research sources", "Heading1"), paragraph(objective),
             paragraph("Source excerpts, not a synthesized answer. Verify sources before use.")]
    for index, source in enumerate(sources, 1):
        url = source["url"]
        page = pages.get(url, {})
        items.extend([
            paragraph(f"{index}. {page.get('title') or source.get('title') or url}", "Heading2"),
            paragraph(url),
        ])
        excerpts = page.get("excerpts") or source.get("excerpts") or []
        if url in errors:
            items.append(paragraph(
                f"Page fetch failed ({errors[url]['error_type']}); showing search excerpts."
            ))
        if not excerpts:
            items.append(paragraph("No excerpts returned for this source."))
        items.extend(paragraph(excerpt) for excerpt in excerpts)

    # Keep writes synchronous on disk for this short-lived pipeline.
    env = {**os.environ, "OFFICECLI_NO_AUTO_RESIDENT": "1", "OFFICECLI_SKIP_UPDATE": "1"}
    def cli(*args, payload=None):
        result = subprocess.run(
            [binary, *args], input=payload, text=True, capture_output=True,
            check=True, timeout=120, env=env,
        )
        return result.stdout

    cli("create", str(output))  # No --force: existing documents are protected.
    cli("batch", str(output), "--json", payload=json.dumps(items))
    cli("close", str(output))  # Flush before another program reads the file.
    print(f"Saved {len(sources)} sources to {output}")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("objective", help="The research question or goal")
    parser.add_argument("--query", action="append", required=True,
                        help="Related 3–6 word query; repeat for multiple queries")
    parser.add_argument("--output", type=Path, default=Path("research-report.docx"))
    parser.add_argument("--max-sources", type=int, choices=range(1, 6), default=3)
    args = parser.parse_args()
    if args.output.suffix.lower() != ".docx":
        parser.error("--output must end in .docx")
    if args.output.exists():
        parser.error("output already exists; choose a new filename")
    binary = shutil.which("officecli")
    if binary is None:
        parser.error("install officecli and add it to PATH first")
    sources, fetched = asyncio.run(research(args.objective, args.query, args.max_sources))
    write_report(binary, args.output.resolve(), args.objective, sources, fetched)


if __name__ == "__main__":
    main()
