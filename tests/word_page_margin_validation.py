"""Run with: python3 tests/word_page_margin_validation.py officecli [or dotnet CLI.dll]."""
import hashlib
import json
import os
import pathlib
import re
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile

COMMAND = sys.argv[1:]
ATTRIBUTES = ("top", "right", "bottom", "left", "header", "footer", "gutter")
ENV = {**os.environ, "OFFICECLI_SKIP_UPDATE": "1", "OFFICECLI_NO_AUTO_RESIDENT": "1"}


def run(*args):
    return subprocess.run(COMMAND + list(map(str, args)), env=ENV,
                          capture_output=True, text=True, timeout=60)


with tempfile.TemporaryDirectory(prefix="word-page-margin-validation-") as directory:
    root = pathlib.Path(directory)
    template = root / "template.docx"
    result = run("create", template, "--locale", "en-US")
    assert result.returncode == 0, result.stdout + result.stderr
    with zipfile.ZipFile(template) as archive:
        parts = {info.filename: archive.read(info) for info in archive.infolist()}
    # Custom XML is opaque data, even when its payload uses a Word namespace.
    parts["customXml/item1.xml"] = (b'<data xmlns:w="http://schemas.openxmlformats.org/'
                                    b'wordprocessingml/2006/main"><w:pgMar/></data>')
    rel_name = "word/_rels/document.xml.rels"
    rels = ET.fromstring(parts[rel_name])
    ET.SubElement(rels, "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship", {
        "Id": "rIdSyntheticCustomXml", "Target": "../customXml/item1.xml",
        "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"})
    ET.register_namespace("", "http://schemas.openxmlformats.org/package/2006/relationships")
    parts[rel_name] = ET.tostring(rels, encoding="utf-8", xml_declaration=True)
    types = ET.fromstring(parts["[Content_Types].xml"])
    ET.SubElement(types, "{http://schemas.openxmlformats.org/package/2006/content-types}Override", {
        "PartName": "/customXml/item1.xml", "ContentType": "application/xml"})
    ET.register_namespace("", "http://schemas.openxmlformats.org/package/2006/content-types")
    parts["[Content_Types].xml"] = ET.tostring(types, encoding="utf-8", xml_declaration=True)

    for missing in (None, *ATTRIBUTES, "all", "element"):
        document = parts["word/document.xml"].decode("utf-8")
        match = re.search(r"<w:pgMar\b[^>]*/>", document)
        assert match, document
        margin = match.group()
        for attribute in ATTRIBUTES:
            if not re.search(r"\bw:" + attribute + r'="', margin):
                value = "720" if attribute != "gutter" else "0"
                margin = margin[:-2] + f' w:{attribute}="{value}"/>'
        if missing == "element":
            margin = ""
        elif missing == "all":
            margin = "<w:pgMar/>"
        elif missing:
            margin = re.sub(r'\s+w:' + missing + r'="[^"]*"', "", margin)
        document = document[:match.start()] + margin + document[match.end():]
        path = root / (str(missing) + ".docx")
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
            for name, data in parts.items():
                archive.writestr(name, document.encode("utf-8")
                                 if name == "word/document.xml" else data)
        before = hashlib.sha256(path.read_bytes()).hexdigest()
        result = run("validate", path, "--json")
        payload = json.loads(result.stdout)
        assert "data" in payload, result.stdout + result.stderr
        errors = payload["data"]["errors"]
        expected = list(ATTRIBUTES) if missing == "all" else [missing] if missing in ATTRIBUTES else []
        actual = [error for error in errors if error["type"] == "Sch_MissRequiredAttribute"]
        assert len(actual) == len(expected), (missing, result.stdout)
        for attribute in expected:
            assert any("w:" + attribute in error["description"] for error in actual), result.stdout
        assert (result.returncode != 0) == bool(expected), (missing, result.stdout)
        assert hashlib.sha256(path.read_bytes()).hexdigest() == before, "validation modified source"
    print("PASS: every missing margin attribute, all missing, complete margins, omitted optional element, source bytes")
