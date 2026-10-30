"""Focused CLI regression: bulk revision decisions include document stories.

Usage: python3 tests/revision_document_stories.py /path/to/officecli
   or: python3 tests/revision_document_stories.py dotnet /path/to/officecli.dll
"""
import os
import pathlib
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile

COMMAND = sys.argv[1:]
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml."


def pair(label, first_id):
    return (
        f'<w:ins w:id="{first_id}" w:author="Reviewer" w:date="2026-01-01T00:00:00Z">'
        f'<w:r><w:t>{label} added.</w:t></w:r></w:ins>'
        f'<w:del w:id="{first_id + 1}" w:author="Reviewer" w:date="2026-01-01T00:00:00Z">'
        f'<w:r><w:delText>{label} removed.</w:delText></w:r></w:del>'
    )


def fixture(file, footer_only=False):
    stories = {
        "word/document.xml": f'<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body>'
        f'<w:p>{"<w:r><w:t>Body unchanged.</w:t></w:r>" if footer_only else pair("Body", 1)}</w:p>'
        '<w:sectPr><w:headerReference w:type="default" r:id="header1"/>'
        '<w:headerReference w:type="even" r:id="header2"/>'
        '<w:footerReference w:type="default" r:id="footer1"/>'
        '<w:footerReference w:type="even" r:id="footer2"/></w:sectPr></w:body></w:document>',
        "word/header1.xml": f'<w:hdr xmlns:w="{W}"><w:p>'
        f'{"<w:r><w:t>Header unchanged.</w:t></w:r>" if footer_only else pair("Header", 1)}</w:p></w:hdr>',
        "word/footer1.xml": f'<w:ftr xmlns:w="{W}"><w:p>'
        + "".join(pair(f"Footer {i}", 50 + i * 2) for i in range(3)) + '</w:p></w:ftr>',
        "word/header2.xml": f'<w:hdr xmlns:w="{W}"><w:p><w:r><w:rPr><w:b/></w:rPr>'
        '<w:t xml:space="preserve">  Unchanged header  </w:t></w:r></w:p></w:hdr>',
        "word/footer2.xml": f'<w:ftr xmlns:w="{W}"><w:p><w:r><w:rPr><w:i/></w:rPr>'
        '<w:t>Unchanged footer</w:t></w:r></w:p></w:ftr>',
        "word/footnotes.xml": f'<w:footnotes xmlns:w="{W}"><w:footnote w:id="1"><w:p>'
        f'{"<w:r><w:t>Note unchanged.</w:t></w:r>" if footer_only else pair("Footnote", 70)}</w:p></w:footnote></w:footnotes>',
        "word/endnotes.xml": f'<w:endnotes xmlns:w="{W}"><w:endnote w:id="1"><w:p>'
        f'{"<w:r><w:t>Note unchanged.</w:t></w:r>" if footer_only else pair("Endnote", 80)}</w:p></w:endnote></w:endnotes>',
        "word/comments.xml": f'<w:comments xmlns:w="{W}"><w:comment w:id="0" w:author="Comment author">'
        '<w:p><w:r><w:t>Retain this review comment.</w:t></w:r></w:p></w:comment></w:comments>',
    }
    part_types = {"document": "document.main", "header": "header", "footer": "footer",
                  "footnotes": "footnotes", "endnotes": "endnotes", "comments": "comments"}
    overrides = []
    relations = []
    for name in stories:
        stem = pathlib.PurePosixPath(name).stem
        kind = stem.rstrip("12")
        overrides.append(f'<Override PartName="/{name}" ContentType="{MIME}{part_types[kind]}+xml"/>')
        if kind != "document":
            relations.append(f'<Relationship Id="{stem}" Type="{R}/{kind}" Target="{stem}.xml"/>')
    parts = {**stories,
        "[Content_Types].xml": '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        + "".join(overrides) + '</Types>',
        "_rels/.rels": f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="main" Type="{R}/officeDocument" Target="word/document.xml"/></Relationships>',
        "word/_rels/document.xml.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(relations) + '</Relationships>',
    }
    with zipfile.ZipFile(file, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, xml in parts.items():
            archive.writestr(name, xml)


def read_parts(file):
    with zipfile.ZipFile(file) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def markers(xml):
    root = ET.fromstring(xml)
    return [node for node in root.iter() if node.tag in {f"{{{W}}}ins", f"{{{W}}}del"}]


def run(file, selector, action):
    env = {**os.environ, "OFFICECLI_SKIP_UPDATE": "1", "OFFICECLI_NO_AUTO_RESIDENT": "1"}
    result = subprocess.run([*COMMAND, "set", str(file), selector, "--prop", f"revision.action={action}"],
                            env=env, capture_output=True, text=True, timeout=60)
    assert result.returncode == 0, (result.stdout, result.stderr)


def content_signature(xml, ignore_generated_ids=False):
    def node_signature(node):
        generated_ids = {"{http://schemas.microsoft.com/office/word/2010/wordml}paraId",
                         "{http://schemas.microsoft.com/office/word/2010/wordml}textId"}
        return (node.tag, sorted((k, v) for k, v in node.attrib.items()
                                if not ignore_generated_ids or k not in generated_ids),
                node.text, node.tail, [node_signature(child) for child in node])
    return node_signature(ET.fromstring(xml))


def prepared_fixture(file, footer_only=False):
    fixture(file, footer_only)
    original = read_parts(file)
    # Editable-open normalizes paragraph IDs before every CLI mutation. A
    # no-match selector establishes that existing baseline without a decision.
    run(file, "/revision[@author=No matching reviewer]", "accept")
    prepared = read_parts(file)
    for name in ("word/header2.xml", "word/footer2.xml", "word/comments.xml"):
        assert content_signature(prepared[name], True) == content_signature(original[name], True)
    return prepared


with tempfile.TemporaryDirectory(prefix="officecli-story-revisions-") as temporary:
    root = pathlib.Path(temporary)
    for action in ("accept", "reject"):
        for footer_only in (False, True):
            file = root / f"{action}-{footer_only}.docx"
            before = prepared_fixture(file, footer_only)
            assert len(markers(before["word/footer1.xml"])) == 6
            run(file, "/revision", action)
            after = read_parts(file)  # New ZIP read after the CLI process exits.
            counts = {name: len(markers(after[name])) for name in
                      ("word/document.xml", "word/header1.xml", "word/footer1.xml", "word/footnotes.xml", "word/endnotes.xml")}
            print(f"{action}: saved revision counts: {counts}")
            for name in ("word/document.xml", "word/header1.xml", "word/footer1.xml", "word/footnotes.xml", "word/endnotes.xml"):
                assert not markers(after[name]), f"{action}: revisions remain in {name}"
                if not footer_only or name == "word/footer1.xml":
                    text = "".join(ET.fromstring(after[name]).itertext())
                    assert ("added." in text) == (action == "accept"), (name, text)
                    assert ("removed." in text) == (action == "reject"), (name, text)
            for name in ("word/header2.xml", "word/footer2.xml", "word/comments.xml"):
                # SDK namespace placement can change on any save. Compare all
                # elements, attributes (including IDs), text and formatting.
                assert content_signature(after[name]) == content_signature(before[name]), f"Untouched part changed: {name}"
            run(file, "/revision", action)
            assert not markers(read_parts(file)["word/footer1.xml"])
            print(f"PASS: {action}, footer_only={footer_only}; saved stories and untouched parts verified")
    for selector, remaining in (("/revision[@id=1]", 1), ("/revision[@author=Reviewer]", 0)):
        file = root / "targeted.docx"
        before = prepared_fixture(file)
        run(file, selector, "accept")
        after = read_parts(file)
        assert len(markers(after["word/document.xml"])) == remaining
        for name in ("word/header1.xml", "word/footer1.xml", "word/footnotes.xml", "word/endnotes.xml"):
            assert content_signature(after[name]) == content_signature(before[name]), f"Targeted decision expanded to {name}"
        print(f"PASS: targeted body selector {selector} keeps other stories unchanged")
