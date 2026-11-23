"""Run with: python3 tests/word_page_margins.py /path/to/officecli [or dotnet CLI.dll]."""
import os
import pathlib
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile

COMMAND = sys.argv[1:]
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
REQUIRED = {"top", "right", "bottom", "left", "header", "footer", "gutter"}
ENV = {**os.environ, "OFFICECLI_SKIP_UPDATE": "1", "OFFICECLI_NO_AUTO_RESIDENT": "1"}


def cli(*args):
    result = subprocess.run(COMMAND + list(map(str, args)), env=ENV,
                            capture_output=True, text=True, timeout=60)
    assert result.returncode == 0, result.stdout + result.stderr


def document(path):
    with zipfile.ZipFile(path) as archive:
        return ET.fromstring(archive.read("word/document.xml"))


def rewrite(path, change):
    with zipfile.ZipFile(path) as archive:
        parts = {info.filename: archive.read(info) for info in archive.infolist()}
    root = ET.fromstring(parts["word/document.xml"])
    change(root)
    parts["word/document.xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)


def margins(path):
    return [{key.removeprefix(W): value for key, value in node.attrib.items()}
            for node in document(path).iter(W + "pgMar")]


with tempfile.TemporaryDirectory(prefix="word-page-margins-") as directory:
    root = pathlib.Path(directory)
    for label, options in [("default", ["--locale", "en-US"]),
                           ("minimal", ["--minimal", "--locale", "en-US"]),
                           ("rtl", ["--locale", "ar"])]:
        path = root / (label + ".docx")
        cli("create", path, *options)
        actual = margins(path)
        assert len(actual) == 1 and REQUIRED <= actual[0].keys(), actual
        assert {key: actual[0][key] for key in ("header", "footer", "gutter")} == {
            "header": "720", "footer": "720", "gutter": "0"}, actual

    path = root / "default.docx"
    cli("set", path, "/", "--prop", "pageWidth=12000", "--prop", "pageHeight=17000",
        "--prop", "marginTop=1100", "--prop", "marginBottom=1200",
        "--prop", "marginLeft=1300", "--prop", "marginRight=1400")
    assert margins(path)[0] == dict(top="1100", bottom="1200", left="1300", right="1400",
                                   header="720", footer="720", gutter="0")

    def remove_margin(root):
        section = root.find("./" + W + "body/" + W + "sectPr")
        section.remove(section.find(W + "pgMar"))

    rewrite(path, remove_margin)
    cli("set", path, "/", "--prop", "marginTop=1000")
    actual = margins(path)[0]
    assert REQUIRED <= actual.keys() and actual["top"] == "1000", actual

    path = root / "add-missing.docx"
    cli("create", path, "--locale", "en-US")
    rewrite(path, remove_margin)
    cli("add", path, "/body", "--type", "section")
    assert REQUIRED <= margins(path)[-1].keys(), margins(path)
    print("PASS: default/minimal/RTL creation, page settings, absent body margins")
