"""Run with: python3 tests/word_section_margin_inheritance.py officecli [or dotnet CLI.dll]."""
import os
import pathlib
import re
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile

COMMAND = sys.argv[1:]
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
ENV = {**os.environ, "OFFICECLI_SKIP_UPDATE": "1", "OFFICECLI_NO_AUTO_RESIDENT": "1"}


def cli(*args):
    result = subprocess.run(COMMAND + list(map(str, args)), env=ENV,
                            capture_output=True, text=True, timeout=60)
    assert result.returncode == 0, result.stdout + result.stderr


with tempfile.TemporaryDirectory(prefix="word-section-margins-") as directory:
    for number, source in enumerate([
        dict(top="1100", bottom="1200", left="1300", right="1400", header="450", footer="550", gutter="150"),
        dict(top="-120", bottom="-240", left="1300", right="1400", header="0", footer="0", gutter="0"),
    ]):
        path = pathlib.Path(directory) / f"source-{number}.docx"
        cli("create", path, "--locale", "en-US")
        with zipfile.ZipFile(path) as archive:
            parts = {info.filename: archive.read(info) for info in archive.infolist()}
        xml = parts["word/document.xml"].decode("utf-8")
        replacement = "<w:pgMar " + " ".join(f'w:{key}="{value}"' for key, value in source.items()) + "/>"
        xml, count = re.subn(r"<w:pgMar\b[^>]*/>", replacement, xml)
        assert count == 1
        parts["word/document.xml"] = xml.encode("utf-8")
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
            for name, data in parts.items():
                archive.writestr(name, data)
        cli("add", path, "/body", "--type", "section")
        with zipfile.ZipFile(path) as archive:
            root = ET.fromstring(archive.read("word/document.xml"))
        margins = [{key.removeprefix(W): value for key, value in node.attrib.items()}
                   for node in root.iter(W + "pgMar")]
        assert len(margins) == 2 and all(margin == source for margin in margins), margins
    print("PASS: saved sections inherit all seven margins, including custom, zero and signed source values")
