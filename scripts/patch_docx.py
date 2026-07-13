from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path
from tempfile import TemporaryDirectory
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}


def qn(tag: str) -> str:
    return f"{{{W}}}{tag}"


def ensure_rfonts(rpr: etree._Element, size_half_points: int | None = None) -> None:
    rfonts = rpr.find(qn("rFonts"))
    if rfonts is None:
        rfonts = etree.SubElement(rpr, qn("rFonts"))
    for key, val in {
        "ascii": "Times New Roman",
        "hAnsi": "Times New Roman",
        "cs": "Times New Roman",
        "eastAsia": "SimSun",
    }.items():
        rfonts.set(qn(key), val)
    if size_half_points:
        for name in ("sz", "szCs"):
            el = rpr.find(qn(name))
            if el is None:
                el = etree.SubElement(rpr, qn(name))
            el.set(qn("val"), str(size_half_points))


def patch_styles(xml: bytes) -> bytes:
    root = etree.fromstring(xml)
    for style in root.findall("w:style", NS):
        style_id = style.get(qn("styleId"), "")
        rpr = style.find("w:rPr", NS)
        if rpr is None:
            rpr = etree.SubElement(style, qn("rPr"))
        size = 20 if style_id in {"FootnoteText", "FootnoteReference"} else None
        ensure_rfonts(rpr, size)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")


def patch_document(xml: bytes) -> bytes:
    root = etree.fromstring(xml)
    for rpr in root.findall(".//w:rPr", NS):
        ensure_rfonts(rpr)
    # Ensure tables do not split rows across pages.
    for tr in root.findall(".//w:tr", NS):
        trpr = tr.find("w:trPr", NS)
        if trpr is None:
            trpr = etree.SubElement(tr, qn("trPr"))
        if trpr.find("w:cantSplit", NS) is None:
            etree.SubElement(trpr, qn("cantSplit"))
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")


def patch(path: Path) -> None:
    with TemporaryDirectory() as tmp:
        tmpdir = Path(tmp)
        with zipfile.ZipFile(path) as zin:
            zin.extractall(tmpdir)
        styles = tmpdir / "word" / "styles.xml"
        document = tmpdir / "word" / "document.xml"
        if styles.exists():
            styles.write_bytes(patch_styles(styles.read_bytes()))
        if document.exists():
            document.write_bytes(patch_document(document.read_bytes()))
        rebuilt = path.with_suffix(".rebuilt.docx")
        with zipfile.ZipFile(rebuilt, "w", zipfile.ZIP_DEFLATED) as zout:
            for p in sorted(tmpdir.rglob("*")):
                if p.is_file():
                    zout.write(p, p.relative_to(tmpdir).as_posix())
        rebuilt.replace(path)


def validate(path: Path) -> dict[str, int | str]:
    with zipfile.ZipFile(path) as zf:
        bad = zf.testzip()
        if bad:
            raise RuntimeError(f"Corrupt ZIP member: {bad}")
        names = set(zf.namelist())
        for required in {"[Content_Types].xml", "word/document.xml", "word/styles.xml"}:
            if required not in names:
                raise RuntimeError(f"Missing {required}")
        document = zf.read("word/document.xml")
        text = " ".join(etree.fromstring(document).xpath("//w:t/text()", namespaces=NS))
        footnote_refs = len(etree.fromstring(document).xpath("//w:footnoteReference", namespaces=NS))
        footnotes = 0
        if "word/footnotes.xml" in names:
            fnroot = etree.fromstring(zf.read("word/footnotes.xml"))
            footnotes = len([n for n in fnroot.xpath("//w:footnote", namespaces=NS)
                             if int(n.get(qn("id"), "-1")) >= 0])
        if "КРИЗИС МОНАРХИЧЕСКОГО ПРОЕКТА" not in text and "КОДИРОВОЧНАЯ МАТРИЦА" not in text:
            raise RuntimeError("Expected title not found")
        return {
            "file": path.name,
            "characters": len(text),
            "footnote_references": footnote_refs,
            "footnotes": footnotes,
        }


def main() -> None:
    if len(sys.argv) < 2:
        raise SystemExit("Usage: patch_docx.py FILE [FILE ...]")
    reports = []
    for arg in sys.argv[1:]:
        p = Path(arg)
        patch(p)
        reports.append(validate(p))
    report_path = Path("outputs/validation_report.txt")
    report_path.parent.mkdir(exist_ok=True)
    report_path.write_text("\n".join(str(r) for r in reports) + "\n", encoding="utf-8")
    for report in reports:
        print(report)


if __name__ == "__main__":
    main()
