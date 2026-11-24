"""用合成文档验证同目录发布包。测试依赖与批量转换模块相同。"""
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile
import hashlib
import os
import subprocess
import sys

import pymupdf
from openpyxl import Workbook
from pptx import Presentation

BASE = Path(__file__).resolve().parent
PACKAGE = BASE.parent / (BASE.name + ".zip")
MARK = "Synthetic conversion test document with searchable office content"


def check():
    with TemporaryDirectory() as tmp:
        root = Path(tmp)
        kit = root / "kit"
        db = root / "db"
        with ZipFile(PACKAGE) as z:
            assert z.testzip() is None
            z.extractall(kit)
        env = dict(os.environ, PYTHONIOENCODING="utf-8", ANTHROPIC_API_KEY="")
        def run(script, *args, expected=0):
            p = subprocess.run([sys.executable, str(script), *map(str, args)],
                               cwd=db if db.exists() else root, env=env,
                               capture_output=True, encoding="utf-8")
            if p.returncode != expected:
                raise AssertionError((args, p.returncode, p.stdout, p.stderr))
            return p.stdout

        run(kit / "setup_kb.py", db, "--no-git")
        kb = db / "kb_tools.py"
        raw = db / "raw/notes"
        def batch(path="raw/notes", expected=0):
            msg = run(kb, "convert-batch", path, expected=expected)
            report = Path(msg.split("报告：")[-1].strip())
            return report, report.read_text(encoding="utf-8")

        a = raw / "A"
        b = raw / "B"
        a.mkdir()
        b.mkdir()
        docx = a / "sample.docx"
        with ZipFile(docx, "w") as z:
            z.writestr("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
            z.writestr("_rels/.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
            z.writestr("word/document.xml", f'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>{MARK}</w:t></w:r></w:p></w:body></w:document>')
        book = Workbook()
        book.active.append(["Example", MARK])
        book.save(a / "sample.xlsx")
        slides = Presentation()
        slides.slides.add_slide(slides.slide_layouts[0]).shapes.title.text = MARK
        slides.save(b / "sample.pptx")
        (b / "sample.txt").write_text(MARK, encoding="utf-8")
        (a / "sample.txt").write_text(MARK + " another folder", encoding="utf-8")
        (a / "sample.html").write_text(f"<h1>{MARK}</h1>", encoding="utf-8")
        with pymupdf.open() as pdf:
            pdf.new_page().insert_text((32, 40), MARK)
            pdf.new_page().draw_rect((50, 50, 100, 100), fill=(0, 0, 0))
            pdf.save(a / "sample.pdf")
        digest = hashlib.sha256(docx.read_bytes()).hexdigest()
        report, summary = batch(expected=2)
        assert "完成 6 份，待复核 1 份，失败 0 份" in summary, summary
        for page in (report.parent / "文件").rglob("阅读稿.md"):
            text = page.read_text(encoding="utf-8")
            assert MARK in text and "converted_from:" in text and "sha256:" in text, page
        assert len(list(report.parent.rglob("阅读稿.md"))) == 7
        assert len(list(report.parent.rglob("*.png"))) == 1 and "第 2 页" in summary
        assert hashlib.sha256(docx.read_bytes()).hexdigest() == digest
        assert digest in (a / "sample.docx.meta.md").read_text(encoding="utf-8")
        report2, _ = batch("raw/notes/B")
        assert report2 != report and report.is_file()
        run(kb, "convert", docx)
        assert MARK in docx.with_suffix(".md").read_text(encoding="utf-8")
        run(kb, "index")
        assert "0 个错误" in run(kb, "lint")
        run(kb, "convert-batch", "outputs", expected=1)
        (b / "broken.pdf").write_bytes(b"invalid PDF")
        _, summary = batch("raw/notes/B", expected=1)
        assert "失败 1 份" in summary and "完成 2 份" in summary, summary
        print("通过：发布包安装、七份样本六种格式、同名隔离、混合 PDF 逐页复核、哈希、重复运行、旧命令、lint、目录边界、坏文件隔离。")


if __name__ == "__main__":
    check()
