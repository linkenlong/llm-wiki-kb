"""知识库批量转换。通过 python kb_tools.py convert-batch raw/notes 调用。
批量引擎依赖：python -m pip install "markitdown[docx,pdf,pptx,xlsx]" pymupdf
旧 Office 文件另需 LibreOffice。可选云端转写另需 anthropic、ANTHROPIC_API_KEY、
FALLBACK_MODEL 和显式 --cloud。无云端转写时，疑难 PDF 页截屏供人工复核。
"""
import base64
import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile
from datetime import date
from pathlib import Path

from kb_tools import cmd_stamp, split_fm

def find_soffice() -> str | None:
    path = shutil.which("soffice") or shutil.which("libreoffice")
    if path:
        return path
    for p in (r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
              "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if os.path.exists(p):
            return p
    return None


# ---------------- 配置 ----------------
SOFFICE = find_soffice()
MODEL = os.environ.get("FALLBACK_MODEL", "")
DPI = 150                    # 送给大模型的图片分辨率
MIN_PAGE_CHARS = 30          # 页面文字少于此值，视为无文字层/扫描页
MAX_GARBLE_RATIO = 0.10      # 乱码字符占比上限
MIN_DOC_CHARS = 20           # 非 PDF 文件输出字数下限
LEGACY = {".doc": ".docx", ".ppt": ".pptx", ".xls": ".xlsx"}
SUPPORTED = {".pdf", ".docx", ".pptx", ".xlsx", ".csv", ".html", ".htm",
             ".epub", ".txt", ".json", ".xml", ".md"} | set(LEGACY)

PROMPT = ("将这一页图片内容逐字转写为 Markdown。只转写，不总结、不补充、不解释。"
          "保留标题层级、列表和表格（用 Markdown 表格）；公式用 LaTeX；"
          "看不清的地方写 [?]。只输出 Markdown 本身。")

md_engine = None
_client = None


# ---------------- 工具函数 ----------------
def sha(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()[:16]


def file_hash(p: Path) -> str:
    h = hashlib.sha256()
    with p.open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def garble_ratio(text: str) -> float:
    if not text:
        return 1.0
    bad = sum(1 for c in text if c == "\ufffd" or (ord(c) < 32 and c not in "\n\r\t"))
    return bad / len(text)


def to_md(path: Path) -> str:
    global md_engine
    if md_engine is None:
        from markitdown import MarkItDown
        md_engine = MarkItDown()
    return md_engine.convert_local(str(path)).text_content or ""


def convert_legacy(src: Path, tmp: Path) -> Path:
    """doc/ppt/xls -> docx/pptx/xlsx，用 LibreOffice 无头模式"""
    if not SOFFICE:
        raise RuntimeError("未安装 LibreOffice，无法转换旧格式")
    target = LEGACY[src.suffix.lower()].lstrip(".")
    subprocess.run([SOFFICE, "--headless", "--convert-to", target,
                    "--outdir", str(tmp), str(src)],
                   check=True, capture_output=True, timeout=180)
    return tmp / (src.stem + "." + target)


# ---------------- 大模型兜底 ----------------
def client():
    global _client
    if _client is None:
        import anthropic
        _client = anthropic.Anthropic()
    return _client


def llm_page(png: bytes, cache_dir: Path) -> str | None:
    """单页图片 -> Markdown，按图片哈希缓存"""
    cf = cache_dir / f"page_{sha(MODEL.encode() + png)}.md"
    if cf.exists():
        return cf.read_text(encoding="utf-8")
    if not os.environ.get("ANTHROPIC_API_KEY"):
        return None
    try:
        resp = client().messages.create(
            model=MODEL, max_tokens=4000,
            messages=[{"role": "user", "content": [
                {"type": "image", "source": {"type": "base64", "media_type": "image/png",
                                             "data": base64.b64encode(png).decode()}},
                {"type": "text", "text": PROMPT}]}])
        text = "".join(b.text for b in resp.content if b.type == "text").strip()
    except Exception as e:
        print(f"    大模型调用失败: {e}")
        return None
    if resp.stop_reason == "max_tokens" or len(text) < 5 or "[?]" in text and text.count("[?]") > 10:
        return None  # 低置信度，进人工队列
    with tempfile.NamedTemporaryFile(dir=cache_dir, delete=False, mode="w", encoding="utf-8") as tmp:
        tmp.write(text)
    os.replace(tmp.name, cf)
    return text


# ---------------- PDF 处理 ----------------
def bad_pdf_pages(pdf: Path) -> list[int]:
    import fitz
    bad = []
    with fitz.open(pdf) as doc:
        for i, page in enumerate(doc):
            t = page.get_text()
            if len(t.strip()) < MIN_PAGE_CHARS or garble_ratio(t) > MAX_GARBLE_RATIO:
                bad.append(i)
    return bad


def convert_pdf(pdf: Path, cache_dir: Path, review_dir: Path, report: list, cloud: bool) -> str:
    import fitz
    bad = bad_pdf_pages(pdf)
    if not bad:
        return to_md(pdf)

    parts = []
    with fitz.open(pdf) as doc, tempfile.TemporaryDirectory() as td:
        for i, page in enumerate(doc):
            if i not in bad:
                one = Path(td) / f"p{i}.pdf"
                with fitz.open() as single:
                    single.insert_pdf(doc, from_page=i, to_page=i)
                    single.save(one)
                parts.append(to_md(one).strip())
                continue
            png = page.get_pixmap(dpi=DPI).tobytes("png")
            text = llm_page(png, cache_dir) if cloud else None
            if text:
                parts.append(text)
                report.append((pdf.name, i + 1, "llm_ok"))
            else:
                shot = review_dir / f"{pdf.stem}_p{i + 1}.png"
                shot.write_bytes(png)
                parts.append(f"<!-- NEEDS_REVIEW: page {i + 1}, screenshot: {shot.name} -->")
                report.append((pdf.name, i + 1, "needs_review"))
    return "\n\n".join(parts)



def convert_batch(root: Path, directory: str, cloud: bool) -> int:
    raw = (root / "raw").resolve()
    folder = (root / directory).resolve()
    if not folder.is_dir() or not folder.is_relative_to(raw):
        sys.exit("输入目录须为此知识库中已存在的 raw/ 或其子目录。")
    if cloud and not (MODEL and os.environ.get("ANTHROPIC_API_KEY")):
        sys.exit("云端转写需要设置 FALLBACK_MODEL 和 ANTHROPIC_API_KEY。")
    files = [p for p in sorted(folder.rglob("*"))
             if p.is_file() and p.suffix.lower() in SUPPORTED
             and p.resolve().is_relative_to(raw) and not p.name.endswith(".meta.md")]
    if not files:
        print("该目录中没有支持的资料。")
        return 0

    output = root / "outputs" / "转换"
    output.mkdir(parents=True, exist_ok=True)
    run = Path(tempfile.mkdtemp(prefix=f"{date.today()}-", dir=output))
    cache = output / ".cache"
    cache.mkdir(exist_ok=True)
    report = ["# 批量转换结果", "", "原文路径相对于知识库根目录。阅读稿和复核截图放在本批次的“文件/原文路径/”下。", ""]
    failed = review = 0
    for src in files:
        rel = src.relative_to(root)
        dest = run / "文件" / rel
        try:
            dest.mkdir(parents=True)
            if cmd_stamp(str(src), ""):
                raise ValueError("原件与已登记哈希不符，请核对归档版本")
            digest = file_hash(src)
            events = []
            # 缓存按输入格式与转写配置隔离。逐页转写仍保留状态记录。
            cache_key = f"{src.suffix.lower()}-{digest}"
            cached = cache / f"{cache_key}.md"
            if src.suffix.lower() == ".pdf":
                text = convert_pdf(src, cache, dest, events, cloud)
            elif cached.is_file():
                text = cached.read_text(encoding="utf-8")
            else:
                with tempfile.TemporaryDirectory() as temp:
                    work = convert_legacy(src, Path(temp)) if src.suffix.lower() in LEGACY else src
                    if work.suffix.lower() in {".md", ".txt"}:
                        _, text, _ = split_fm(work.read_text(encoding="utf-8-sig"))
                    else:
                        text = to_md(work)
                if len(text.strip()) >= MIN_DOC_CHARS and garble_ratio(text) <= MAX_GARBLE_RATIO:
                    with tempfile.NamedTemporaryFile(dir=cache, delete=False, mode="w",
                                                     encoding="utf-8") as tmp:
                        tmp.write(text)
                    os.replace(tmp.name, cached)
            if not text.strip() or (not events and (len(text.strip()) < MIN_DOC_CHARS
                                                     or garble_ratio(text) > MAX_GARBLE_RATIO)):
                events.append((src.name, 0, "needs_review_whole_file"))
                text = "<!-- NEEDS_REVIEW: whole file -->\n" + text
            needs_review = any(e[2].startswith("needs_review") for e in events)
            status = "待复核" if needs_review else "完成"
            header = (f"---\nconverted_from: {json.dumps(rel.as_posix(), ensure_ascii=False)}\n"
                      f"source_file_sha256: {digest}\nconverted: {date.today()}\n"
                      f"review_status: {status}\n---\n")
            page = dest / "阅读稿.md"
            page.write_text(header + text, encoding="utf-8")
            cmd_stamp(str(page), "")
            review += needs_review
            report.append(f"- {rel.as_posix()} — {status}")
            for _, number, event in events:
                report.append(f"  - {('第 ' + str(number) + ' 页') if number else '整份'} — {event}")
        except Exception as exc:
            failed += 1
            report.append(f"- {rel.as_posix()} — 失败：{exc}")
    report.append(f"\n总计 {len(files)} 份，完成 {len(files)-review-failed} 份，待复核 {review} 份，失败 {failed} 份。")
    path = run / "report.md"
    path.write_text("\n".join(report) + "\n", encoding="utf-8")
    print(report[-1])
    print(f"报告：{path}")
    return 1 if failed else 2 if review else 0
