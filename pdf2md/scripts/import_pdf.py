#!/usr/bin/env python3
"""PDF를 <root>/pdf/ 에 복사하고 <root>/md/ 에 페이지별 Markdown으로 변환한다.

사용: import_pdf.py [--root DIR] [--force] PDF [PDF ...]
기존 md 포맷(pdftotext -layout, 페이지마다 <!-- p.N --> / ## p.N)과 바이트 단위로 동일하게 생성한다.
"""
import argparse, re, shutil, subprocess, sys
from pathlib import Path

DEFAULT_ROOT = Path.home() / "Desktop" / "이전제안서"


def page_count(pdf: Path) -> int:
    out = subprocess.run(["pdfinfo", str(pdf)], capture_output=True, text=True, check=True).stdout
    return int(re.search(r"Pages:\s+(\d+)", out).group(1))


def to_markdown(pdf: Path) -> tuple[str, int, int]:
    n = page_count(pdf)
    txt = subprocess.run(["pdftotext", "-layout", str(pdf), "-"],
                         capture_output=True, text=True, check=True).stdout
    pages = txt.split("\f")
    parts = [f"# {pdf.stem}\n\n원본: ../pdf/{pdf.name} · {n}쪽\n"]
    empty = 0
    for i in range(n):
        t = pages[i].strip() if i < len(pages) else ""
        empty += not t
        parts.append(f"<!-- p.{i+1} -->\n## p.{i+1}\n\n{t}\n")
    return "\n".join(parts), n, empty


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdfs", nargs="+", type=Path)
    ap.add_argument("--root", type=Path, default=DEFAULT_ROOT)
    ap.add_argument("--force", action="store_true", help="이미 있는 파일도 덮어쓰기")
    a = ap.parse_args()

    pdf_dir, md_dir = a.root / "pdf", a.root / "md"
    pdf_dir.mkdir(parents=True, exist_ok=True)
    md_dir.mkdir(parents=True, exist_ok=True)
    failed = False

    for src in a.pdfs:
        src = src.expanduser().resolve()
        if not src.is_file() or src.suffix.lower() != ".pdf":
            print(f"[SKIP] PDF 파일이 아님: {src}"); failed = True; continue
        dst_pdf = pdf_dir / (src.stem + ".pdf")
        dst_md = md_dir / (src.stem + ".md")
        if not a.force and (dst_pdf.exists() and dst_pdf.resolve() != src or dst_md.exists()):
            print(f"[EXISTS] 이미 존재: {dst_pdf.name} / {dst_md.name} (덮어쓰려면 --force)")
            failed = True; continue
        if dst_pdf.resolve() != src:
            shutil.copy2(src, dst_pdf)
        md, n, empty = to_markdown(dst_pdf)
        dst_md.write_text(md, encoding="utf-8")
        warn = f"  ⚠ 텍스트 없는 페이지 {empty}/{n} (스캔 이미지일 수 있음)" if empty > n // 2 else ""
        print(f"[OK] {src.name} → pdf/{dst_pdf.name}, md/{dst_md.name} ({n}쪽, 빈 페이지 {empty}){warn}")

    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
