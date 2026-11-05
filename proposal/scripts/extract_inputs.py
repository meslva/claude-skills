"""RFP 입력 파일을 읽기 쉬운 텍스트와 이미지로 추출한다.

사용: python extract_inputs.py <입력 파일 또는 폴더> <출력 폴더> [--render-all] [--min-chars 150]

- PDF: <이름>_pages.txt (페이지 표시 포함), 글자가 적은 페이지는 <이름>_p{번호}.png 로 렌더링
       (--render-all 이면 모든 페이지 렌더링). 구성도·간트·표 이미지 페이지를 놓치지 않기 위함.
- XLSX: <이름>_<시트>.txt (행마다 ' | ' 구분, 줄바꿈은 ' / '), 시트별 행·열 수 요약,
        시트에 붙은 이미지는 <이름>_<시트>_img{n}.png
- DOCX: <이름>.txt (문단 + 표)
- HWP: 읽을 수 없으므로 목록에만 남긴다(PDF 변환 요청).
마지막에 _summary.txt 에 파일별 요약을 쓴다.
"""
import argparse
import os
import re
import sys


def safe(name):
    return re.sub(r'[\\/:*?"<>|\s]+', "_", name)[:60]


def do_pdf(path, out, render_all, min_chars, summary):
    import pdfplumber
    import pypdfium2 as pdfium
    base = safe(os.path.splitext(os.path.basename(path))[0])
    low = []
    with pdfplumber.open(path) as pdf, open(os.path.join(out, f"{base}_pages.txt"), "w",
                                             encoding="utf-8") as f:
        n = len(pdf.pages)
        for i, pg in enumerate(pdf.pages, 1):
            t = pg.extract_text() or ""
            f.write(f"\n=== p{i} ===\n{t}\n")
            body = re.sub(r"\s+", "", t)
            area = float(pg.width * pg.height) or 1.0
            img_area = sum(abs((im["x1"] - im["x0"]) * (im["bottom"] - im["top"])) for im in pg.images)
            # 로고 같은 작은 이미지는 무시하고, 페이지의 20% 이상을 차지하는 그림만 본다
            if len(body) < min_chars or img_area / area >= 0.2:
                low.append(i)
    doc = pdfium.PdfDocument(path)
    targets = range(1, len(doc) + 1) if render_all else low
    for i in targets:
        doc[i - 1].render(scale=1.6).to_pil().save(os.path.join(out, f"{base}_p{i:02d}.png"))
    summary.append(f"[PDF] {os.path.basename(path)}: {n}쪽, 이미지로 확인할 페이지 {list(targets)}")


def do_xlsx(path, out, summary):
    import openpyxl
    base = safe(os.path.splitext(os.path.basename(path))[0])
    wb = openpyxl.load_workbook(path, data_only=True)
    lines = [f"[XLSX] {os.path.basename(path)}"]
    for ws in wb.worksheets:
        name = safe(ws.title)
        cnt = 0
        with open(os.path.join(out, f"{base}_{name}.txt"), "w", encoding="utf-8") as f:
            f.write(f"##### {ws.title} ({ws.max_row}x{ws.max_column})\n")
            for r_i, row in enumerate(ws.iter_rows(values_only=True), 1):
                if any(c is not None for c in row):
                    cnt += 1
                    f.write(f"{r_i}: " + " | ".join("" if c is None else str(c).replace("\n", " / ")
                                                     for c in row) + "\n")
        imgs = getattr(ws, "_images", [])
        for k, im in enumerate(imgs, 1):
            try:
                with open(os.path.join(out, f"{base}_{name}_img{k}.png"), "wb") as g:
                    g.write(im._data())
            except Exception as e:  # noqa: BLE001
                lines.append(f"  - {ws.title}: 이미지 {k} 추출 실패 ({e})")
        lines.append(f"  - {ws.title}: 값 있는 행 {cnt}, 크기 {ws.max_row}x{ws.max_column}, 이미지 {len(imgs)}")
    summary.extend(lines)


def do_docx(path, out, summary):
    import docx
    base = safe(os.path.splitext(os.path.basename(path))[0])
    d = docx.Document(path)
    with open(os.path.join(out, f"{base}.txt"), "w", encoding="utf-8") as f:
        for p in d.paragraphs:
            if p.text.strip():
                f.write(p.text + "\n")
        for ti, t in enumerate(d.tables, 1):
            f.write(f"\n[표 {ti}]\n")
            for row in t.rows:
                f.write(" | ".join(c.text.replace("\n", " / ") for c in row.cells) + "\n")
    summary.append(f"[DOCX] {os.path.basename(path)}: 문단 {len(d.paragraphs)}, 표 {len(d.tables)}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("out")
    ap.add_argument("--render-all", action="store_true")
    ap.add_argument("--min-chars", type=int, default=150)
    a = ap.parse_args()
    os.makedirs(a.out, exist_ok=True)
    files = []
    if os.path.isdir(a.src):
        for fn in sorted(os.listdir(a.src)):
            if not fn.startswith("~$"):
                files.append(os.path.join(a.src, fn))
    else:
        files.append(a.src)
    summary = []
    for p in files:
        ext = os.path.splitext(p)[1].lower()
        try:
            if ext == ".pdf":
                do_pdf(p, a.out, a.render_all, a.min_chars, summary)
            elif ext in (".xlsx", ".xlsm"):
                do_xlsx(p, a.out, summary)
            elif ext == ".docx":
                do_docx(p, a.out, summary)
            elif ext in (".hwp", ".hwpx"):
                summary.append(f"[HWP] {os.path.basename(p)}: 직접 읽을 수 없음 → PDF 변환 요청")
        except Exception as e:  # noqa: BLE001
            summary.append(f"[오류] {os.path.basename(p)}: {e}")
    with open(os.path.join(a.out, "_summary.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(summary) + "\n")
    print("\n".join(summary))


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
