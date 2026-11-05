"""맥: PowerPoint로 장표를 PNG로 렌더링한다(윈도우는 render_pptx.ps1).

사용: python render_pptx_mac.py <pptx> <출력폴더> [--slides 3,5-7] [--dpi 80]

- 원본은 열지 않는다. 확인할 장만 남긴 사본을 PowerPoint 샌드박스 폴더에 고유 이름(claude_xxxx.pptx)으로 만들어 연다.
- 저장·닫기는 반드시 그 이름으로만 지정한다. 사용자가 열어 둔 다른 문서는 건드리지 않는다.
- PowerPoint를 사용자가 쓰고 있으면 실행 전에 사용자에게 묻는다(스킬 규칙).
- Claude 실행 환경에서는 샌드박스 밖 실행이 필요할 수 있다.
"""
import argparse
import glob
import os
import shutil
import subprocess
import sys
import time
import uuid

from pptx import Presentation

BOX = os.path.expanduser("~/Library/Containers/com.microsoft.Powerpoint/Data/tmp_claude")


def keep_only(src, dst, slides):
    prs = Presentation(src)
    ids = prs.slides._sldIdLst
    for i, sid in reversed(list(enumerate(list(ids), 1))):
        if slides and i not in slides:
            prs.part.drop_rel(sid.rId)
            ids.remove(sid)
    prs.save(dst)


def parse(spec):
    out = set()
    for part in (spec or "").split(","):
        if part:
            a, _, b = part.partition("-")
            out |= set(range(int(a), int(b or a) + 1))
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("out")
    ap.add_argument("--slides")
    ap.add_argument("--dpi", type=int, default=80)
    a = ap.parse_args()
    slides = parse(a.slides)
    os.makedirs(BOX, exist_ok=True)
    os.makedirs(a.out, exist_ok=True)
    name = f"claude_{uuid.uuid4().hex[:8]}.pptx"
    copy, pdf = os.path.join(BOX, name), os.path.join(BOX, name[:-5] + ".pdf")
    keep_only(a.pptx, copy, slides)
    script = f'''
tell application "Microsoft PowerPoint"
    open (POSIX file "{copy}")
    repeat 60 times
        if (name of every presentation) contains "{name}" then exit repeat
        delay 0.5
    end repeat
    save presentation "{name}" in (POSIX file "{pdf}") as save as PDF
    close presentation "{name}" saving no
end tell'''
    try:
        subprocess.run(["osascript", "-e", script], check=True, timeout=300)
        for _ in range(60):
            if os.path.exists(pdf) and os.path.getsize(pdf) > 0:
                break
            time.sleep(0.5)
        order = sorted(slides) if slides else None
        prefix = os.path.join(a.out, "slide")
        subprocess.run(["pdftoppm", "-png", "-r", str(a.dpi), pdf, prefix], check=True)
        pages = sorted(glob.glob(prefix + "-*.png"))
        for k, p in enumerate(pages):   # 원래 장 번호로 이름을 바꾼다
            no = order[k] if order else k + 1
            os.replace(p, os.path.join(a.out, f"slide_{no:03d}.png"))
        print(f"렌더링 {len(pages)}장 → {a.out}")
    finally:
        for f in (copy, pdf):
            if os.path.exists(f):
                os.remove(f)
        if os.path.isdir(BOX) and not os.listdir(BOX):
            shutil.rmtree(BOX, ignore_errors=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
