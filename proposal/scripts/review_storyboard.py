"""제안서 원고(스토리보드) 자동 검수.

사용: python review_storyboard.py <제안작업 폴더> [--owner B]
  --owner: 담당별 검수. 원고의 "담당" 필드와 추적대상.csv의 "담당" 열이 그 값인 것만 본다.
           결과 파일 이름에 _담당 값이 붙는다(예: 검수보고서_B.md).

읽는 것
  04_원고/*.md       — "### [장-절-번호] 제목" 블록과 "- 필드: 값" 줄
  03_전략/추적대상.csv — 키,출처,내용,필수 (없으면 추적표 생략)
만드는 것
  06_검수/검수보고서.md, 06_검수/추적표.xlsx, 06_검수/완성도점검표.md
"""
import csv
import glob
import os
import re
import sys
from collections import defaultdict

AMBIG = [
    (r"할\s*수\s*있(다|음|습니다|으며|는)", "~할 수 있다 → ~한다"),
    (r"될\s*수\s*있(다|음|습니다)", "~될 수 있다 → ~된다"),
    (r"가능(하다|함|합니다|할 것)", "~가능하다 → ~한다/~를 제공한다"),
    (r"사용\s*가능", "사용 가능 → 사용한다"),
    (r"공급될\s*수", "공급될 수 → 공급한다"),
    (r"고려(하겠|하고자|하고 싶|할 예정|토록)", "고려하겠다 → 적용한다(대상·시점)"),
    (r"검토(하겠|할 예정|예정)", "검토 예정 → 수행한다 또는 권고안 명시"),
    (r"노력(하겠|함|합니다)", "노력하겠다 → 구체적 방법·횟수·기준"),
    (r"최선을\s*다", "최선을 다하겠다 → 구체적 방법"),
    (r"최대한|가급적|가능한\s*한", "최대한/가급적 → 수치·기준"),
    (r"필요\s*시", "필요 시 → 조건 명시(○○ 발생 시)"),
]
TAGS = {"근거 필요": r"\[근거 필요[^\]]*\]", "자산 필요": r"\[자산 필요[^\]]*\]", "가정": r"\[가정[^\]]*\]"}
NUM = re.compile(r"(\d+(?:\.\d+)?)\s*(개월|M/M|MM|명|주|영업일|%|건|본|회)")
HEAD = re.compile(r"^###\s*\[([^\]]+)\]\s*(.*)$")
FIELD = re.compile(r"^-\s*(헤드메시지|본문|슬로건|키워드|도식|요약 바|근거|대응|담당|상태)\s*:\s*(.*)$")


def parse(path):
    slides, cur, field = [], None, None
    for ln in open(path, encoding="utf-8"):
        ln = ln.rstrip("\n")
        m = HEAD.match(ln)
        if m:
            cur = {"id": m.group(1).strip(), "title": m.group(2).strip(), "file": os.path.basename(path),
                   "헤드메시지": "", "본문": "", "슬로건": "", "키워드": "", "도식": "", "요약 바": "", "근거": "", "대응": "", "담당": "", "상태": "",
                   "raw": ""}
            slides.append(cur)
            field = None
            continue
        if cur is None:
            continue
        cur["raw"] += ln + "\n"
        f = FIELD.match(ln)
        if f:
            field = f.group(1)
            cur[field] = f.group(2).strip()
        elif field == "본문" and ln.strip().startswith("-"):
            cur["본문"] += "\n" + ln.strip()[1:].strip()
    return slides


def diagram_type(v):
    return v.split("|")[0].strip() if v else ""


def main(root, owner=None):
    files = sorted(glob.glob(os.path.join(root, "04_원고", "*.md")))
    slides = [s for f in files for s in parse(f)]
    if owner:
        slides = [s for s in slides if s["담당"].strip() == owner]
    sfx = f"_{owner}" if owner else ""
    out_dir = os.path.join(root, "06_검수")
    os.makedirs(out_dir, exist_ok=True)
    title = f"# 검수보고서 — 담당 {owner}" if owner else "# 검수보고서"
    R = [title, "", f"- 원고 파일 {len(files)}개, 장표 {len(slides)}장", ""]

    # 추적표
    trace_rows, missing = [], []
    tpath = os.path.join(root, "03_전략", "추적대상.csv")
    if os.path.exists(tpath):
        with open(tpath, encoding="utf-8-sig") as f:
            targets = list(csv.DictReader(f))
        if owner:
            targets = [t for t in targets if t.get("담당", "").strip() == owner]
        cover = defaultdict(list)
        for s in slides:
            for k in re.split(r"[,\s]+", s["대응"]):
                if k:
                    cover[k.strip()].append(s["id"])
        for t in targets:
            key = t.get("키", "").strip()
            ids = cover.get(key, [])
            trace_rows.append([key, t.get("출처", ""), t.get("내용", ""), t.get("필수", ""), ", ".join(ids)])
            if not ids and t.get("필수", "").strip().upper() == "Y":
                missing.append(t)
        unknown = [] if owner else sorted(k for k in cover if k not in {t.get("키", "").strip() for t in targets})
        R += ["## 1. 추적표", "", f"- 추적 대상 {len(targets)}건, 대응 장표 없음(필수) **{len(missing)}건**", ""]
        if missing:
            R += ["| 키 | 출처 | 내용 |", "|---|---|---|"] + [f"| {m['키']} | {m.get('출처','')} | {m.get('내용','')} |" for m in missing] + [""]
        if unknown:
            R += [f"- 추적대상에 없는 대응 키(오타 의심): {', '.join(unknown)}", ""]
    else:
        R += ["## 1. 추적표", "", "- `03_전략/추적대상.csv`가 없어 생략했다.", ""]

    # 모호 표현
    amb = []
    for s in slides:
        text = s["헤드메시지"] + "\n" + s["본문"]
        for pat, fix in AMBIG:
            for m in re.finditer(pat, text):
                a, b = max(0, m.start() - 20), min(len(text), m.end() + 15)
                amb.append((s["file"], s["id"], text[a:b].replace("\n", " "), fix))
    R += ["## 2. 모호 표현", "", f"- {len(amb)}건", ""]
    if amb:
        R += ["| 파일 | 장표 | 문맥 | 바꾸는 방향 |", "|---|---|---|---|"] + [f"| {a} | {b} | …{c}… | {d} |" for a, b, c, d in amb] + [""]

    # 태그
    R += ["## 3. 태그", ""]
    tag_by_slide = defaultdict(lambda: defaultdict(int))
    for s in slides:
        for name, pat in TAGS.items():
            for m in re.finditer(pat, s["raw"]):
                tag_by_slide[s["id"]][name] += 1
                R.append(f"- [{s['id']}] {m.group(0)}")
    if not any(tag_by_slide.values()):
        R.append("- 없음")
    R.append("")

    # 도식·필드
    R += ["## 4. 도식·필드 점검", ""]
    prev, run = None, 0
    for s in slides:
        typ = diagram_type(s["도식"])
        is_body = not re.search(r"표지|목차|간지", s["title"])
        if is_body and (not typ or typ in ("없음", "-")):
            R.append(f"- [{s['id']}] 도식 없음 — 텍스트 장표")
        for fld in ("헤드메시지", "근거", "대응"):
            if is_body and not s[fld]:
                R.append(f"- [{s['id']}] {fld} 비어 있음")
        run = run + 1 if typ and typ == prev else 1
        if run == 3:
            R.append(f"- [{s['id']}] 같은 도식 유형 '{typ}' 3장 연속")
        prev = typ
    R.append("")

    # 수치 대조
    R += ["## 5. 수치 대조 (사람이 확인)", "", "단위별로 장표마다 등장한 값이다. 같은 대상인데 값이 다르면 고친다.", ""]
    by_unit = defaultdict(lambda: defaultdict(set))
    for s in slides:
        for v, u in NUM.findall(s["헤드메시지"] + " " + s["본문"]):
            by_unit[u][v].add(s["id"])
    for u, vals in by_unit.items():
        if len(vals) > 1:
            R.append(f"- **{u}**: " + "; ".join(f"{v}{u} ({', '.join(sorted(ids))})" for v, ids in sorted(vals.items())))
    R.append("")

    # 완성도
    chap = defaultdict(lambda: defaultdict(int))
    for s in slides:
        c = s["id"].split("-")[0]
        chap[c]["장표"] += 1
        chap[c][s["상태"] or "미기재"] += 1
        typ = diagram_type(s["도식"])
        if not typ or typ in ("없음", "-"):
            chap[c]["도식 없음"] += 1
        for name in TAGS:
            chap[c][name] += tag_by_slide[s["id"]][name]
        chap[c]["모호 표현"] += sum(1 for a in amb if a[1] == s["id"])
    miss_by_ch = defaultdict(int)
    C = ["# 완성도점검표", "", "| 장 | 장표 | 초안 | 검토 | 완료 | 도식 없음 | 근거 필요 | 자산 필요 | 가정 | 모호 표현 |",
         "|---|---|---|---|---|---|---|---|---|---|"]
    for c in sorted(chap, key=lambda x: (len(x), x)):
        d = chap[c]
        C.append(f"| {c}장 | {d['장표']} | {d['초안']} | {d['검토']} | {d['완료']} | {d['도식 없음']} | "
                 f"{d['근거 필요']} | {d['자산 필요']} | {d['가정']} | {d['모호 표현']} |")
    C += ["", f"- 추적 누락(필수): {len(missing)}건", ""]
    open(os.path.join(out_dir, f"완성도점검표{sfx}.md"), "w", encoding="utf-8").write("\n".join(C))
    open(os.path.join(out_dir, f"검수보고서{sfx}.md"), "w", encoding="utf-8").write("\n".join(R))

    if trace_rows:
        try:
            import openpyxl
            from openpyxl.styles import Font, PatternFill
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "추적표"
            ws.append(["키", "출처", "내용", "필수", "대응 장표"])
            for c in ws[1]:
                c.font = Font(bold=True, color="FFFFFF")
                c.fill = PatternFill("solid", fgColor="14325C")
            for r in trace_rows:
                ws.append(r)
                if not r[4] and str(r[3]).upper() == "Y":
                    ws.cell(ws.max_row, 5).fill = PatternFill("solid", fgColor="FBE6E6")
                    ws.cell(ws.max_row, 5).value = "누락"
            for col, w in zip("ABCDE", (18, 22, 60, 6, 30)):
                ws.column_dimensions[col].width = w
            ws.freeze_panes = "A2"
            wb.save(os.path.join(out_dir, f"추적표{sfx}.xlsx"))
        except PermissionError:
            print("추적표.xlsx가 열려 있어 저장하지 못했다. 닫고 다시 실행한다.")

    print(f"장표 {len(slides)} | 추적 누락 {len(missing)} | 모호 표현 {len(amb)} | "
          f"태그 {sum(sum(v.values()) for v in tag_by_slide.values())}")
    print(f"→ {os.path.join(out_dir, f'검수보고서{sfx}.md')}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    own = None
    if "--owner" in sys.argv:
        i = sys.argv.index("--owner")
        own = sys.argv[i + 1] if i + 1 < len(sys.argv) else None
        args = [a for a in args if a != own]
    main(args[0] if args else "제안작업", own)
