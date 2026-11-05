"""PPT 자기 검증: 완료 조건을 숫자로 검사한다. 통과할 때까지 고치고 다시 돌린다.

사용:
  python verify_pptx.py <결과.pptx> [--base <원본.pptx>] [--plan <theme plan.json>] [--slides 3,5-7]
                        [--redesigned 1,3] [--json out.json]

--base를 주면(기존 PPT를 고친 작업) 문제를 "새로 생긴 것"과 "원본에도 있던 것"으로 나눈다.
완료 조건은 새로 생긴 것 0개다. 원본에도 있던 것은 보고만 한다(고치라는 요청이 없으면 손대지 않는다).
--redesigned: 일부러 다시 디자인한 장 — 원본과 줄 수 비교(줄늘어남)를 하지 않는다.

검사 항목(완료 조건 — 모두 0이어야 통과)
  넘침      글자가 보이는 상자(채우기·테두리 있음)나 표 칸을 넘친다
  이탈      도형이 장표 밖으로 0.05인치 넘게 나간다
  겹침      글자가 있는 상자끼리 10% 넘게 겹친다
  작은글자  8pt 미만 글자(원본에서 이미 그랬던 것은 제외)
  빽빽한표  표 행 위아래 여백(행 높이 − 글자 높이)이 0.14인치 미만 — 한 구역만 빽빽하면 억지로 끼워 넣은 장표로 보인다
  두줄라벨  흐름 도형(셰브런·오각 화살표) 안 라벨이 두 줄 이상 — PowerPoint 화면에서 칸을 넘친다
  튀는색    한 요소의 색이 장표 시각 비중을 독차지하거나(50% 이상), 강한 색 면적이 넓거나(30% 이상),
            강조색 계열이 4개 이상 — 다른 영역이 눈에 들어오지 않는다
  줄늘어남  --base: 원본보다 줄 수·높이가 늘어난 글상자
  미적용    --plan: 바꾸기로 한 색이 남아 있음 / 전면 배경 그림이 원본과 같음
정보(실패 아님): 글꼴 파일이 없어 넘침을 잴 수 없는 상자 수 — "검증 못 함"으로 보고한다.
확인(실패 아님): 크기혼용 — 한 글상자 안에 글자 크기가 섞임. 의도한 강조(핵심 문구·제목)가 아니면 하나로 맞춘다.
"""
import argparse
import hashlib
import json
import os
import re
import sys
from collections import defaultdict

from lxml import etree
from pptx import Presentation
from pptx.oxml.ns import qn

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from reskin import (Fonts, background_images, compare, hsl, hue_gap, lab, measure_deck,  # noqa: E402
                    para_runs, simulate, text_bodies, theme_fonts)
from pptx.opc.constants import RELATIONSHIP_TYPE as RT  # noqa: E402

EMU_IN = 914400
VIVID_C = 45          # Lab 채도(C*) 이 이상이면 '강한 색'
DOMINANT_SHARE = 0.5  # 한 요소가 시각 비중의 50% 이상
LOUD_AREA = 0.30      # 강한 색 면적이 장표의 30% 이상
MAX_ACCENT_HUES = 3   # 강조색 계열 3개까지
MIN_SZ = 800          # 글자 하한 8pt(1/100pt 단위)
ROW_PAD_IN = 0.14     # 표 행 위아래 여백 최소(인치)
FLOW_SHAPES = {"chevron", "homePlate"}
BULLET_GLYPHS = {"▪", "•", "◦", "·", "-", "※"}


def parse_slides(spec, n):
    if not spec:
        return set(range(1, n + 1))
    out = set()
    for part in spec.split(","):
        a, _, b = part.partition("-")
        out |= set(range(int(a), int(b or a) + 1))
    return out


# ------------------------------------------------------------------ 색 읽기
def scheme_map(slide):
    part = slide.slide_layout.slide_master.part.part_related_by(RT.THEME)
    scheme = etree.fromstring(part.blob).find(".//" + qn("a:clrScheme"))
    out = {}
    for el in scheme if scheme is not None else []:
        v = el.find(qn("a:srgbClr"))
        s = el.find(qn("a:sysClr"))
        out[etree.QName(el).localname] = (v.get("val") if v is not None else (s.get("lastClr") if s is not None else None))
    for a, b in (("tx1", "dk1"), ("bg1", "lt1"), ("tx2", "dk2"), ("bg2", "lt2")):
        out[a] = out.get(b)
    return out


def color_of(el, schemes):
    """srgbClr/schemeClr(+lumMod/lumOff) → 'RRGGBB'"""
    if el is None:
        return None
    c = el.find(qn("a:srgbClr"))
    if c is None:
        c = el.find(qn("a:schemeClr"))
        if c is None:
            return None
        base = schemes.get(c.get("val"))
    else:
        base = c.get("val")
    if not base:
        return None
    h, s, l = hsl(base)
    mod, off = c.find(qn("a:lumMod")), c.find(qn("a:lumOff"))
    if mod is not None:
        l *= int(mod.get("val")) / 100000
    if off is not None:
        l += int(off.get("val")) / 100000
    import colorsys
    r, g, b = colorsys.hls_to_rgb(h / 360, min(max(l, 0), 1), s)
    return "".join(f"{round(v * 255):02X}" for v in (r, g, b))


def fill_of(sp, schemes):
    pr = sp.find(qn("p:spPr"))
    if pr is None:
        return None
    if pr.find(qn("a:noFill")) is not None:
        return None
    sf = pr.find(qn("a:solidFill"))
    if sf is not None:
        return color_of(sf, schemes)
    gf = pr.find(qn("a:gradFill"))
    if gf is not None:
        gs = gf.find(".//" + qn("a:gs"))
        return color_of(gs, schemes)
    st = sp.find(qn("p:style"))   # 테마 스타일 채우기
    if st is not None and st.find(qn("a:fillRef")) is not None and int(st.find(qn("a:fillRef")).get("idx", 0)) > 0:
        return color_of(st.find(qn("a:fillRef")), schemes)
    return None


def has_line(sp):
    ln = sp.find(qn("p:spPr") + "/" + qn("a:ln"))
    if ln is not None:
        return ln.find(qn("a:noFill")) is None
    st = sp.find(qn("p:style"))
    return st is not None and st.find(qn("a:lnRef")) is not None and int(st.find(qn("a:lnRef")).get("idx", 0)) > 0


def xfrm(el):
    x = el.find(".//" + qn("a:xfrm"))
    if x is None or x.find(qn("a:off")) is None:
        return None
    o, e = x.find(qn("a:off")), x.find(qn("a:ext"))
    return int(o.get("x")), int(o.get("y")), int(e.get("cx")), int(e.get("cy"))


def top_level_shapes(slide):
    """그룹 안 도형은 그룹 좌표를 장표 좌표로 바꿔서 돌려준다."""
    out = []

    def walk(parent, tf):
        for el in parent:
            tag = etree.QName(el).localname
            if tag == "grpSp":
                gx = el.find(qn("p:grpSpPr") + "/" + qn("a:xfrm"))
                if gx is None:
                    walk(el, tf)
                    continue
                o, e = gx.find(qn("a:off")), gx.find(qn("a:ext"))
                co, ce = gx.find(qn("a:chOff")), gx.find(qn("a:chExt"))
                sx = int(e.get("cx")) / max(int(ce.get("cx")), 1)
                sy = int(e.get("cy")) / max(int(ce.get("cy")), 1)
                ax, ay, asx, asy = tf
                ntf = (ax + asx * (int(o.get("x")) - int(co.get("x")) * sx), ay + asy * (int(o.get("y")) - int(co.get("y")) * sy),
                       asx * sx, asy * sy)
                walk(el, ntf)
            elif tag in ("sp", "pic", "graphicFrame", "cxnSp"):
                b = xfrm(el)
                if b:
                    ax, ay, asx, asy = tf
                    out.append((el, (ax + b[0] * asx, ay + b[1] * asy, b[2] * asx, b[3] * asy)))
    walk(slide.shapes._spTree, (0, 0, 1, 1))
    return out


def rotated(el, box):
    """90°/270° 회전한 도형은 화면에 보이는 영역으로 바꾼다."""
    x = el.find(".//" + qn("a:xfrm"))
    rot = int(x.get("rot", 0)) // 60000 % 180 if x is not None else 0
    if 45 <= rot < 135:
        cx, cy = box[0] + box[2] / 2, box[1] + box[3] / 2
        return cx - box[3] / 2, cy - box[2] / 2, box[3], box[2]
    return box


def text_region(el, box, fonts, tf):
    """글자가 실제로 차지하는 영역(정렬·세로 맞춤 반영). 잴 수 없으면 None"""
    body = el.find(qn("p:txBody"))
    bp = body.find(qn("a:bodyPr"))
    li, ri = int(bp.get("lIns", 91440)), int(bp.get("rIns", 91440))
    ti, bi = int(bp.get("tIns", 45720)), int(bp.get("bIns", 45720))
    x, y, w, h = box
    avail = (w - li - ri) / 12700
    res = simulate(body, avail, bp.get("wrap") != "none", fonts, tf)
    if res is None:
        return None
    widest = 0.0
    for p in body.findall(qn("a:p")):
        line = 0.0
        for r, t, sz, face, b in para_runs(body, p, tf):
            m = fonts.get(face, b) if t != "\n" else None
            if m:
                line += m[0](t) * sz / 10000
        widest = max(widest, line)
    tw = min(widest, avail) * 12700
    th = res[1] * 12700
    ppr = body.find(qn("a:p") + "/" + qn("a:pPr"))
    algn = ppr.get("algn", "l") if ppr is not None else "l"
    tx = x + li + {"ctr": (w - li - ri - tw) / 2, "r": (w - li - ri - tw)}.get(algn, 0)
    anchor = bp.get("anchor", "t")
    ty = y + ti + {"ctr": (h - ti - bi - th) / 2, "b": (h - ti - bi - th)}.get(anchor, 0)
    return tx, ty, tw, th


def name_of(el):
    c = el.find(".//" + qn("p:cNvPr"))
    return c.get("name") if c is not None else "?"


def text_of(el):
    return "".join(el.itertext()).strip()


# ------------------------------------------------------------------ 검사
def check_color(slide, area, bg="FFFFFF"):
    """튀는 색: 한 요소 독차지 / 강한 색 면적 / 강조색 계열 수"""
    schemes = scheme_map(slide)
    items = []
    for el, (x, y, w, h) in top_level_shapes(slide):
        if etree.QName(el).localname != "sp" or w * h <= 0:
            continue
        frac = w * h / area
        c = fill_of(el, schemes)
        if c is None or frac >= 0.9:      # 전면 배경은 제외
            continue
        L, a, b = lab(c)
        chroma = (a * a + b * b) ** 0.5
        contrast = abs(L - lab(bg)[0]) / 100
        items.append((name_of(el), c, frac, chroma, contrast, frac * min(chroma / 80, 1) * contrast))
    issues = []
    total = sum(i[5] for i in items) or 1
    others = sum(1 for el, _ in top_level_shapes(slide) if text_of(el))
    for n, c, frac, chroma, contrast, s in items:
        if chroma >= VIVID_C and s / total >= DOMINANT_SHARE and frac >= 0.08 and others >= 4:
            issues.append(f"'{n}' 색 {c}이 장표 시각 비중의 {s / total:.0%}(면적 {frac:.0%}) — 다른 영역이 눈에 안 들어옴")
    loud = sum(i[2] for i in items if i[3] >= VIVID_C and i[4] >= 0.3)
    if loud >= LOUD_AREA:
        issues.append(f"강한 색 면적이 장표의 {loud:.0%} (기준 {LOUD_AREA:.0%} 미만)")
    hues = []
    for _, c, frac, chroma, _, _ in sorted(items, key=lambda i: -i[2]):
        if chroma >= VIVID_C and frac >= 0.002:
            h = hsl(c)[0]
            if all(hue_gap(h, x) >= 30 for x in hues):
                hues.append(h)
    if len(hues) > MAX_ACCENT_HUES:
        issues.append(f"강조색 계열 {len(hues)}개 (기준 {MAX_ACCENT_HUES}개 이하)")
    return issues


def plain(body):
    return " ".join(x.text or "" for x in body.iter(qn("a:t"))).strip()


def check_slide(slide, prs, fonts, base_small):
    W, H = prs.slide_width, prs.slide_height
    tf = theme_fonts(slide.part)
    out = defaultdict(list)
    unmeasured = 0
    shapes = [(el, rotated(el, b)) for el, b in top_level_shapes(slide)]
    regions = {}
    for el, b in shapes:
        if etree.QName(el).localname == "sp" and text_of(el) and el.find(qn("p:txBody")) is not None:
            regions[id(el)] = text_region(el, b, fonts, tf)
    tol = 0.05 * EMU_IN
    for el, b in shapes:
        x, y, w, h = regions.get(id(el)) or b    # 글상자는 글자 영역 기준
        if x < -tol or y < -tol or x + w > W + tol or y + h > H + tol:
            out["이탈"].append(f"'{name_of(el)}' ({x / EMU_IN:.2f},{y / EMU_IN:.2f},{w / EMU_IN:.2f}x{h / EMU_IN:.2f})")
    texts = [(el, regions[id(el)]) for el, _ in shapes if regions.get(id(el))]
    for i, (a, ba) in enumerate(texts):
        for b, bb in texts[i + 1:]:
            ix = max(0, min(ba[0] + ba[2], bb[0] + bb[2]) - max(ba[0], bb[0]))
            iy = max(0, min(ba[1] + ba[3], bb[1] + bb[3]) - max(ba[1], bb[1]))
            small = min(ba[2] * ba[3], bb[2] * bb[3]) or 1
            if ix * iy / small > 0.10:
                out["겹침"].append(f"'{name_of(a)}' ↔ '{name_of(b)}' {ix * iy / small:.0%}")
    lines_of = {}
    sps = {el.find(qn("p:txBody")): el for el, _ in shapes
           if etree.QName(el).localname == "sp" and el.find(qn("p:txBody")) is not None}   # 요소 자체를 키로(id는 바뀔 수 있음)
    for body, width, wrap, box_h in text_bodies(slide):
        if not "".join(body.itertext()).strip():
            continue
        for p in body.findall(qn("a:p")):
            for r, t, sz, face, bold in para_runs(body, p, tf):
                if t.strip() and sz and sz < MIN_SZ and (t, sz) not in base_small:
                    out["작은글자"].append(f"{sz / 100:g}pt '{t[:15]}'")
        sizes = {sz for p in body.findall(qn("a:p")) for r, t, sz, face, bold in para_runs(body, p, tf)
                 if sz and t.strip() and t.strip() not in BULLET_GLYPHS}
        if len(sizes) > 1:
            out["크기혼용"].append(f"{'/'.join(f'{x / 100:g}' for x in sorted(sizes))}pt '{''.join(x.text or '' for x in body.iter(qn('a:t'))).strip()[:18]}'")
        res = simulate(body, width, wrap, fonts, tf)
        if res is None:
            unmeasured += 1
            continue
        lines_of[body] = res[0]
        sp = sps.get(body)
        tc = body.getparent()
        if (sp is None and etree.QName(tc).localname == "tc" and int(tc.get("rowSpan", "1")) == 1   # 표 칸(병합 칸 제외)
                and tc.getparent().getprevious() is not None and etree.QName(tc.getparent().getprevious()).localname == "tr"):   # 머리행 제외
            pad = (box_h - res[1]) / 72
            if pad < ROW_PAD_IN:
                out["빽빽한표"].append(f"여백 {pad:.2f}in < {ROW_PAD_IN} '{''.join(x.text or '' for x in body.iter(qn('a:t'))).strip()[:15]}'")
        visible = sp is None or fill_of(sp, scheme_map(slide)) is not None or has_line(sp)   # 표 칸이거나 보이는 상자
        bp = body.find(qn("a:bodyPr"))
        auto = bp is not None and bp.find(qn("a:normAutofit")) is not None
        if sp is not None:   # 안쪽 여백까지 합친 도형 높이와 비교(여백을 조금 넘어도 도형 안이면 보기에는 괜찮다)
            box_h += (int(bp.get("tIns", 45720)) + int(bp.get("bIns", 45720))) / 12700
        if visible and not auto and res[1] > box_h * 1.05 + 2:
            label = name_of(sp) if sp is not None else "표 칸"
            out["넘침"].append(f"'{label}' 글자 높이 {res[1]:.0f}pt > 상자 {box_h:.0f}pt '{''.join(body.itertext()).strip()[:15]}'")
    flows = []
    for el, b in shapes:
        g = el.find(".//" + qn("a:prstGeom"))
        if g is not None and g.get("prst") in FLOW_SHAPES:
            flows.append((el, b))
    for fel, (fx, fy, fw, fh) in flows:
        own = fel.find(qn("p:txBody"))
        if own is not None and lines_of.get(own, 1) > 1:
            out["두줄라벨"].append(f"'{name_of(fel)}' {lines_of[own]}줄 '{plain(own)[:15]}'")
        for el, (x, y, w, h) in shapes:
            body = el.find(qn("p:txBody"))
            if el is fel or body is None or lines_of.get(body, 1) < 2:
                continue
            cx, cy = x + w / 2, y + h / 2
            if fx <= cx <= fx + fw and fy <= cy <= fy + fh and h <= fh + 0.1 * EMU_IN:
                out["두줄라벨"].append(f"'{name_of(el)}' {lines_of[body]}줄 '{plain(body)[:15]}'")
    out["튀는색"] = check_color(slide, W * H)
    return out, unmeasured


def small_runs(prs):
    s = set()
    for slide in prs.slides:
        tf = theme_fonts(slide.part)
        for body, *_ in text_bodies(slide):
            for p in body.findall(qn("a:p")):
                for r, t, sz, face, b in para_runs(body, p, tf):
                    if sz and sz < MIN_SZ:
                        s.add((t, sz))
    return s


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--base")
    ap.add_argument("--plan")
    ap.add_argument("--slides")
    ap.add_argument("--json")
    ap.add_argument("--redesigned")
    a = ap.parse_args()
    sys.stdout.reconfigure(encoding="utf-8")
    prs = Presentation(a.pptx)
    fonts = Fonts()
    base = Presentation(a.base) if a.base else None
    base_small = small_runs(base) if base else set()
    sel = parse_slides(a.slides, len(prs.slides))
    redesigned = parse_slides(a.redesigned, len(prs.slides)) if a.redesigned else set()
    result, unmeasured, existing, notes = {}, 0, {}, {}
    key = lambda k, msg: k + re.sub(r"[\d.,]+%?|pt|\(.*?\)", "", msg)
    for i, s in enumerate(prs.slides, 1):
        if i in sel:
            r, u = check_slide(s, prs, fonts, base_small)
            unmeasured += u
            r = {k: v for k, v in r.items() if v}
            if r.get("크기혼용"):
                notes.setdefault(i, {})["크기혼용"] = r.pop("크기혼용")
            if base and i <= len(base.slides):      # 원본에도 있던 문제는 따로
                old, _ = check_slide(base.slides[i - 1], base, fonts, set())
                old_keys = {key(k, m) for k, v in old.items() for m in v}
                for k in list(r):
                    keep = [m for m in r[k] if key(k, m) not in old_keys]
                    gone = [m for m in r[k] if key(k, m) in old_keys]
                    if gone:
                        existing.setdefault(i, {})[k] = gone
                    if keep:
                        r[k] = keep
                    else:
                        del r[k]
            result[i] = r

    if base:
        for (si, bi), x, y in compare(measure_deck(base, fonts), measure_deck(prs, fonts)):
            if si in sel and si not in redesigned:
                result.setdefault(si, {}).setdefault("줄늘어남", []).append(f"글상자 {bi}: 줄 {x[0]}→{y[0]}, 높이 {x[1]:.0f}→{y[1]:.0f}pt")
    if a.plan:
        plan = json.load(open(a.plan, encoding="utf-8"))
        gone = {c["from"].upper() for c in plan["colors"] if c["from"].upper() != c["to"].upper()}
        for i, s in enumerate(prs.slides, 1):
            left = {x.get("val").upper() for x in s._element.iter(qn("a:srgbClr"))} & gone
            if left and i in sel:
                result.setdefault(i, {}).setdefault("미적용", []).append(f"바꾸기로 한 색이 남음: {', '.join(sorted(left))}")
        if base and plan.get("images"):
            before = {str(p.partname): hashlib.md5(p.blob).hexdigest() for p in background_images(base)}
            for p in background_images(prs):
                if before.get(str(p.partname)) == hashlib.md5(p.blob).hexdigest():
                    result.setdefault(0, {}).setdefault("미적용", []).append(f"전면 배경 그림이 원본과 같음: {p.partname}")

    result = {k: v for k, v in result.items() if v}
    total = defaultdict(int)
    for v in result.values():
        for k, items in v.items():
            total[k] += len(items)
    print(f"검사: {a.pptx}  ({len(sel)}장)")
    label = "새로 생긴 문제 " if base else ""
    print("완료 조건: " + ("통과" if not total else f"미통과 — {label}" + ", ".join(f"{k} {n}" for k, n in total.items())))
    if existing:
        ex = defaultdict(int)
        for v in existing.values():
            for k, items in v.items():
                ex[k] += len(items)
        print("원본에도 있던 문제(보고만): " + ", ".join(f"{k} {n}" for k, n in ex.items())
              + f" — {', '.join(map(str, sorted(existing)))}장")
    if notes:
        print("확인(실패 아님) — 의도한 강조가 아니면 맞춘다: " + ", ".join(
            f"{i}장 크기혼용 {len(v['크기혼용'])}" for i, v in sorted(notes.items())))
        for i, v in sorted(notes.items()):
            for it in v["크기혼용"][:4]:
                print(f"  [{i}장] {it}")
    if unmeasured:
        print(f"검증 못 함: 글꼴 파일이 없어 넘침을 못 잰 글상자 {unmeasured}개")
    for si in sorted(result):
        print(f"\n[{si}장]" if si else "\n[파일 전체]")
        for k, items in result[si].items():
            for it in items[:8]:
                print(f"  {k}: {it}")
            if len(items) > 8:
                print(f"  {k}: … 외 {len(items) - 8}건")
    if a.json:
        json.dump({"total": total, "unmeasured": unmeasured, "slides": result}, open(a.json, "w", encoding="utf-8"),
                  ensure_ascii=False, indent=1)
    sys.exit(1 if total else 0)


if __name__ == "__main__":
    main()
