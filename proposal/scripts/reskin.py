"""테마(스타일) 적용: 이미 만든 PPT에 템플릿의 색·글꼴 스타일을 입힌다.

구조·틀·배치·내용은 바꾸지 않는다(윈도우에 맥 테마를 입히듯 생김새만 바꾼다).
템플릿은 PPTX나 PDF 모두 된다. 판형이 달라도 된다. 원본은 고치지 않고 새 파일로 저장한다.

- 색: 도형·글자·표·차트·테마 색 + 표지·간지의 전면 배경 그림(밝기 기준으로 템플릿 팔레트에 맞춰 다시 칠함)
- 글꼴: 굵기가 같은 템플릿 글꼴로. 글자 폭·줄 높이가 달라지는 만큼 글자 크기·줄 간격을 보정해
  원래 줄 수·높이를 지킨다(레이아웃 보존). 적용 후 글상자마다 줄 수를 원본과 비교해 보고한다.

사용:
  python reskin.py analyze <템플릿.pptx|.pdf> <대상.pptx> <plan.json>   # 대응표 초안 + asis/tobe 출력
  python reskin.py apply   <대상.pptx> <plan.json> <결과.pptx>
  python reskin.py check   <원본.pptx> <결과.pptx>                       # 줄 수·높이 비교만
  python reskin.py brand   <템플릿.pptx|.pdf>                            # 새 PPT용 BRAND 값만 출력

plan.json (analyze가 초안을 만들고, 사람이 승인·수정한다)
  colors   [{from, to, count, note}]            from==to면 유지
  fonts    [{from, to, bold, scale, line, count}]  scale: 글자 크기 배율, line: 줄 간격 배율
  palette  배경 그림을 다시 칠할 때 쓰는 템플릿 색(어두운 색→밝은 색)
  images   true면 전면 배경 그림도 다시 칠한다
  brand    새 PPT(build_deck.py)에 넣을 BRAND 추천값
"""
import colorsys
import glob
import io
import json
import math
import os
import re
import sys
from collections import Counter

from lxml import etree
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn

FONT_TAGS = (qn("a:latin"), qn("a:ea"), qn("a:cs"))
RPR_TAGS = (qn("a:rPr"), qn("a:defRPr"), qn("a:endParaRPr"))
R_EMBED = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
EMU_PT = 12700
MIN_SZ = 600   # 글자 크기 하한 6pt
# PDF에 박힌 글꼴 이름(PostScript) → PowerPoint 글꼴 이름
PS_FONTS = {"MalgunGothic": "맑은 고딕", "NanumSquare": "나눔스퀘어", "NanumGothic": "나눔고딕",
            "NanumBarunGothic": "나눔바른고딕", "NanumMyeongjo": "나눔명조", "Dotum": "돋움",
            "Gulim": "굴림", "Batang": "바탕", "AppleSDGothicNeo": "Apple SD Gothic Neo",
            "ArialMT": "Arial"}
FONT_DIRS = ["~/Library/Fonts", "/Library/Fonts", "/System/Library/Fonts", "/System/Library/Fonts/Supplemental",
             "/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts",
             "C:/Windows/Fonts", "~/AppData/Local/Microsoft/Windows/Fonts"]


# ------------------------------------------------------------------ 색 계산
def hsl(hex_):
    r, g, b = (int(hex_[i:i + 2], 16) / 255 for i in (0, 2, 4))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    return h * 360, s, l


def hex_from_hls(h, l, s):
    r, g, b = colorsys.hls_to_rgb(h / 360, l, s)
    return "".join(f"{round(v * 255):02X}" for v in (r, g, b))


def lab(hex_):
    def lin(v):
        v /= 255
        return v / 12.92 if v <= 0.04045 else ((v + 0.055) / 1.055) ** 2.4
    r, g, b = (lin(int(hex_[i:i + 2], 16)) for i in (0, 2, 4))
    x, y, z = ((r * .4124 + g * .3576 + b * .1805) / .9505, r * .2126 + g * .7152 + b * .0722,
               (r * .0193 + g * .1192 + b * .9505) / 1.089)
    f = [t ** (1 / 3) if t > 0.008856 else 7.787 * t + 16 / 116 for t in (x, y, z)]
    return 116 * f[1] - 16, 500 * (f[0] - f[1]), 200 * (f[1] - f[2])


def delta_e(a, b):
    return sum((p - q) ** 2 for p, q in zip(lab(a), lab(b))) ** 0.5


def is_neutral(hex_):
    _, s, l = hsl(hex_)
    return s < 0.15 or l > 0.94 or l < 0.06


def is_warning(hex_):
    h, s, l = hsl(hex_)
    return (h < 15 or h > 345) and s > 0.5 and 0.3 < l < 0.7


def hue_gap(a, b):
    d = abs(a - b) % 360
    return min(d, 360 - d)


# ------------------------------------------------------------------ 글꼴 파일·치수
class Fonts:
    """글꼴 이름 → 글꼴 파일. 글자 폭과 줄 높이를 잰다(크기 100 기준)."""

    def __init__(self):
        from fontTools.ttLib import TTCollection, TTFont
        self.files, self.cache = {}, {}
        for d in FONT_DIRS:
            for p in glob.glob(os.path.expanduser(d) + "/*"):
                if not p.lower().endswith((".ttf", ".otf", ".ttc")):
                    continue
                try:
                    fonts = TTCollection(p).fonts if p.lower().endswith(".ttc") else [TTFont(p, lazy=True)]
                except Exception:
                    continue
                for idx, f in enumerate(fonts):
                    names = {}
                    for r in f["name"].names:
                        if r.nameID in (1, 2, 4, 16, 17):
                            try:
                                names.setdefault(r.nameID, set()).add(r.toUnicode())
                            except Exception:
                                pass
                    style = " ".join(names.get(2, set()) | names.get(17, set())).lower()
                    bold = "bold" in style or "굵게" in style
                    for n in names.get(4, set()):
                        self.files.setdefault((n, None), (p, idx))
                    for n in names.get(1, set()) | names.get(16, set()):
                        if bold or "regular" in style or "보통" in style or style.strip() in ("", "normal"):
                            self.files.setdefault((n, bold), (p, idx))

    def get(self, name, bold=False):
        """(폭 재는 함수, 줄 높이 배율) 또는 None"""
        key = (name, bold)
        if key not in self.cache:
            hit = (self.files.get((name, bold)) or self.files.get((name, None))
                   or self.files.get((name, not bold)))   # '페이퍼로지 7 Bold'처럼 이름에 굵기가 든 글꼴
            if hit is None:
                self.cache[key] = None
            else:
                from fontTools.ttLib import TTCollection, TTFont
                from PIL import ImageFont
                f = ImageFont.truetype(hit[0], 100, index=hit[1])
                tt = TTCollection(hit[0]).fonts[hit[1]] if hit[0].lower().endswith(".ttc") else TTFont(hit[0], lazy=True)
                cmap = set(tt.getBestCmap() or {})
                a, d = f.getmetrics()

                def measure(t, f=f, cmap=cmap):
                    # 글꼴에 없는 글자(▶·♡ 같은 기호)는 PowerPoint가 다른 글꼴로 대신 그린다 → 1em으로 셈
                    have = "".join(c for c in t if ord(c) in cmap or c.isspace())
                    return f.getlength(have) + 100 * (len(t) - len(have))
                self.cache[key] = (measure, (a + d) / 100)
        return self.cache[key]


# ------------------------------------------------------------------ 글자 서식 읽기
def theme_fonts(part):
    """마스터/레이아웃/장표 part → 테마 글꼴 {'+mn-ea': '…', …}"""
    p = part
    while p is not None and RT.THEME not in [r.reltype for r in p.rels.values()]:
        nxt = None
        for r in p.rels.values():
            if r.reltype in (RT.SLIDE_LAYOUT, RT.SLIDE_MASTER):
                nxt = r.target_part
        p = nxt
    if p is None:
        return {}
    root = etree.fromstring(p.part_related_by(RT.THEME).blob)
    out = {}
    for tag, key in ((qn("a:majorFont"), "mj"), (qn("a:minorFont"), "mn")):
        el = root.find(".//" + tag)
        for sub, k2 in ((qn("a:latin"), "lt"), (qn("a:ea"), "ea")):
            s = el.find(sub) if el is not None else None
            if s is not None and s.get("typeface"):
                out[f"+{key}-{k2}"] = s.get("typeface")
    return out


def lvl_def(body, p):
    ppr = p.find(qn("a:pPr"))
    lvl = int(ppr.get("lvl", "0")) if ppr is not None else 0
    ls = body.find(qn("a:lstStyle"))
    el = ls.find(qn(f"a:lvl{lvl + 1}pPr")) if ls is not None else None
    return el.find(qn("a:defRPr")) if el is not None else None


def run_props(body, p, rpr, tfonts):
    """(크기 1/100pt, 글꼴 이름, 굵게)"""
    d = lvl_def(body, p)
    sz = (rpr.get("sz") if rpr is not None else None) or (d.get("sz") if d is not None else None) or "1800"
    face = None
    for src in (rpr, d):
        if src is None or face:
            continue
        for t in (qn("a:ea"), qn("a:latin")):
            el = src.find(t)
            if el is not None and el.get("typeface"):
                face = el.get("typeface")
                break
    face = face or "+mn-ea"
    face = tfonts.get(face, face) if face.startswith("+") else face
    b = (rpr.get("b") if rpr is not None else None) or (d.get("b") if d is not None else None)
    return int(sz), face, b == "1"


def para_runs(body, p, tfonts):
    out = []
    for r in p:
        tag = etree.QName(r).localname
        if tag in ("r", "fld"):
            sz, face, b = run_props(body, p, r.find(qn("a:rPr")), tfonts)
            out.append((r, r.findtext(qn("a:t")) or "", sz, face, b))
        elif tag == "br":
            out.append((r, "\n", 0, None, False))
    return out


# ------------------------------------------------------------------ 줄바꿈 계산(검증용)
def text_bodies(slide):
    """(txBody, 쓸 수 있는 폭 pt, 줄바꿈 여부, 상자 높이 pt) — 도형과 표 칸"""
    out = []
    for sp in slide._element.iter(qn("p:sp")):
        body = sp.find(qn("p:txBody"))
        ext = sp.find(qn("p:spPr") + "/" + qn("a:xfrm") + "/" + qn("a:ext"))
        if body is None or ext is None:
            continue
        bp = body.find(qn("a:bodyPr"))
        ins = int(bp.get("lIns", 91440)) + int(bp.get("rIns", 91440))
        vins = int(bp.get("tIns", 45720)) + int(bp.get("bIns", 45720))
        vert = bp.get("vert", "horz") != "horz"
        w, h = int(ext.get("cx")), int(ext.get("cy"))
        if vert:
            w, h = h, w
        out.append((body, (w - ins) / EMU_PT, bp.get("wrap") != "none", (h - vins) / EMU_PT))
    for tbl in slide._element.iter(qn("a:tbl")):
        cols = [int(g.get("w")) for g in tbl.iter(qn("a:gridCol"))]
        for tr in tbl.iter(qn("a:tr")):
            ci = 0
            for tc in tr.findall(qn("a:tc")):
                span = int(tc.get("gridSpan", "1"))
                w = sum(cols[ci:ci + span])
                ci += span
                body = tc.find(qn("a:txBody"))
                pr = tc.find(qn("a:tcPr"))
                ins = (int(pr.get("marL", 91440)) + int(pr.get("marR", 91440))) if pr is not None else 182880
                if body is not None and not tc.get("hMerge") and not tc.get("vMerge"):
                    out.append((body, (w - ins) / EMU_PT, True, int(tr.get("h", 0)) / EMU_PT))
    return out


def simulate(body, width, wrap, fonts, tfonts):
    """(줄 수, 글자 높이 pt). 글꼴 파일이 없으면 None"""
    lines_total, height = 0, 0.0
    for p in body.findall(qn("a:p")):
        runs = para_runs(body, p, tfonts)
        ppr = p.find(qn("a:pPr"))
        mar = (int(ppr.get("marL", 0)) + int(ppr.get("indent", 0)) * 0) / EMU_PT if ppr is not None else 0
        avail = max(width - mar, 1)
        spc = 1.0
        ln = ppr.find(qn("a:lnSpc")) if ppr is not None else None
        pts = None
        if ln is not None:
            if ln.find(qn("a:spcPct")) is not None:
                spc = int(ln.find(qn("a:spcPct")).get("val")) / 100000
            elif ln.find(qn("a:spcPts")) is not None:
                pts = int(ln.find(qn("a:spcPts")).get("val")) / 100
        end = p.find(qn("a:endParaRPr"))
        if not runs or all(t == "" for _, t, *_ in runs):
            sz, face, b = run_props(body, p, end, tfonts)
            m = fonts.get(face, b)
            if m is None:
                return None
            lines_total += 1
            height += pts or m[1] * sz / 100 * spc
            continue
        n_lines, cur, lh = 1, 0.0, 0.0
        for r, text, sz, face, b in runs:
            if text == "\n":
                n_lines += 1
                cur = 0.0
                continue
            m = fonts.get(face, b)
            if m is None:
                return None
            measure, lhr = m
            lh = max(lh, lhr * sz / 100)
            rpr = r.find(qn("a:rPr"))
            kern = int(rpr.get("spc", 0)) / 100 if rpr is not None else 0.0   # 자간(pt). 줄 간격 spc와 다른 값
            for tok in re.findall(r"\S+\s*|\s+", text):
                w = measure(tok) * sz / 10000 + kern * len(tok)
                if wrap and cur + measure(tok.rstrip()) * sz / 10000 + kern * len(tok.rstrip()) > avail and cur > 0:
                    n_lines += 1
                    cur = 0.0
                if wrap and w > avail:         # 한 단어가 폭보다 길면 글자 단위로 끊는다(글자 하나는 쪼개지 않음)
                    extra = min(len(tok.strip()) - 1, math.ceil(w / avail) - 1)
                    n_lines += max(extra, 0)
                    w = w - avail * max(extra, 0) if extra > 0 else w
                cur += w
        lines_total += n_lines
        height += n_lines * (pts or lh * spc)
    return lines_total, height


def measure_deck(prs, fonts):
    out = {}
    for si, s in enumerate(prs.slides, 1):
        tf = theme_fonts(s.part)
        for bi, (body, w, wrap, box_h) in enumerate(text_bodies(s)):
            if not "".join(body.itertext()).strip():
                continue
            out[(si, bi)] = (simulate(body, w, wrap, fonts, tf), box_h, body)
    return out


def compare(before, after):
    """줄 수가 늘었거나 상자를 넘치게 된 글상자 목록"""
    worse = []
    for k, (a, box_h, _) in before.items():
        b = after.get(k, (None,))[0]
        if a is None or b is None:
            continue
        taller = b[1] > a[1] * 1.02 and b[1] > box_h
        if b[0] > a[0] or taller:
            worse.append((k, a, b))
    return worse


# ------------------------------------------------------------------ 템플릿 스타일 뽑기
WEIGHTS = (("light", ("light", "thin")), ("bold", ("bold", "heavy", "black")), ("medium", ("medium",)))


def weight(name, bold_ratio=0.0):
    low = name.lower()
    for w, words in WEIGHTS:
        if any(x in low for x in words):
            return w
    return "bold" if bold_ratio > 0.5 else "regular"


def font_counts(elements):
    c, bold = Counter(), Counter()
    for el in elements:
        for rpr in el.iter(*RPR_TAGS):
            for t in FONT_TAGS:
                f = rpr.find(t)
                if f is not None and not f.get("typeface", "+").startswith("+"):
                    c[f.get("typeface")] += 1
                    if rpr.get("b") == "1":
                        bold[f.get("typeface")] += 1
    return c, bold


def all_elements(prs):
    return ([s._element for s in prs.slides] + [m._element for m in prs.slide_masters]
            + [lo._element for m in prs.slide_masters for lo in m.slide_layouts])


def style_from_pptx(path):
    prs = Presentation(path)
    cols = Counter(x.get("val").upper() for el in all_elements(prs) for x in el.iter(qn("a:srgbClr")))
    fonts, bold = font_counts([s._element for s in prs.slides])
    by_w = {}
    for f, n in fonts.most_common():
        by_w.setdefault(weight(f, bold[f] / n), {"font": f, "bold": False})
    return cols, by_w


def style_from_pdf(path):
    import pymupdf
    doc = pymupdf.open(path)
    fill, text, fonts, bold = Counter(), Counter(), Counter(), Counter()
    to_hex = lambda c: "".join(f"{round(v * 255):02X}" for v in c[:3])
    for page in doc:
        for d in page.get_drawings():
            if d.get("fill"):
                fill[to_hex(d["fill"])] += d["rect"].width * d["rect"].height
            if d.get("color"):
                fill[to_hex(d["color"])] += (d["rect"].width + d["rect"].height) * 2   # 선은 길이만큼
        for b in page.get_text("dict")["blocks"]:
            for ln in b.get("lines", []):
                for sp in ln["spans"]:
                    n = len(sp["text"].strip())
                    text[f'{sp["color"]:06X}'] += n
                    ps = sp["font"].split("+")[-1]
                    base = re.sub(r"[-,]?(Bold|SemiBold|ExtraBold|Light|Regular|Medium|Thin)$", "", ps)
                    fam = PS_FONTS.get(base, base)
                    fonts[fam] += n
                    if base != ps and "Bold" in ps[len(base):]:
                        bold[fam] += n
    cols = Counter()
    for src in (fill, text):   # 면적·글자 수를 각각 비율로 바꿔 더한다(흰 배경이 전부를 차지하지 않게)
        tot = sum(v for k, v in src.items() if k != "FFFFFF") or 1
        for k, v in src.items():
            cols[k] += round(v / tot * 10000)
    by_w = {}
    for f, n in fonts.most_common():
        by_w.setdefault("regular", {"font": f, "bold": False})
        if bold[f]:
            by_w.setdefault("bold", {"font": f, "bold": True})
    return cols, by_w


def template_style(path):
    return style_from_pdf(path) if path.lower().endswith(".pdf") else style_from_pptx(path)


def palette(cols, min_share=0.001):
    """쓸 만큼 쓴 색만 남긴다: (유채색 목록, 무채색 목록) — 많이 쓴 순."""
    tot = sum(cols.values()) or 1
    keep = [c for c, n in cols.most_common() if n / tot >= min_share]
    return [c for c in keep if not is_neutral(c)], [c for c in keep if is_neutral(c)]


def brand_from(path):
    cols, by_w = template_style(path)
    chroma, neutral = palette(cols)
    dark = [c for c in chroma if hsl(c)[2] < 0.55]
    texts = [c for c in neutral if hsl(c)[2] < 0.35]
    brand = {
        "primary": dark[0] if dark else (chroma[0] if chroma else None),
        "secondary": next((c for c in dark[1:] if delta_e(c, dark[0]) > 15), None) if dark else None,
        "accent": max(chroma[:4], key=lambda c: hsl(c)[1]) if chroma else None,
        "light": next((c for c in neutral + chroma if 0.9 < hsl(c)[2] < 0.99), None),
        "text": texts[0] if texts else None,
        "font": (by_w.get("regular") or by_w.get("medium") or {}).get("font"),
        "font_bold": (by_w.get("bold") or {}).get("font"),
    }
    return {k: v for k, v in brand.items() if v}


# ------------------------------------------------------------------ 분석
def theme_parts(prs):
    seen = {}
    for m in prs.slide_masters:
        t = m.part.part_related_by(RT.THEME)
        seen[id(t)] = t
    return list(seen.values())


def target_colors(prs):
    """대상의 색 사용량(장표·레이아웃·마스터). 테마 색(schemeClr)도 실제 색으로 바꿔 함께 센다."""
    els = all_elements(prs)
    cols = Counter(x.get("val").upper() for el in els for x in el.iter(qn("a:srgbClr")))
    names = {"tx1": "dk1", "bg1": "lt1", "tx2": "dk2", "bg2": "lt2"}
    theme = {}
    for t in theme_parts(prs):
        scheme = etree.fromstring(t.blob).find(".//" + qn("a:clrScheme"))
        for el in scheme if scheme is not None else []:
            v = el.find(qn("a:srgbClr"))
            if v is not None:
                theme.setdefault(etree.QName(el).localname, set()).add(v.get("val").upper())
    use = Counter(x.get("val") for el in els for x in el.iter(qn("a:schemeClr")))
    for name, n in use.items():
        for hexv in theme.get(names.get(name, name), ()):
            cols[hexv] += n
    return cols, theme


def map_color(c, chroma, neutral, tgt_hue):
    """대상 색 하나 → 템플릿 스타일 색"""
    if is_warning(c):
        return c, "경고색 유지"
    if is_neutral(c):
        if c in ("FFFFFF", "000000") or not neutral:
            return c, "무채색 유지"
        near = min(neutral, key=lambda t: delta_e(t, c))
        if abs(hsl(near)[2] - hsl(c)[2]) <= 0.08:
            return near, f"템플릿 무채색(색차 {delta_e(near, c):.0f})"
        return c, "무채색 유지(가까운 템플릿 색 없음)"
    if not chroma:
        return c, "템플릿에 유채색 없음"
    near = min(chroma, key=lambda t: delta_e(t, c))
    if delta_e(near, c) <= 30:
        return near, f"템플릿 색(색차 {delta_e(near, c):.0f})"
    h, s, l = hsl(c)
    if tgt_hue is not None and (hue_gap(h, tgt_hue) < 45 or l < 0.25 or s < 0.35):
        # 주색 계열(진한 남색·탁한 파랑 등): 템플릿 색 중 밝기가 가장 가까운 색. 밝기 차이가 크면 그 색의 색상·채도에 원래 밝기
        t = min(chroma, key=lambda x: abs(lab(x)[0] - lab(c)[0]))
        th, ts, tl = hsl(t)
        if abs(tl - l) <= 0.08:
            return t, "주색 계열 → 밝기 가까운 템플릿 색"
        return hex_from_hls(th, l, ts), f"주색 계열 → 템플릿 {t}의 색조, 원래 밝기"
    # 다른 계열 강조색(주황·노랑·초록 등): 템플릿에서 가장 선명한 색(템플릿 강조색)
    acc = max(chroma[:6], key=lambda x: hsl(x)[1])
    return acc, f"다른 계열 강조색 → 템플릿 강조색 {acc}"


def ramp(chroma):
    """배경 그림용 팔레트: 검정 → 템플릿 유채색(밝기 순) → 흰색"""
    pts = sorted({c: lab(c)[0] for c in chroma}.items(), key=lambda kv: kv[1])
    out, last = [("000000", 0.0)], 0.0
    for c, L in pts:
        if L - last >= 6:
            out.append((c, round(L, 1)))
            last = L
    out.append(("FFFFFF", 100.0))
    return out


def font_scales(prs, plan_fonts, fonts):
    """글꼴 쌍마다: 같은 글자를 같은 폭으로 쓰려면 몇 배로 줄여야 하는지, 줄 높이는 몇 배인지"""
    fmap = {f["from"]: f for f in plan_fonts}
    w_old, w_new, lh_old, lh_new, missing = Counter(), Counter(), {}, {}, set()
    for s in prs.slides:
        tf = theme_fonts(s.part)
        for body, *_ in text_bodies(s):
            for p in body.findall(qn("a:p")):
                for _, text, sz, face, b in para_runs(body, p, tf):
                    if text == "\n" or face not in fmap:
                        continue
                    m = fmap[face]
                    a, z = fonts.get(face, b), fonts.get(m["to"], b or m["bold"])
                    if a is None or z is None:
                        missing.add(face if a is None else m["to"])
                        continue
                    w_old[face] += a[0](text)
                    w_new[face] += z[0](text)
                    lh_old[face], lh_new[face] = a[1], z[1]
    out = {}
    for f in w_old:
        scale = min(1.0, w_old[f] / w_new[f]) if w_new[f] else 1.0
        line = min(1.0, lh_old[f] / (lh_new[f] * scale))
        out[f] = (round(scale, 3), round(line, 3))
    return out, missing


def analyze(tpl_path, tgt_path, plan_path):
    tgt = Presentation(tgt_path)
    tpl_cols, tpl_w = template_style(tpl_path)
    chroma, neutral = palette(tpl_cols)
    cols, theme = target_colors(tgt)
    tgt_chroma = [c for c, _ in cols.most_common() if not is_neutral(c) and not is_warning(c)]
    tgt_hue = hsl(tgt_chroma[0])[0] if tgt_chroma else None
    theme_of = {v: k for k, vs in theme.items() for v in vs}

    colors, total, acc = [], sum(cols.values()), 0
    for c, n in cols.most_common(80):
        if acc / max(total, 1) > 0.995 and c not in theme_of:
            continue
        acc += n
        to, note = map_color(c, chroma, neutral, tgt_hue)
        if c in theme_of:
            note += f" · 테마 {theme_of[c]}"
        colors.append({"from": c, "to": to, "count": n, "note": note})
    for name, vs in theme.items():   # 쓰지 않은 테마 색도 같이 맞춘다(나중에 새로 쓰는 도형용)
        for c in vs:
            if c not in {x["from"] for x in colors}:
                to, note = map_color(c, chroma, neutral, tgt_hue)
                colors.append({"from": c, "to": to, "count": 0, "note": note + f" · 테마 {name}"})

    tgt_fonts, tgt_bold = font_counts(all_elements(tgt) + [etree.fromstring(t.blob) for t in theme_parts(tgt)])
    fonts = []
    for f, n in tgt_fonts.most_common(15):
        w = weight(f, tgt_bold[f] / n)
        t = tpl_w.get(w) or (tpl_w.get("bold") if w == "bold" else None) or tpl_w.get("regular") or tpl_w.get("medium")
        if t:
            fonts.append({"from": f, "to": t["font"], "bold": bool(t["bold"] and w == "bold"), "count": n})
    fx = Fonts()
    scales, missing = font_scales(tgt, fonts, fx)
    for f in fonts:
        f["scale"], f["line"] = scales.get(f["from"], (1.0, 1.0))

    plan = {"template": tpl_path, "target": tgt_path, "colors": colors, "fonts": fonts,
            "palette": ramp(chroma), "images": True, "brand": brand_from(tpl_path)}
    with open(plan_path, "w", encoding="utf-8") as f:
        json.dump(plan, f, ensure_ascii=False, indent=1)

    p = print
    p(f"템플릿 유채색: {', '.join(chroma[:12])}\n템플릿 무채색: {', '.join(neutral[:8])}\n")
    p("## 색 (asis → tobe)\n\n| asis | tobe | 사용 수 | 비고 |\n|---|---|---:|---|")
    for c in colors:
        p(f"| {c['from']} | {c['to']} | {c['count']} | {c['note']} |")
    p("\n## 글꼴 (글자 크기·줄 간격은 원래 줄 수를 지키도록 보정)\n\n"
      "| asis | tobe | 굵게 | 글자 크기 | 줄 간격 | 사용 수 |\n|---|---|---|---:|---:|---:|")
    for f in fonts:
        p(f"| {f['from']} | {f['to']} | {'예' if f['bold'] else ''} | ×{f['scale']} | ×{f['line']} | {f['count']} |")
    if missing:
        p(f"\n※ 글꼴 파일을 못 찾아 보정하지 못하는 글꼴: {', '.join(sorted(missing))}")
    p("\n## 배경 그림\n\n전면 배경 그림(표지·간지 등)은 밝기에 따라 이 팔레트로 다시 칠한다: "
      + " → ".join(c for c, _ in plan["palette"]))
    p("\n## 새 PPT용 BRAND 추천\n\n" + json.dumps(plan["brand"], ensure_ascii=False))


# ------------------------------------------------------------------ 적용
def restyle(root, cmap, fmap, stats):
    for x in root.iter(qn("a:srgbClr")):
        to = cmap.get(x.get("val").upper())
        if to and to != x.get("val").upper():
            x.set("val", to)
            stats["색"] += 1
    for rpr in list(root.iter(*RPR_TAGS)) + list(root.iter(qn("a:majorFont"), qn("a:minorFont"))):
        bold = False
        for t in FONT_TAGS:
            f = rpr.find(t)
            m = fmap.get(f.get("typeface")) if f is not None else None
            if m:
                if m["to"] != f.get("typeface"):
                    f.set("typeface", m["to"])
                    stats["글꼴"] += 1
                bold = bold or m["bold"]
        if bold and rpr.tag in RPR_TAGS:
            rpr.set("b", "1")


def shrink(sz, k):
    """크기 sz(1/100pt)를 k배로(0.1pt 단위 내림). 6pt 아래로는 내리지 않되, 원래 6pt보다 작던 글자는 키우지 않는다."""
    return max(min(MIN_SZ, sz), math.floor(sz * k / 10) * 10)


def set_sz(r, sz):
    rpr = r.find(qn("a:rPr"))
    if rpr is None:
        rpr = etree.Element(qn("a:rPr"))
        r.insert(0, rpr)
    rpr.set("sz", str(int(sz)))


def rescale_body(body, fmap, tfonts, factor=None):
    """글꼴 바꾸기 전에 부른다: 글꼴 쌍의 배율만큼 글자 크기를, 줄 높이 배율만큼 줄 간격을 줄인다.
    factor를 주면 글꼴과 상관없이 모든 글자를 그 배율로 줄인다(보정 후에도 넘치는 상자용)."""
    for p in body.findall(qn("a:p")):
        line = 1.0
        for r, text, sz, face, b in para_runs(body, p, tfonts):
            if text == "\n":
                continue
            m = fmap.get(face)
            k = factor or (m["scale"] if m else 1.0)
            if factor is None and m:
                line = min(line, m["line"])
            if k < 1.0:
                set_sz(r, shrink(sz, k))
        end = p.find(qn("a:endParaRPr"))
        if end is not None and end.get("sz"):   # 크기를 직접 지정한 경우만(상속값을 줄이면 매번 바뀐 것으로 잘못 판정)
            sz, face, b = run_props(body, p, end, tfonts)
            m = fmap.get(face)
            k = factor or (m["scale"] if m else 1.0)
            if k < 1.0:
                end.set("sz", str(shrink(sz, k)))
        if line < 0.995:
            ppr = p.find(qn("a:pPr"))
            if ppr is None:
                ppr = etree.Element(qn("a:pPr"))
                p.insert(0, ppr)
            ln = ppr.find(qn("a:lnSpc"))
            if ln is None:
                ln = etree.Element(qn("a:lnSpc"))
                etree.SubElement(ln, qn("a:spcPct")).set("val", "100000")
                ppr.insert(0, ln)
            pct = ln.find(qn("a:spcPct"))
            if pct is not None:   # 고정 간격(spcPts)은 글꼴과 무관하므로 그대로
                pct.set("val", str(round(int(pct.get("val")) * line)))


def tighten_body(body):
    """글자를 더 줄일 수 없는 상자(6pt 이하): 자간을 0.1pt씩(최대 -0.8pt), 줄 간격을 5%씩(최소 80%, 한 문단짜리는 70%) 좁힌다."""
    changed = False
    for p in body.findall(qn("a:p")):
        for r in p.findall(qn("a:r")):
            rpr = r.find(qn("a:rPr"))
            if rpr is None:
                rpr = etree.Element(qn("a:rPr"))
                r.insert(0, rpr)
            spc = int(rpr.get("spc", 0))
            if spc > -80:
                rpr.set("spc", str(spc - 10))
                changed = True
        ppr = p.find(qn("a:pPr"))
        if ppr is None:
            ppr = etree.Element(qn("a:pPr"))
            p.insert(0, ppr)
        ln = ppr.find(qn("a:lnSpc"))
        if ln is None:
            ln = etree.Element(qn("a:lnSpc"))
            etree.SubElement(ln, qn("a:spcPct")).set("val", "100000")
            ppr.insert(0, ln)
        pct = ln.find(qn("a:spcPct"))
        floor = 70000 if len(body.findall(qn("a:p"))) == 1 else 80000   # 한 줄 라벨은 70%까지
        if pct is not None and int(pct.get("val")) > floor:
            pct.set("val", str(max(floor, round(int(pct.get("val")) * 0.95))))
            changed = True
    return changed


def recolor_image(blob, ramp_pts):
    """밝기(L*)에 따라 팔레트 색으로 다시 칠한다. 투명도는 유지."""
    import numpy as np
    from PIL import Image
    im = Image.open(io.BytesIO(blob))
    fmt = im.format
    rgba = np.asarray(im.convert("RGBA")).astype(np.float32) / 255
    lin = np.where(rgba[..., :3] <= 0.04045, rgba[..., :3] / 12.92, ((rgba[..., :3] + 0.055) / 1.055) ** 2.4)
    y = lin[..., 0] * .2126 + lin[..., 1] * .7152 + lin[..., 2] * .0722
    L = np.where(y > 0.008856, 116 * np.cbrt(y) - 16, 903.3 * y)
    xs = [L_ for _, L_ in ramp_pts]
    out = np.empty_like(rgba)
    for ch in range(3):
        ys = [int(c[ch * 2:ch * 2 + 2], 16) / 255 for c, _ in ramp_pts]
        out[..., ch] = np.interp(L, xs, ys)
    out[..., 3] = rgba[..., 3]
    res = Image.fromarray((out * 255).round().astype(np.uint8), "RGBA")
    buf = io.BytesIO()
    if fmt == "JPEG":
        res.convert("RGB").save(buf, "JPEG", quality=92)
    else:
        res.save(buf, "PNG", optimize=True)
    return buf.getvalue()


def background_images(prs):
    """장표 면적의 절반 이상을 덮는 그림과 배경 채우기 그림의 part 목록"""
    area = prs.slide_width * prs.slide_height
    hits = {}
    owners = list(prs.slides) + list(prs.slide_masters) + [lo for m in prs.slide_masters for lo in m.slide_layouts]
    for o in owners:
        for pic in o._element.iter(qn("p:pic")):
            ext = pic.find(qn("p:spPr") + "/" + qn("a:xfrm") + "/" + qn("a:ext"))
            if ext is None or int(ext.get("cx")) * int(ext.get("cy")) < area * 0.5:
                continue
            for b in pic.iter(qn("a:blip")):
                if b.get(R_EMBED):
                    part = o.part.related_part(b.get(R_EMBED))
                    hits[id(part)] = part
        for bg in o._element.iter(qn("p:bg")):
            for b in bg.iter(qn("a:blip")):
                if b.get(R_EMBED):
                    part = o.part.related_part(b.get(R_EMBED))
                    hits[id(part)] = part
    return [p for p in hits.values() if p.content_type in ("image/png", "image/jpeg")]


def apply(tgt_path, plan_path, out_path):
    if out_path == tgt_path:
        sys.exit("결과 파일은 원본과 다른 이름이어야 합니다.")
    plan = json.load(open(plan_path, encoding="utf-8"))
    cmap = {c["from"].upper(): c["to"].upper() for c in plan["colors"]}
    fmap = {f["from"]: f for f in plan["fonts"]}
    prs = Presentation(tgt_path)
    fonts = Fonts()
    before = measure_deck(prs, fonts)
    stats = Counter()

    # 1) 글자 크기·줄 간격 보정(글꼴 이름이 바뀌기 전에, 원래 글꼴 기준으로)
    owners = list(prs.slides) + list(prs.slide_masters) + [lo for m in prs.slide_masters for lo in m.slide_layouts]
    for o in owners:
        tf = theme_fonts(o.part)
        bodies = [b for b, *_ in text_bodies(o)] if hasattr(o, "slide_layout") else list(o._element.iter(qn("p:txBody")))
        for body in bodies:
            rescale_body(body, fmap, tf)

    # 2) 색·글꼴
    for part in prs.part.package.iter_parts():
        if not part.content_type.endswith("xml"):
            continue
        el = getattr(part, "_element", None)
        if el is not None:
            restyle(el, cmap, fmap, stats)
        elif b"srgbClr" in part.blob or b"typeface" in part.blob:   # 차트·다이어그램·테마 등
            root = etree.fromstring(part.blob)
            restyle(root, cmap, fmap, stats)
            part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    # 3) 그래도 줄 수·높이가 늘어난 글상자는 그 상자만 3%씩 더 줄이고(하한 6pt),
    #    더 줄일 수 없으면 자간·줄 간격을 좁힌다
    fixed = {"size": set(), "tight": set()}
    for _ in range(12):
        after = measure_deck(prs, fonts)
        worse = compare(before, after)
        if not worse:
            break
        for (si, bi), *_ in worse:
            body = after[(si, bi)][2]
            snap = etree.tostring(body)
            rescale_body(body, {}, theme_fonts(prs.slides[si - 1].part), factor=0.97)
            if etree.tostring(body) != snap:
                fixed["size"].add((si, bi))
            elif tighten_body(body):
                fixed["tight"].add((si, bi))
    stats["글자를 더 줄인 글상자"] = len(fixed["size"])
    stats["자간·줄 간격을 좁힌 글상자"] = len(fixed["tight"])

    # 4) 전면 배경 그림
    if plan.get("images"):
        for part in background_images(prs):
            part._blob = recolor_image(part.blob, plan["palette"])
            stats["다시 칠한 배경 그림"] += 1

    prs.save(out_path)
    print(f"저장: {out_path} ({len(prs.slides)}장)")
    for k, v in stats.items():
        print(f"  {k}: {v}")
    report(before, measure_deck(Presentation(out_path), fonts))


def report(before, after):
    worse = compare(before, after)
    unmeasured = sum(1 for k, (a, *_) in before.items() if a is None or after.get(k, (None,))[0] is None)
    print(f"  줄 수 비교: 글상자 {len(before)}개 중 늘어난 것 {len(worse)}개, 글꼴 파일이 없어 못 잰 것 {unmeasured}개")
    for (si, bi), a, b in worse[:40]:
        print(f"    ※ {si}장 글상자 {bi}: 줄 {a[0]}→{b[0]}, 높이 {a[1]:.0f}→{b[1]:.0f}pt")


def check(orig_path, out_path):
    fonts = Fonts()
    report(measure_deck(Presentation(orig_path), fonts), measure_deck(Presentation(out_path), fonts))


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    cmd, args = (sys.argv[1], sys.argv[2:]) if len(sys.argv) > 1 else ("", [])
    if cmd == "analyze" and len(args) == 3:
        analyze(*args)
    elif cmd == "apply" and len(args) == 3:
        apply(*args)
    elif cmd == "check" and len(args) == 2:
        check(*args)
    elif cmd == "brand" and len(args) == 1:
        print(json.dumps(brand_from(args[0]), ensure_ascii=False, indent=1))
    else:
        print(__doc__)
