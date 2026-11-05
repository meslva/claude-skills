"""제안서 PPT 도식 컴포넌트 라이브러리 (python-pptx).

build_deck.py에서 이렇게 쓴다:

    import sys; sys.path.insert(0, r"<스킬 폴더>/scripts")
    from deck import Deck
    d = Deck(template=None, brand={"company": "우리회사"})
    s = d.content_slide("무중단 인프라 전환 전략",
                        headline="경매가 없는 시간대 일괄 전환과 2회 리허설로 중단 없이 전환합니다.",
                        section="Ⅱ. 구축방안 > 3. 인프라 전환")
    left, right = d.hsplit(d.area(), [0.55, 0.45])
    d.timeline(s, left, periods=[...], rows=[...])
    d.kpi_cards(s, right, [{"value": "0분", "label": "경매 중단"}])
    d.save("제안서_v1.pptx")

원칙
- 모든 좌표 단위는 인치. 박스는 (x, y, w, h) 튜플.
- 글자는 도형 자체의 텍스트 프레임에 넣는다(도형 위에 글상자를 겹치지 않는다).
  텍스트 없는 도형(구분선, 배경 띠, 강조 막대)만 따로 그린다.
- 도형의 테마 스타일(그림자 등)을 제거하고 색·선·글꼴을 명시해 평면적이고 정돈된 느낌을 낸다.
- 차트는 네이티브 차트(편집 가능)로 만든다.
- 본문 텍스트 안의 **굵게** 표기는 강조색 굵은 글씨로 바뀐다.
"""
from __future__ import annotations

import copy
import glob
import os
import re

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

DEFAULT_BRAND = {
    # 회사 공통 가이드 기본값(엠투엠글로벌). 템플릿·고객 키 컬러가 있으면 덮어쓴다.
    "font": "나눔스퀘어",
    "primary": "002060",     # 주 남색: 제목, 헤더 바, 번호 배지
    "secondary": "123E8C",   # 주 파랑: 주요 도형, 헤드라인 강조
    "accent": "0076C0",      # 강조 파랑(권고안·마일스톤·**굵게**). 고객 키 컬러가 있으면 교체
    "positive": "3B7D23",    # 초록: 개선·목표
    "negative": "DC2F34",    # 빨강: 수치 강조·위험·예시 도장·비용 각주
    "light": "F5F7FB",       # 패널 배경
    "light2": "F7F8FA",
    "line": "D9DDE3",        # 테두리·구분선
    "text": "414141",
    "subtext": "808080",
    "white": "FFFFFF",
    "seq": ["163E64", "3B7D23", "0B76A0", "7B2672"],  # 01~04 순서색(남색→초록→청록→보라)
    "company": "",
    "doc_title": "",
    "slide_w": 10.83,        # A4 가로(회사 판형)
    "slide_h": 7.5,
    "margin": 0.25,
    "content_top": 1.85,
    "content_bottom": 7.0,
    "title_layout": None,     # 템플릿 레이아웃 이름(부분 일치). None이면 빈 레이아웃
    "use_title_placeholder": False,
    "draw_footer": True,
    "font_bold": None,        # 굵은 글씨 전용 글꼴(예: "페이퍼로지 7 Bold"). 없으면 font + bold
    "text_scale": 0.75,      # 도식 안 글자 배율. 밀도와 무관하다(처음부터 줄여 채우지 않는다)
    "min_size": 8,           # 글자 하한(사용자 확정: 8pt 미만 금지)
    "line_pt": 1.2,          # 본문 줄 간격을 '글자 크기 × 이 값' pt로 고정(None이면 배수 그대로). 배수 간격은
                             # PowerPoint 화면이 PDF 렌더링보다 줄을 크게 잡아 두 줄 글자가 칸을 넘친다
}

ALIGN = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}
ANCHOR = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}


def rgb(hex_: str) -> RGBColor:
    return RGBColor.from_string(hex_.lstrip("#").upper())


def lerp(c1: str, c2: str, t: float) -> str:
    a = [int(c1[i:i + 2], 16) for i in (0, 2, 4)]
    b = [int(c2[i:i + 2], 16) for i in (0, 2, 4)]
    return "".join(f"{round(x + (y - x) * t):02X}" for x, y in zip(a, b))

ROMANS = {"Ⅰ": 1, "Ⅱ": 2, "Ⅲ": 3, "Ⅳ": 4, "Ⅴ": 5, "Ⅵ": 6, "Ⅶ": 7, "Ⅷ": 8, "Ⅸ": 9, "Ⅹ": 10}


def split_chapter_no(text):
    """'Ⅱ. 구축 방안' → (2, '구축 방안'). 번호가 없으면 (None, text)."""
    t = (text or "").strip()
    m = re.match(r"^([ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]|[IVX]+|\d+)\.?\s*(.*)$", t)
    if not m:
        return None, t
    tok = m.group(1)
    if tok in ROMANS:
        return ROMANS[tok], m.group(2)
    if tok.isdigit():
        return int(tok), m.group(2)
    vals = {"I": 1, "V": 5, "X": 10}
    n = sum(vals[c] for c in tok)
    for a, b in zip(tok, tok[1:]):
        if vals[a] < vals[b]:
            n -= 2 * vals[a]
    return n, m.group(2)


DENSITY = {
    # 밀도 옵션 1~5 = 장표에 담는 내용의 양. 기본 3. 높을수록 내용을 더해 빈 공간을 줄이고,
    # 낮을수록 중요하지 않은 내용을 뺀다. 글자 크기·간격은 바꾸지 않는다(항목 수는 pick/cut).
    # brand["density"] 또는 Deck(density=)로 준다. 예전 표기 상/중/하는 4/3/2로 읽는다.
    5: {"fill": 0.95},   # 최대: 빈 곳 거의 없이 채움
    4: {"fill": 0.90},   # 높음: 제출용 상세 장표(예전 '상')
    3: {"fill": 0.80},   # 보통(기본, 예전 '중')
    2: {"fill": 0.75},   # 낮음(예전 '하')
    1: {"fill": 0.65},   # 최소: 발표용
}
DENSITY_ALIAS = {"상": 4, "중": 3, "하": 2, "높음": 4, "보통": 3, "낮음": 2,
                 "high": 4, "mid": 3, "medium": 3, "low": 2}
DENSITY_DEFAULT = 3


def density_level(v) -> int:
    """밀도 값(1~5 숫자·'3'·상/중/하)을 1~5 정수로 바꾼다. 알 수 없으면 기본 3."""
    if v is None or v == "":
        return DENSITY_DEFAULT
    if isinstance(v, (int, float)):
        return max(1, min(5, int(v)))
    t = str(v).strip().replace("밀도", "").replace("=", "").strip()
    if t.isdigit():
        return max(1, min(5, int(t)))
    return DENSITY_ALIAS.get(t, DENSITY_DEFAULT)


class Deck:
    # ------------------------------------------------------------------ 기본
    def __init__(self, template: str | None = None, brand: dict | None = None,
                 proto: dict | None = None, density: int | str | None = None):
        """template: 템플릿 pptx 경로. proto: 템플릿 원형 장표 복제 설정(use 'python deck.py --inspect'
        로 도형 이름을 확인해 만든다). proto가 있으면 템플릿 장표를 원형으로 남겨 두고 저장할 때 지운다.
        density: 1~5 — 장표에 담는 내용의 양. pick()·cut()의 항목 수와 채움 목표(fill)를 정한다.
        글자 크기·간격은 바꾸지 않는다(brand["density"]로도 준다). **기본값은 3**.
        예전 표기 상/중/하는 4/3/2로 읽는다."""
        self.b = {**DEFAULT_BRAND, **(brand or {})}
        dk = density_level(density if density is not None else self.b.get("density"))
        self.b["density"] = dk
        d = DENSITY[dk]
        self.density = dk           # 1~5
        self.dg = 1.0               # 컴포넌트 내부 간격 배율(밀도와 무관)
        self.fill = d["fill"]       # 본문 영역을 채울 목표 비율(렌더링 확인 기준)
        self.proto = proto or {}
        self._protos = {}
        self._n_proto = 0
        self._tpl_shapes = set()
        self._raw = False
        self._chapter = None
        if template:
            self.prs = Presentation(template)
            self.b["slide_w"] = self.prs.slide_width / 914400
            self.b["slide_h"] = self.prs.slide_height / 914400
            if self.proto:
                slides = list(self.prs.slides)
                self._n_proto = len(slides)
                for kind, spec in self.proto.items():
                    self._protos[kind] = slides[spec["slide"] - 1]
                self.b["draw_footer"] = False
            else:
                self._clear_slides()
        else:
            self.prs = Presentation()
            self.prs.slide_width = Inches(self.b["slide_w"])
            self.prs.slide_height = Inches(self.b["slide_h"])
        self.page = 0

    def save(self, path: str, keep_prototypes: bool = False):
        if self.b.get("line_pt"):
            self.fix_line_spacing(self.b["line_pt"])
        if self._n_proto and not keep_prototypes:
            lst = self.prs.slides._sldIdLst
            for sid in list(lst)[:self._n_proto]:
                self.prs.part.drop_rel(sid.rId)
                lst.remove(sid)
            self._n_proto = 0
        self.prs.save(path)

    def fix_line_spacing(self, k=1.2, top=None):
        """본문 영역(top 인치 아래) 문단의 줄 간격을 고정 pt(가장 큰 글자 × k)로 바꾼다. save()가 자동으로 부른다.
        배수 줄 간격은 보는 프로그램마다 글꼴 줄 높이를 다르게 잡으므로, 렌더링 확인 화면과 PowerPoint 화면을 같게 만든다.
        템플릿 원형에서 복제한 도형(머리말·제안 포인트·간지 번호 등)은 템플릿 간격을 그대로 둔다."""
        if top is None:
            top = (self.proto.get("content", {}).get("area") or [0, self.b["content_top"]])[1] - 0.1

        def paras(sh):
            if sh.shape_type == 6:
                for c in sh.shapes:
                    yield from paras(c)
                return
            if sh.top is None or sh.top / 914400 < top:
                return
            if sh.has_text_frame:
                yield from sh.text_frame.paragraphs
            if getattr(sh, "has_table", False) and sh.has_table:
                for r in sh.table.rows:
                    for c in r.cells:
                        yield from c.text_frame.paragraphs

        for sl in self.prs.slides:
            for sh in sl.shapes:
                if (sl.slide_id, str(sh.shape_id)) in self._tpl_shapes:
                    continue
                for p in paras(sh):
                    sz = [r.font.size.pt for r in p.runs if r.font.size and r.text.strip()]
                    if sz:
                        p.line_spacing = Pt(round(max(sz) * k, 1))

    @staticmethod
    def row_heights(lines, size, total=None, head=0.27, pad=0.14):
        """표 행 높이(인치): 행마다 줄 수 × 글자 크기 × 1.2 + 위아래 여백. total(표 전체 높이)을 주면 남는 높이를
        모든 행에 같은 여백으로 나눈다(한 행·한 구역만 빽빽하거나 헐렁하지 않게). 여백이 pad보다 작아지면 경고한다."""
        lh = size * 1.2 / 72
        if total is not None:
            pad_ = (total - head - sum(lines) * lh) / len(lines)
            if pad_ < pad:
                print(f"[row_heights] 행 여백 {pad_:.2f}in < {pad} — 위 구역(도식·카드)을 줄여 표에 자리를 주거나 내용을 줄인다")
            pad = pad_
        return [n * lh + pad for n in lines]

    def step_label(self, slide, box, no, title, color, size=10, floor=9):
        """흐름 도형(셰브런·오각 화살표) 라벨: 번호와 이름을 한 줄에 같은 크기로 넣는다(줄바꿈 없음).
        칸에 안 들어가면 0.5pt씩 줄이고(하한 floor) 그래도 넘치면 ValueError — 이름을 줄이거나 칸을 넓힌다.
        한 줄(행)에 놓인 라벨은 모두 같은 size로 부른다."""
        x, y, w, h = box
        avail = (w - 0.04) * 72
        em = lambda t: sum(1.0 if ord(c) > 0x3000 else 0.35 if c == " " else 0.55 for c in t)
        text = f"{no}  {title}" if no else title
        sz = size
        while em(text) * sz > avail and sz > floor:
            sz -= 0.5
        if em(text) * sz > avail:
            raise ValueError(f"흐름 라벨이 한 줄에 안 들어감: '{text}' {em(text) * sz:.0f}pt > {avail:.0f}pt")
        tb = self.text(slide, box, [{"text": text, "size": sz, "color": color}], align="c", anchor="m", space_after=0)
        tb.text_frame.word_wrap = False
        return tb

    def _clear_slides(self):
        lst = self.prs.slides._sldIdLst
        for sid in list(lst):
            self.prs.part.drop_rel(sid.rId)
            lst.remove(sid)

    def _all_layouts(self):
        return [lo for m in self.prs.slide_masters for lo in m.slide_layouts]

    def _layout(self, name: str | None):
        layouts = self.prs.slide_layouts
        if name:
            alls = self._all_layouts()
            for lo in alls:
                if lo.name == name:
                    return lo
            for lo in alls:
                if name.lower() in lo.name.lower():
                    return lo
        for key in ("blank", "빈 화면", "빈화면"):
            for lo in layouts:
                if key in lo.name.lower():
                    return lo
        return layouts[len(layouts) - 1]

    def area(self) -> tuple:
        """현재 장표의 본문 영역. 슬로건·키워드를 넣었으면 그 아래부터."""
        if self.proto.get("content", {}).get("area"):
            x, y, w, h = self.proto["content"]["area"]
        else:
            m = self.b["margin"]
            x, y, w, h = (m, self.b["content_top"], self.b["slide_w"] - 2 * m,
                          self.b["content_bottom"] - self.b["content_top"])
        off = getattr(self, "_top_off", 0)
        return (x, y + off, w, h - off)

    def pick(self, *vals):
        """밀도에 따라 값을 고른다. 값은 높은 밀도부터 적는다.
        5개: `d.pick(7, 6, 5, 4, 3)` — 밀도 5면 7, 1이면 3.
        3개(예전 상/중/하 방식): `d.pick(6, 4, 3)` — 밀도 5·4는 6, 3은 4, 2·1은 3."""
        if len(vals) == 5:
            return vals[5 - self.density]
        if len(vals) == 3:
            return vals[{5: 0, 4: 0, 3: 1, 2: 2, 1: 2}[self.density]]
        raise ValueError("pick()에는 값을 5개(밀도 5→1) 또는 3개(상·중·하) 넣는다")

    def cut(self, items):
        """밀도에 따라 목록을 잘라 쓴다(5·4=전부, 3=최대 5, 2=최대 4, 1=최대 3)."""
        return items[:self.pick(len(items), len(items), 5, 4, 3)]

    @staticmethod
    def hsplit(box, ratios, gap=0.3):
        x, y, w, h = box
        usable = w - gap * (len(ratios) - 1)
        tot = sum(ratios)
        out, cx = [], x
        for r in ratios:
            cw = usable * r / tot
            out.append((cx, y, cw, h))
            cx += cw + gap
        return out

    @staticmethod
    def vsplit(box, ratios, gap=0.25):
        x, y, w, h = box
        usable = h - gap * (len(ratios) - 1)
        tot = sum(ratios)
        out, cy = [], y
        for r in ratios:
            ch = usable * r / tot
            out.append((x, cy, w, ch))
            cy += ch + gap
        return out

    @staticmethod
    def inset(box, dx=0.1, dy=None):
        dy = dx if dy is None else dy
        x, y, w, h = box
        return (x + dx, y + dy, w - 2 * dx, h - 2 * dy)

    # ------------------------------------------------------------ 저수준 도구
    def _strip_style(self, shp):
        st = shp._element.find(qn("p:style"))
        if st is not None:
            shp._element.remove(st)

    def _font(self, run, size=12, bold=False, color=None, italic=False):
        f = run.font
        if not self._raw:
            size = max(self.b["min_size"], round(size * self.b["text_scale"] * 2) / 2)
        f.size = Pt(size)
        face = self.b["font"]
        if bold and self.b.get("font_bold"):
            face, bold = self.b["font_bold"], False
        f.bold = bold
        f.italic = italic
        f.name = face
        f.color.rgb = rgb(color or self.b["text"])
        rpr = run._r.get_or_add_rPr()
        for tag in ("a:ea", "a:cs"):
            el = rpr.find(qn(tag))
            if el is None:
                el = etree.SubElement(rpr, qn(tag))
            el.set("typeface", face)

    def _bullet(self, p, char="•", level=0, indent=0.18):
        ppr = p._p.get_or_add_pPr()
        ppr.set("marL", str(int(Inches(indent * (level + 1)))))
        ppr.set("indent", str(int(-Inches(indent))))
        for tag in ("a:buNone", "a:buChar", "a:buAutoNum"):
            el = ppr.find(qn(tag))
            if el is not None:
                ppr.remove(el)
        old = ppr.find(qn("a:buFont"))
        if old is not None:
            ppr.remove(old)
        etree.SubElement(ppr, qn("a:buFont")).set("typeface", "Arial")
        bu = etree.SubElement(ppr, qn("a:buChar"))
        bu.set("char", char)

    def _runs(self, p, text, size, bold, color, hl_color=None):
        """**굵게** 구간을 강조색 굵은 글씨로."""
        parts = re.split(r"(\*\*.+?\*\*)", text)
        for part in parts:
            if not part:
                continue
            r = p.add_run()
            if part.startswith("**") and part.endswith("**"):
                r.text = part[2:-2]
                self._font(r, size, True, hl_color or self.b["accent"])
            else:
                r.text = part
                self._font(r, size, bold, color)

    def _fill_text(self, tf, paras, size=12, bold=False, color=None, align="l",
                   anchor="m", margin=0.08, spacing=1.1, space_after=3, hl_color=None):
        """paras: str | list[str | dict(text,size,bold,color,level,bullet,align,space_after)]."""
        tf.word_wrap = True
        tf.vertical_anchor = ANCHOR[anchor]
        for side in ("left", "right"):
            setattr(tf, f"margin_{side}", Inches(margin))
        for side in ("top", "bottom"):
            setattr(tf, f"margin_{side}", Inches(margin * 0.6))
        if isinstance(paras, str):
            paras = paras.split("\n")
        first = True
        for item in paras:
            d = item if isinstance(item, dict) else {"text": item}
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            p.alignment = ALIGN[d.get("align", align)]
            p.line_spacing = spacing
            p.space_after = Pt(d.get("space_after", space_after))
            if d.get("bullet"):
                self._bullet(p, d.get("bullet") if isinstance(d.get("bullet"), str) else "•",
                             d.get("level", 0))
            self._runs(p, str(d.get("text", "")), d.get("size", size), d.get("bold", bold),
                       d.get("color", color), hl_color)

    def box(self, slide, box, text="", fill=None, line=None, shape="rect", size=12,
            bold=False, color=None, align="c", anchor="m", margin=0.08, radius=0.08,
            line_w=0.75, dash=False, hl_color=None, space_after=3):
        """텍스트를 품은 단일 도형. fill/line에 None이면 없음."""
        kinds = {"rect": MSO_SHAPE.RECTANGLE, "round": MSO_SHAPE.ROUNDED_RECTANGLE,
                 "oval": MSO_SHAPE.OVAL, "chevron": MSO_SHAPE.CHEVRON,
                 "pentagon": MSO_SHAPE.PENTAGON, "arrow": MSO_SHAPE.RIGHT_ARROW,
                 "down": MSO_SHAPE.DOWN_ARROW, "diamond": MSO_SHAPE.DIAMOND,
                 "tri": MSO_SHAPE.ISOSCELES_TRIANGLE}
        x, y, w, h = box
        shp = slide.shapes.add_shape(kinds[shape], Inches(x), Inches(y), Inches(w), Inches(h))
        self._strip_style(shp)
        if shape == "round":
            shp.adjustments[0] = radius
        if fill:
            shp.fill.solid()
            shp.fill.fore_color.rgb = rgb(fill)
        else:
            shp.fill.background()
        if line:
            shp.line.color.rgb = rgb(line)
            shp.line.width = Pt(line_w)
            if dash:
                shp.line.dash_style = MSO_LINE_DASH_STYLE.DASH
        else:
            shp.line.fill.background()
        if text != "" and text is not None:
            self._fill_text(shp.text_frame, text, size, bold, color, align, anchor, margin,
                            hl_color=hl_color, space_after=space_after)
        return shp

    def text(self, slide, box, paras, size=12, bold=False, color=None, align="l", anchor="t",
             margin=0.02, spacing=1.1, space_after=4, hl_color=None):
        x, y, w, h = box
        tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        self._fill_text(tb.text_frame, paras, size, bold, color, align, anchor, margin,
                        spacing, space_after, hl_color)
        return tb

    def line(self, slide, x1, y1, x2, y2, color=None, width=1.0, arrow=False, dash=False,
             elbow=False):
        kind = MSO_CONNECTOR.ELBOW if elbow else MSO_CONNECTOR.STRAIGHT
        c = slide.shapes.add_connector(kind, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
        self._strip_style(c)
        c.line.color.rgb = rgb(color or self.b["line"])
        c.line.width = Pt(width)
        if dash:
            c.line.dash_style = MSO_LINE_DASH_STYLE.DASH
        if arrow:
            ln = c.line._get_or_add_ln()
            t = etree.SubElement(ln, qn("a:tailEnd"))
            t.set("type", "triangle")
            t.set("w", "med")
            t.set("len", "med")
        return c

    def image(self, slide, box, path):
        """이미지를 박스 안에 비율 유지로 맞춰 넣는다(아키텍처 등 복잡 도식, 원본은 _도식원본/에 보관)."""
        from PIL import Image
        x, y, w, h = box
        with Image.open(path) as im:
            iw, ih = im.size
        r = min(w / iw, h / ih)
        pw, ph = iw * r, ih * r
        return slide.shapes.add_picture(path, Inches(x + (w - pw) / 2), Inches(y + (h - ph) / 2),
                                        Inches(pw), Inches(ph))

    @staticmethod
    def est_h(items, width, size=11.5, extra=0.3):
        """개조식 항목 목록의 대략적 높이(인치). 한글 1자 너비 ≈ size*0.0145 인치로 어림한다."""
        per_line = max(int(width / (size * 0.0145)), 8)
        lines = sum(max(1, -(-len(str(t)) // per_line)) for t in items)
        return lines * size * 0.0205 + len(items) * 0.06 + extra

    # ------------------------------------------------------------- 장표 뼈대
    def _new_slide(self, layout=None):
        self.page += 1
        return self.prs.slides.add_slide(self._layout(layout or self.b["title_layout"]))

    def _footer(self, s, source=None):
        b = self.b
        if not b["draw_footer"]:
            return
        W, H, m = b["slide_w"], b["slide_h"], b["margin"]
        if b["doc_title"] or b["company"]:
            self.text(s, (m, H - 0.3, 5, 0.22), f"{b['doc_title']}  {b['company']}".strip(),
                      size=7, color=b["subtext"])
        self.text(s, (W / 2 - 0.4, H - 0.3, 0.8, 0.22), str(self.page), size=8,
                  color=b["subtext"], align="c")
        if source:
            self.text(s, (m, b["content_bottom"] + 0.02, W - 2 * m, 0.2), f"※ {source}",
                      size=7, color=b["subtext"])

    def content_slide(self, title, headline=None, section=None, source=None, notes=None,
                      layout=None, slogan=None, keywords=None):
        """제목 + 헤드메시지 + 바닥글을 갖춘 본문 장표. section은 "장 > 절" 형식.
        템플릿 원형 모드면 원형 장표를 복제해 제목·헤드메시지·장·절 글자만 바꾼다."""
        b = self.b
        if "content" in self.proto:
            chap, _, sec = (section or "").partition(">")
            no, chap = split_chapter_no(chap)
            s = self.clone("content", {"title": title, "headline": headline or "",
                                       "chapter": chap.strip(), "section": sec.strip()},
                           autonum={"chapter": no or self._chapter})
            self._top_off = 0
            self._slogan(s, slogan, keywords)
            if source:
                x, y, w, h = self.area()
                self.text(s, (x, y + h - 0.02, w, 0.25), f"※ {source}", size=8, color=b["subtext"])
            if notes:
                s.notes_slide.notes_text_frame.text = notes
            return s
        self._raw = True
        self._top_off = 0
        s = self._new_slide(layout)
        W, m = b["slide_w"], b["margin"]
        # 회사 공통 가이드 틀: 좌상 제목(18pt) + 가는 선, 우상 장·절(10/12pt), Governing Message 상자
        if b["use_title_placeholder"] and s.shapes.title is not None:
            s.shapes.title.text = title
        else:
            self.text(s, (0.15, 0.12, 7.4, 0.4), title, size=18, bold=True, color=b["primary"],
                      anchor="m")
        self.line(s, 0.15, 0.56, W - 0.15, 0.56, b["primary"], 0.75)
        if section:
            chap, _, sec = section.partition(">")
            self.text(s, (W - 3.2, 0.04, 3.05, 0.24), chap.strip(), size=10, bold=True,
                      color=b["primary"], align="r")
            self.text(s, (W - 3.2, 0.27, 3.05, 0.26), sec.strip(), size=12, color=b["primary"],
                      align="r")
        if headline:
            self.box(s, (0.19, 0.72, 1.3, 0.2), "Governing Message", fill=b["primary"], size=7,
                     bold=True, color=b["white"], margin=0.02)
            self.box(s, (0.19, 0.92, W - 0.38, 0.05), fill=b["primary"])
            self.box(s, (0.19, 0.97, W - 0.38, 0.62), headline, fill=b["white"], line=b["line"],
                     size=13, bold=True, color=b["text"], align="l", margin=0.15)
        self._footer(s, source)
        self._raw = False
        self._slogan(s, slogan, keywords)
        if notes:
            s.notes_slide.notes_text_frame.text = notes
        return s

    def cover(self, title, subtitle="", date="", company=""):
        b = self.b
        if "cover" in self.proto:
            return self.clone("cover", {"title": title, "slogan": subtitle,
                                        "date": [date, company or b["company"]]})
        s = self._new_slide()
        self._raw = True
        W, H = b["slide_w"], b["slide_h"]
        self.box(s, (0, 0, W, H), fill=b["primary"])
        self.box(s, (0.8, 2.35, 0.12, 1.6), fill=b["accent"])
        self.text(s, (1.15, 2.2, W - 2, 1.3), title, size=34, bold=True, color=b["white"],
                  anchor="m")
        if subtitle:
            self.text(s, (1.15, 3.5, W - 2, 0.6), subtitle, size=18, color="C9D4E5")
        self.text(s, (1.15, H - 1.3, W - 2, 0.8), [date, company or b["company"]], size=14,
                  color="C9D4E5")
        self._raw = False
        return s

    def section(self, number, title, desc="", items=None):
        """간지. number: 1, 2 … (템플릿 원형 모드면 번호별 레이아웃 사용), items: 절 목록."""
        b = self.b
        self._chapter = number if isinstance(number, int) else split_chapter_no(str(number))[0]
        if "section" in self.proto:
            spec = self.proto["section"]
            lo = spec.get("layouts", {}).get(number) or spec.get("layouts", {}).get(str(number))
            return self.clone("section", {"title": title, "items": items or [desc]}, layout=lo)
        s = self._new_slide()
        self._raw = True
        W, H = b["slide_w"], b["slide_h"]
        self.box(s, (0, 0, W * 0.38, H), fill=b["primary"])
        self.text(s, (0.8, H / 2 - 1.2, W * 0.38 - 1.2, 1.4), str(number), size=72, bold=True,
                  color=b["white"], anchor="m")
        self.text(s, (W * 0.38 + 0.8, H / 2 - 0.8, W * 0.55, 0.9), title, size=32, bold=True,
                  color=b["primary"], anchor="m")
        if desc or items:
            self.text(s, (W * 0.38 + 0.8, H / 2 + 0.15, W * 0.52, 2.5), items or desc, size=14,
                      color=b["subtext"])
        self._raw = False
        return s

    def toc(self, items=None, title="목차"):
        """items: [(번호, 제목, 설명)]. 템플릿 원형 모드면 원형 목차를 그대로 복제한다."""
        b = self.b
        if "toc" in self.proto:
            return self.clone("toc")
        s = self._new_slide()
        self._raw = True
        W = b["slide_w"]
        self.text(s, (0.8, 0.6, 5, 0.8), title, size=30, bold=True, color=b["primary"])
        n = len(items)
        top, h = 1.7, min(0.95, 5.0 / max(n, 1))
        for i, it in enumerate(items):
            num, name = it[0], it[1]
            desc = it[2] if len(it) > 2 else ""
            y = top + i * h
            self.box(s, (0.8, y + 0.1, 0.75, h - 0.2), str(num), fill=b["primary"], size=18,
                     bold=True, color=b["white"])
            self.text(s, (1.75, y + 0.05, 4.5, h - 0.1), name, size=18, bold=True,
                      color=b["text"], anchor="m")
            if desc:
                self.text(s, (6.3, y + 0.05, W - 7.1, h - 0.1), desc, size=12,
                          color=b["subtext"], anchor="m")
            self.line(s, 0.8, y + h, W - 0.8, y + h, b["line"], 0.5)
        self._raw = False
        return s

    def end(self, text="감사합니다."):
        """마지막 장. 템플릿 원형 모드면 원형을 복제한다."""
        b = self.b
        if "end" in self.proto:
            return self.clone("end")
        s = self._new_slide()
        self._raw = True
        self.box(s, (0, 0, b["slide_w"], b["slide_h"]), text, fill=b["primary"], size=36,
                 bold=True, color=b["white"])
        self._raw = False
        return s

    # ====================================================== 회사 스타일 컴포넌트
    # 수주 제안서(양부대 스마트물류·NH 싱씽몰)와 회사 공통 가이드에서 뽑은 도식.
    # 한 장표 = 주 도식 + 보조 블록 + 하단 요약 바 조합을 기본으로 한다.

    ICON_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "assets", "icons")
    _icon_index = None

    def seq(self, i):
        """01~04 순서색."""
        q = self.b["seq"]
        return q[i % len(q)]

    def _slogan(self, s, slogan, keywords):
        """본문 카드 맨 위 가운데: 따옴표 슬로건(명사형 대구) + 영문 키워드 한 줄."""
        if not slogan and not keywords:
            return
        x, y, w, h = self.area()
        self._raw = True
        used = 0
        if slogan:
            t = slogan if slogan.startswith("“") else f"“ {slogan} ”"
            self.text(s, (x, y, w, 0.3), t, size=14, color=self.b["accent"], align="c")
            used += 0.3
        if keywords:
            k = keywords if isinstance(keywords, str) else "  ·  ".join(keywords)
            self.text(s, (x, y + used, w, 0.22), k, size=8, color=self.b["subtext"], align="c")
            used += 0.24
        self._raw = False
        self._top_off = used + 0.08

    def _icon_path(self, name):
        if not name:
            return None
        if os.path.isfile(str(name)):
            return name
        if re.fullmatch(r"i\d{3}", str(name)):
            hits = glob.glob(os.path.join(self.ICON_DIR, f"{name}_*.png"))
            return hits[0] if hits else None
        if Deck._icon_index is None:
            Deck._icon_index = []
            idx = os.path.join(self.ICON_DIR, "index.md")
            if os.path.exists(idx):
                for ln in open(idx, encoding="utf-8"):
                    m = re.match(r"\|\s*(i\d{3})(?:~i\d{3})?\s*\|\s*(.+?)\s*\|", ln)
                    if m:
                        Deck._icon_index.append((m.group(1), m.group(2)))
        for code, kw in Deck._icon_index:
            if str(name) in kw:
                return self._icon_path(code)
        return None

    def icon(self, s, box, name, circle=None, pad=0.18):
        """아이콘. name: 'i245' 또는 assets/icons/index.md 키워드('보안').
        circle: 원 배경색(흰 아이콘 묶음 C를 쓸 때는 진한 색). 없으면 자리표시 원."""
        x, y, w, h = box
        d = min(w, h)
        x, y = x + (w - d) / 2, y + (h - d) / 2
        path = self._icon_path(name)
        if circle == "ring":
            self.box(s, (x, y, d, d), fill=self.b["white"], line="C9CED6", line_w=1, shape="oval")
            x, y, d = x + d * 0.22, y + d * 0.22, d * 0.56
        elif circle:
            self.box(s, (x, y, d, d), fill=circle, shape="oval")
            x, y, d = x + d * pad, y + d * pad, d * (1 - 2 * pad)
        if path:
            return self.image(s, (x, y, d, d), path)
        return self.box(s, (x, y, d, d), "", fill=self.b["light"], line=self.b["line"], shape="oval")

    def section_bar(self, s, box, text, color=None, size=12):
        """도식 제목 띠(명사구). 절차면 'A → B → C'를 그대로 쓴다."""
        return self.box(s, box, text, fill=color or self.b["primary"], size=size, bold=True,
                        color=self.b["white"], shape="round", radius=0.12)

    def tab_panel(self, s, box, label, color=None, fill=None):
        """좌상단에 탭 라벨이 걸친 패널. 반환: 안쪽 영역."""
        b = self.b
        x, y, w, h = box
        col = color or b["primary"]
        self.box(s, (x, y + 0.13, w, h - 0.13), fill=fill or b["light2"], line=b["line"])
        tw = min(w * 0.6, 0.25 + len(label) * 0.11)
        self.box(s, (x + 0.08, y, tw, 0.26), label, fill=col, size=9, bold=True, color=b["white"],
                 shape="round", radius=0.3, margin=0.04)
        return (x + 0.1, y + 0.36, w - 0.2, h - 0.44)

    def stamp(self, s, x, y, text="예시", w=0.62, h=0.24):
        """예시·Illustrative 도장(빨간 테두리, 기울임). 실물이 아닌 화면·가정 수치에 붙인다."""
        shp = self.box(s, (x, y, w, h), text, fill=self.b["white"], line=self.b["negative"], line_w=1.25,
                       size=9, bold=True, color=self.b["negative"], margin=0.02)
        shp.rotation = -12
        return shp

    def header_card(self, s, box, title, bullets=(), no=None, en=None, sub=None, color=None,
                    icon=None, style="bar", size=11):
        """번호 카드: [01 + 제목(+영문)] 머리띠 → 아이콘 → 한 줄 부제 → 개조식.
        style='bar'(색 머리띠) | 'border'(색 테두리 + 색 글자 제목)."""
        b = self.b
        x, y, w, h = box
        col = color or b["primary"]
        head = f"{no}  {title}" if no else title
        hh = 0.5 if en else 0.38
        if style == "bar":
            paras = [{"text": head, "bold": True, "size": size + 2, "color": b["white"], "space_after": 0,
                      "align": "c"}]
            if en:
                paras.append({"text": en, "size": size - 2, "color": "DCE6F5", "align": "c"})
            self.box(s, (x, y, w, hh), paras, fill=col, margin=0.04)
            self.box(s, (x, y + hh, w, h - hh), fill=b["white"], line=b["line"])
        else:
            self.box(s, (x, y, w, h), fill=b["white"], line=col, line_w=1.5, shape="round", radius=0.04)
            paras = [{"text": head, "bold": True, "size": size + 3, "color": col, "space_after": 0, "align": "c"}]
            if en:
                paras.append({"text": en, "size": size - 2, "color": b["subtext"], "align": "c"})
            self.text(s, (x, y + 0.06, w, hh), paras, align="c")
        cy = y + hh + 0.08
        n_lines = len(bullets) + (1 if sub else 0)
        est = (0.5 if icon else 0) + n_lines * size * self.b["text_scale"] / 72 * 1.55 + 0.1
        cy += max(0.0, min((y + h - cy - est) / 2 - 0.05, 0.25))
        if icon:
            self.icon(s, (x + w / 2 - 0.22, cy, 0.44, 0.44), icon)
            cy += 0.5
        paras = []
        if sub:
            paras.append({"text": sub, "size": size - 1, "color": b["subtext"], "align": "c", "space_after": 4})
        paras += [{"text": t, "bullet": "•", "size": size, "space_after": 3} for t in bullets]
        if paras:
            self.text(s, (x + 0.08, cy, w - 0.16, y + h - cy - 0.05), paras)

    def numbered_cards(self, s, box, cards, gap=0.18, style="bar", icon_row=True, size=11):
        """번호 카드 가로 배열. cards: [{title, en?, sub?, bullets, icon?}] — 순서색 자동."""
        cells = self.hsplit(box, [1] * len(cards), gap)
        for i, (c, cell) in enumerate(zip(cards, cells)):
            self.header_card(s, cell, c["title"], c.get("bullets", []), no=f"{i + 1:02d}", en=c.get("en"),
                             sub=c.get("sub"), color=c.get("color", self.seq(i)),
                             icon=c.get("icon") if icon_row else None, style=style, size=size)

    def summary_bar(self, s, box, items, size=11):
        """하단 요약 바: 흰 상자를 세로선으로 나눈 칸마다 [원 아이콘 + 색 제목 + 한 줄 설명]."""
        b = self.b
        x, y, w, h = box
        self.box(s, box, fill=b["white"], line=b["line"])
        n = len(items)
        cw = w / n
        d = min(h - 0.16, 0.46)
        for i, it in enumerate(items):
            cx = x + i * cw
            if i:
                self.line(s, cx, y + 0.1, cx, y + h - 0.1, b["line"], 0.75)
            col = it.get("color", self.seq(i))
            self.icon(s, (cx + 0.1, y + (h - d) / 2, d, d), it.get("icon"), circle=it.get("circle"))
            self.text(s, (cx + 0.18 + d, y + 0.06, cw - d - 0.26, h - 0.12),
                      [{"text": it["title"], "bold": True, "size": size + 1, "color": col, "space_after": 1},
                       {"text": it.get("desc", ""), "size": size - 2, "color": b["text"]}], anchor="m")

    def kpi_strip(self, s, box, items, size=11):
        """KPI 띠: 둥근 테두리 안에 [아이콘 + 큰 숫자 + 라벨] 칸, 세로 구분선."""
        b = self.b
        x, y, w, h = box
        self.box(s, box, fill=b["white"], line=b["line"], shape="round", radius=0.3)
        n = len(items)
        cw = w / n
        d = min(h * 0.6, 0.42)
        for i, it in enumerate(items):
            cx = x + i * cw
            if i:
                self.line(s, cx, y + h * 0.22, cx, y + h * 0.78, b["line"], 0.75)
            self.icon(s, (cx + 0.15, y + (h - d) / 2, d, d), it.get("icon"))
            self.text(s, (cx + 0.22 + d, y + 0.04, cw - d - 0.3, h - 0.08),
                      [{"text": it["value"], "bold": True, "size": size + 6, "color": b["primary"], "space_after": 0},
                       {"text": it["label"], "size": size - 1, "color": b["subtext"]}], anchor="m")

    def pillar_matrix(self, s, box, pillars, row_labels, size=11, label_w=0.85):
        """기둥 매트릭스: 열=전략/기둥(순서색 2줄 머리), 행=라벨(핵심 업무/산출물/검증 포인트).
        pillars: [{title, en?, rows: [[bullets], [bullets], ...]}]"""
        b = self.b
        x, y, w, h = box
        hh = 0.55
        n = len(pillars)
        cells = self.hsplit((x + label_w + 0.08, y, w - label_w - 0.08, h), [1] * n, 0.12)
        nr = len(row_labels)
        rh = (h - hh - 0.08) / nr
        for r, lab in enumerate(row_labels):
            ry = y + hh + 0.08 + r * rh
            self.box(s, (x, ry + 0.04, label_w, rh - 0.08), lab, fill=b["light"], size=size, bold=True,
                     color=b["primary"], margin=0.04)
            if r:
                self.line(s, x + label_w + 0.08, ry, x + w, ry, b["line"], 0.75, dash=True)
        for i, (pl, cell) in enumerate(zip(pillars, cells)):
            cx, cy, cw, ch = cell
            col = pl.get("color", self.seq(i))
            paras = [{"text": pl["title"], "bold": True, "size": size + 2, "color": b["white"], "space_after": 0,
                      "align": "c"}]
            if pl.get("en"):
                paras.append({"text": pl["en"], "size": size - 1, "color": "DCE6F5", "align": "c"})
            self.box(s, (cx, cy, cw, hh), paras, fill=col, margin=0.03)
            self.box(s, (cx, cy + hh, cw, ch - hh), fill=None, line=col, line_w=1)
            for r, items in enumerate(pl.get("rows", [])):
                ry = y + hh + 0.08 + r * rh
                self.text(s, (cx + 0.08, ry + 0.04, cw - 0.16, rh - 0.08),
                          [{"text": t, "bullet": "•", "size": size - 1, "space_after": 1} for t in items],
                          anchor="m")

    def hub_quadrant(self, s, box, center, cards, size=10):
        """허브 + 4사분면: 가운데 3겹 원(핵심 시스템), 네 모서리 카드(제목·부제·개조식), 꺾은 연결선.
        center: (제목, 부제). cards: 4개 [{title, sub?, bullets, icon?, color?}]"""
        b = self.b
        x, y, w, h = box
        cw, chh = w * 0.36, h * 0.47
        pos = [(x, y), (x + w - cw, y), (x, y + h - chh), (x + w - cw, y + h - chh)]
        d = min(w - 2 * cw - 0.2, h * 0.75)
        cx, cy = x + w / 2, y + h / 2
        for i, (c, (px, py)) in enumerate(zip(cards, pos)):
            col = c.get("color", self.seq(i))
            mid_x = px + cw if i % 2 == 0 else px
            mid_y = py + chh / 2
            self.line(s, mid_x, mid_y, cx + (-d / 2 if i % 2 == 0 else d / 2) * 0.7, cy, col, 1.5, arrow=False)
            self.header_card(s, (px, py, cw, chh), c["title"], c.get("bullets", []), no=f"{i + 1:02d}",
                             sub=c.get("sub"), color=col, icon=c.get("icon"), style="border", size=size + 1)
        self.box(s, (cx - d / 2, cy - d / 2, d, d), fill=b["light"], shape="oval", line=b["line"])
        d2 = d * 0.84
        self.box(s, (cx - d2 / 2, cy - d2 / 2, d2, d2), fill=b["white"], shape="oval", line=b["primary"], line_w=1.5)
        d3 = d * 0.7
        t, sub = (center, "") if isinstance(center, str) else center
        self.box(s, (cx - d3 / 2, cy - d3 / 2, d3, d3),
                 [{"text": t, "bold": True, "size": size + 5, "color": b["white"], "space_after": 1, "align": "c"},
                  {"text": sub, "size": size - 1, "color": "DCE6F5", "align": "c"}], fill=b["primary"], shape="oval")

    def screen_callout(self, s, box, callouts, image=None, markers=None, stamp=True, title="", size=11,
                       menu=None, kpis=None, mode="dashboard"):
        """화면 목업 + 번호 콜아웃: 왼쪽 설명 카드(번호·제목·설명), 오른쪽 화면(이미지 없으면 도형 목업),
        주황 연결선과 화면 위 노란 번호 마커. markers: 화면 안 상대 좌표 [(0~1, 0~1)].
        menu: 목업 왼쪽 메뉴 이름 목록, kpis: 목업 상단 지표 [(값, 라벨)] 3개(가짜 수치 대신 "00" 권장).
        mode: dashboard(차트·도넛, 현황 화면) | list(검색 조건 + 목록 그리드, 관리 화면) | cards(카드형 목록, 고객 화면)."""
        b = self.b
        x, y, w, h = box
        lw = min(2.4, w * 0.3)
        sx, sw = x + lw + 0.5, w - lw - 0.5
        # 화면
        if image and os.path.exists(image):
            self.image(s, (sx, y, sw, h), image)
        else:
            self.box(s, (sx, y, sw, h), fill=b["white"], line="9AA3AE", line_w=1)
            self.box(s, (sx, y, sw, 0.26), f"   ●  ●  ●     {title}", fill="E4E7EB", size=8, color="6B7280",
                     align="l", margin=0.04)
            mw = min(1.2, sw * 0.2)
            self.box(s, (sx, y + 0.26, mw, h - 0.26), fill=b["primary"])
            for k, mn in enumerate(menu or ["대시보드", "조회", "관리", "통계", "설정"]):
                self.box(s, (sx + 0.08, y + 0.4 + k * 0.3, mw - 0.16, 0.24), mn, fill=b["accent"] if k == 0 else None,
                         size=8, color=b["white"], align="l", margin=0.06)
            gx, gy = sx + mw + 0.15, y + 0.4
            gw = sw - mw - 0.3
            for k in range(3):
                kx = gx + k * (gw / 3)
                kv = (kpis or [("00", "지표 1"), ("00", "지표 2"), ("00", "지표 3")])[k]
                self.box(s, (kx, gy, gw / 3 - 0.1, 0.55), [{"text": kv[0], "bold": True,
                         "size": 13, "color": b["primary"], "space_after": 0},
                         {"text": kv[1], "size": 8, "color": b["subtext"]}],
                         fill=b["light"], line=b["line"], margin=0.05)
            body = (gx, gy + 0.68, gw, y + h - (gy + 0.68) - 0.1)
            {"list": self._mock_list, "cards": self._mock_cards}.get(mode, self._mock_dashboard)(s, body)
        n = len(callouts)
        ch = (h - 0.12 * (n - 1)) / n
        markers = markers or [0.3 + 0.15 * i for i in range(n)]
        for i, c in enumerate(callouts):
            cy = y + i * (ch + 0.12)
            if not isinstance(markers[i], (tuple, list)):
                markers[i] = (markers[i], (cy + ch / 2 - y) / h)
            self.box(s, (x, cy, lw, ch), fill=b["white"], line=b["line"], shape="round", radius=0.08)
            self.box(s, (x + 0.08, cy + 0.08, 0.24, 0.24), str(i + 1), fill=b["primary"], shape="oval", size=9,
                     bold=True, color=b["white"], margin=0)
            desc = c.get("desc", "")
            self.text(s, (x + 0.38, cy + 0.06, lw - 0.44, ch - 0.1),
                      [{"text": c["title"], "bold": True, "size": size + 1, "color": b["primary"], "space_after": 2}] +
                      [{"text": t, "bullet": "•" if isinstance(desc, list) else None, "size": size - 1,
                        "color": b["text"], "space_after": 1} for t in (desc if isinstance(desc, list) else [desc])])
            mx, my = sx + sw * markers[i][0], y + h * markers[i][1]
            self.line(s, x + lw, cy + ch / 2, mx - 0.13, my, "F28C28", 1, elbow=abs(cy + ch / 2 - my) > 0.02)
            self.box(s, (mx - 0.13, my - 0.13, 0.26, 0.26), str(i + 1), fill="FFC000", shape="oval", size=9,
                     bold=True, color=b["primary"], margin=0)
        if stamp:
            self.stamp(s, sx + sw - 0.8, y + 0.36)

    def panel(self, s, box, title, no=None, color=None, size=11):
        """번호 패널(수주 제안서 '01 업무보고 체계' 형식): 색 머리띠 + 같은 색 테두리 본문. 반환: 안쪽 영역."""
        b = self.b
        x, y, w, h = box
        col = color or b["primary"]
        head = f"{no}  {title}" if no else title
        self.box(s, (x, y, w, 0.3), head, fill=col, size=size, bold=True, color=b["white"], margin=0.04)
        self.box(s, (x, y + 0.3, w, h - 0.3), fill=b["white"], line=col, line_w=1)
        return (x + 0.08, y + 0.38, w - 0.16, h - 0.46)

    def item_rows(self, s, box, rows, label_w=0.85, size=10, gap=0.06):
        """라벨 상자 + 설명 행 목록(패널 안을 빽빽하게 채우는 기본 행). rows: [(라벨, 설명 str|list)]"""
        b = self.b
        x, y, w, h = box
        gap = gap * self.dg
        n = len(rows)
        rh = (h - gap * (n - 1)) / n
        for i, (lab, desc) in enumerate(rows):
            ry = y + i * (rh + gap)
            self.box(s, (x, ry, label_w, rh), lab, fill=b["light"], line=b["line"], shape="round", radius=0.12,
                     size=size, bold=True, color=b["primary"], margin=0.03)
            lines = desc if isinstance(desc, list) else [desc]
            self.text(s, (x + label_w + 0.08, ry, w - label_w - 0.08, rh),
                      [{"text": t, "size": size - 1, "color": b["text"], "space_after": 1} for t in lines], anchor="m")
            if i < n - 1:
                self.line(s, x + label_w + 0.08, ry + rh + gap / 2, x + w, ry + rh + gap / 2, "E3E7EC", 0.5)

    def icon_rows(self, s, box, rows, size=10, numbered=True):
        """번호 원 + 아이콘 + 제목·설명 행(수주 제안서 '02 이슈/의사결정 체계' 형식). rows: [(아이콘, 제목, 설명)]"""
        b = self.b
        x, y, w, h = box
        n = len(rows)
        rh = h / n
        for i, (ic, title, desc) in enumerate(rows):
            ry = y + i * rh
            dx = x
            if numbered:
                d = min(0.26, rh * 0.6)
                self.box(s, (x, ry + (rh - d) / 2, d, d), str(i + 1), fill=b["primary"], shape="oval", size=9,
                         bold=True, color=b["white"], margin=0)
                dx = x + d + 0.08
            self.box(s, (dx, ry + 0.03, w - (dx - x), rh - 0.06), fill=b["white"], line=b["line"], shape="round",
                     radius=0.1)
            idim = min(0.34, rh - 0.12)
            if ic:
                self.icon(s, (dx + 0.06, ry + (rh - idim) / 2, idim, idim), ic)
            tx = dx + (idim + 0.14 if ic else 0.1)
            self.text(s, (tx, ry + 0.03, x + w - tx - 0.05, rh - 0.06),
                      [{"text": title, "bold": True, "size": size, "color": b["primary"], "space_after": 0},
                       {"text": desc, "size": size - 1.5, "color": b["text"]}], anchor="m")

    def sub_title(self, s, x, y, w, text, size=11, color=None):
        """패널 밖 소제목('04 보고 체계 (Report Matrix)'). 반환: 소제목 아래 y."""
        self.text(s, (x, y, w, 0.26), text, size=size, bold=True, color=color or self.b["primary"], anchor="m")
        return y + 0.28

    # ------------------------------------------------ 수주 제안서·공통 가이드 고밀도 도식
    def icon_flow(self, s, box, items, size=9, arrow=True, color=None):
        """원 안 아이콘 + 아래 라벨을 가로로 잇는 공정 흐름(양부대 '입고예정 ▸ 입고/검수 ▸ …').
        items: [(아이콘, 라벨)]. 반환: 사용한 아래쪽 y."""
        b = self.b
        x, y, w, h = box
        n = len(items)
        aw = 0.16 if arrow else 0
        cw = (w - aw * (n - 1)) / n
        d = min(cw * 0.62, h - 0.28, 0.62)
        for i, (ic, lab) in enumerate(items):
            cx = x + i * (cw + aw)
            self.icon(s, (cx + (cw - d) / 2, y, d, d), ic, circle="ring")
            self.text(s, (cx - 0.04, y + d + 0.03, cw + 0.08, 0.24), lab, size=size, color=b["text"], align="c")
            if arrow and i < n - 1:
                self.box(s, (cx + cw + aw * 0.2, y + d / 2 - 0.05, aw * 0.6, 0.1), fill=color or "8A94A6", shape="tri"
                         ).rotation = 90
        return y + d + 0.3

    def bullet_block(self, s, box, items, size=9, color=None, char="•", highlight=()):
        """개조식 블록(작은 글씨). highlight: 주황 강조할 항목 번호."""
        b = self.b
        paras = [{"text": t, "bullet": char, "size": size, "space_after": 1,
                  "color": "E26B0A" if i in highlight else (color or b["text"])} for i, t in enumerate(items)]
        return self.text(s, box, paras)

    def hub_cards(self, s, box, center, cards, size=10):
        """허브 + 4개 사분면 카드(양부대 '사업범위 내용' 형식).
        center: (제목, [부제 줄...]). cards: 4개 {title, sub, flow:[(아이콘,라벨)], bullets:[...], color?}"""
        b = self.b
        x, y, w, h = box
        cw, ch = w * 0.37, (h - 0.2) / 2
        pos = [(x, y), (x + w - cw, y), (x, y + ch + 0.2), (x + w - cw, y + ch + 0.2)]
        cx, cy = x + w / 2, y + h / 2
        d = min(w - 2 * cw - 0.5, h * 0.62)
        for i, (c, (px, py)) in enumerate(zip(cards, pos)):
            col = c.get("color", self.seq(i))
            self.box(s, (px, py, cw, ch), fill=b["white"], line="D5DAE1", line_w=0.75, shape="round", radius=0.05)
            self.text(s, (px + 0.12, py + 0.06, cw - 0.2, 0.26), f"{i + 1:02d} {c['title']}", size=size + 3, bold=True,
                      color=col, anchor="m")
            self.text(s, (px + 0.3, py + 0.32, cw - 0.4, 0.2), c.get("sub", ""), size=size - 1, bold=True,
                      color=b["text"], anchor="m")
            fy = self.icon_flow(s, (px + 0.12, py + 0.56, cw - 0.24, ch * 0.42), c.get("flow", []), size=size - 1.5)
            self.bullet_block(s, (px + 0.2, fy, cw - 0.3, py + ch - fy - 0.05), c.get("bullets", []), size=size - 2)
            # 허브 → 카드 꺾은 화살표
            left = i % 2 == 0
            ex = px + cw if left else px
            ey = py + ch * 0.18 if i < 2 else py + ch * 0.82
            sx = cx + (-d * 0.35 if left else d * 0.35)
            sy = cy - d * 0.36 if i < 2 else cy + d * 0.36
            self.line(s, sx, sy, sx, ey, col, 2)
            self.line(s, sx, ey, ex + (0.02 if left else -0.02), ey, col, 2, arrow=True)
        self.box(s, (cx - d / 2 - 0.08, cy - d / 2 - 0.08, d + 0.16, d + 0.16), fill=None, line="D5DAE1", shape="oval")
        t, subs = center
        paras = [{"text": t, "bold": True, "size": size + 8, "color": b["white"], "space_after": 6, "align": "c"}]
        paras += [{"text": sline, "size": size - 1, "bold": k > 0, "color": "DCE6F5", "space_after": 2, "align": "c"}
                  for k, sline in enumerate(subs)]
        self.box(s, (cx - d / 2, cy - d / 2, d, d), paras, fill=b["primary"], shape="oval")

    def pillar_detail(self, s, box, pillars, size=9.5):
        """기둥 카드(양부대 '사업 추진 전략' 형식): 색 머리(국문+영문 대문자) → 회색 설명 → 소제목별 개조식.
        pillars: [{title, en, desc, sections: [(소제목, [항목...], [강조 번호])], color?}]"""
        b = self.b
        cells = self.hsplit(box, [1] * len(pillars), 0.14)
        for i, (p, (x, y, w, h)) in enumerate(zip(pillars, cells)):
            col = p.get("color", self.seq(i))
            self.box(s, (x, y, w, 0.55), [{"text": p["title"], "bold": True, "size": size + 3, "color": b["white"],
                                           "space_after": 0, "align": "c"},
                                          {"text": p.get("en", "").upper(), "bold": True, "size": size, "color": b["white"],
                                           "align": "c"}], fill=col, margin=0.03)
            self.box(s, (x, y + 0.55, w, h - 0.55), fill="F7F8FA", line="D5DAE1", line_w=0.75)
            self.text(s, (x + 0.1, y + 0.6, w - 0.2, 0.42), p.get("desc", ""), size=size - 1, color=b["subtext"],
                      align="c", anchor="m")
            secs = p.get("sections", [])
            top = y + 1.06
            avail = y + h - top - 0.06
            weights = [len(sc[1]) + 1.3 for sc in secs]
            for k, sc in enumerate(secs):
                name, items = sc[0], sc[1]
                hl = sc[2] if len(sc) > 2 else ()
                sh = avail * weights[k] / sum(weights)
                if k:
                    self.line(s, x + 0.12, top - 0.02, x + w - 0.12, top - 0.02, "D5DAE1", 0.75)
                self.text(s, (x + 0.14, top, w - 0.2, 0.22), name, size=size, bold=True, color=b["text"], anchor="m")
                self.bullet_block(s, (x + 0.2, top + 0.22, w - 0.3, sh - 0.24), items, size=size - 1.5, highlight=hl)
                top += sh

    def flow_cards(self, s, box, items, size=10):
        """하단 흐름 카드: [아이콘 원 + 색 제목 + 설명] ➜ … (양부대 '업무 연속성 확보 ➜ …')."""
        b = self.b
        x, y, w, h = box
        self.box(s, box, fill=b["white"], line="D5DAE1", line_w=0.75, shape="round", radius=0.08)
        n = len(items)
        aw = 0.32
        cw = (w - 0.2 - aw * (n - 1)) / n
        d = min(h - 0.24, 0.5)
        for i, it in enumerate(items):
            cx = x + 0.1 + i * (cw + aw)
            col = it.get("color", self.seq(i))
            self.icon(s, (cx, y + (h - d) / 2, d, d), it.get("icon"), circle=col)
            self.text(s, (cx + d + 0.08, y + 0.08, cw - d - 0.1, h - 0.16),
                      [{"text": it["title"], "bold": True, "size": size + 1, "color": col, "space_after": 1},
                       {"text": it.get("desc", ""), "size": size - 2, "color": b["text"]}], anchor="m")
            if i < n - 1:
                self.box(s, (cx + cw + 0.05, y + h / 2 - 0.1, aw - 0.1, 0.2), fill=b["accent"], shape="arrow")

    def col_summary(self, s, box, items, size=10):
        """하단 열 요약(양부대 '정확한 재고관리 | 업무 자동화 | …'): 색 제목 가운데 + 개조식 2~3줄."""
        b = self.b
        x, y, w, h = box
        self.box(s, box, fill=b["white"], line="D5DAE1", line_w=0.75, shape="round", radius=0.06)
        n = len(items)
        cw = w / n
        for i, it in enumerate(items):
            cx = x + i * cw
            if i:
                self.line(s, cx, y + 0.12, cx, y + h - 0.12, "D5DAE1", 0.75)
            col = it.get("color", self.seq(i))
            self.text(s, (cx + 0.1, y + 0.08, cw - 0.2, 0.3), it["title"], size=size + 1, bold=True, color=col,
                      align="c", anchor="m")
            self.bullet_block(s, (cx + 0.2, y + 0.4, cw - 0.3, h - 0.46), it.get("bullets", []), size=size - 1.5)

    def tile_row(self, s, box, title, tiles, size=9.5, color=None):
        """남색 제목 띠 + 아이콘 타일 행(양부대 '기술 기반'). tiles: [(아이콘, 굵은 라벨, 작은 설명)]"""
        b = self.b
        x, y, w, h = box
        self.box(s, (x, y, w, 0.28), title, fill=color or b["primary"], size=size + 1, bold=True, color=b["white"],
                 margin=0.02)
        n = len(tiles)
        cells = self.hsplit((x, y + 0.36, w, h - 0.36), [1] * n, 0.1)
        for (ic, lab, sub), (cx, cy, cw, chh) in zip(tiles, cells):
            self.box(s, (cx, cy, cw, chh), fill=b["white"], line="9AA3AE", line_w=0.75)
            d = min(0.36, chh * 0.42)
            self.icon(s, (cx + (cw - d) / 2, cy + 0.06, d, d), ic)
            self.text(s, (cx + 0.03, cy + d + 0.08, cw - 0.06, chh - d - 0.1),
                      [{"text": lab, "bold": True, "size": size, "color": b["text"], "space_after": 0, "align": "c"},
                       {"text": sub, "size": size - 1.5, "color": b["subtext"], "align": "c"}], align="c")

    def lane_stack(self, s, box, lanes, ratios=None, arrows=None, label_w=1.05, size=11):
        """세로 3단 레인(양부대 '스마트 병원 통합 물류 관리 체계' 형식).
        왼쪽에 `01 기존 시스템 연계` 같은 색 라벨, 오른쪽 각 단의 작업 영역을 돌려준다.
        lanes: [(번호, 제목(줄바꿈 가능), 색)] · arrows: 단 사이 설명(아래 화살표). 반환: 단별 영역 목록."""
        b = self.b
        x, y, w, h = box
        n = len(lanes)
        ratios = ratios or [1] * n
        ah = 0.36 if arrows else 0.12
        avail = h - ah * (n - 1)
        tot = sum(ratios)
        out, cy = [], y
        for i, (no, title, col) in enumerate(lanes):
            lh = avail * ratios[i] / tot
            parts = title.splitlines()
            self.text(s, (x, cy, label_w - 0.1, lh),
                      [{"text": f"{no} {parts[0]}", "bold": True, "size": size + 2, "color": col, "space_after": 0}] +
                      [{"text": t, "bold": True, "size": size + 2, "color": col, "space_after": 0} for t in parts[1:]],
                      anchor="m")
            out.append((x + label_w, cy, w - label_w, lh))
            if i < n - 1:
                ly = cy + lh
                self.line(s, x, ly + ah / 2, x + w, ly + ah / 2, "D5DAE1", 0.75)
                cap = arrows[i] if arrows and i < len(arrows) else None
                if cap:
                    tw = min(w * 0.5, max(len(cap.splitlines()[0]) * 0.08 + 0.6, 1.8))
                    cx = x + label_w + (w - label_w) / 2
                    self.text(s, (cx - tw / 2 - 0.25, ly + 0.01, tw, ah - 0.04),
                              [{"text": t, "bold": k == 0, "size": size - (1 if k == 0 else 2),
                                "color": b["text"] if k == 0 else b["subtext"], "space_after": 0, "align": "c"}
                               for k, t in enumerate(cap.splitlines())], anchor="m")
                    self.box(s, (cx + tw / 2 - 0.15, ly + 0.05, 0.28, ah - 0.12), fill="9AA3AE", shape="down")
            cy += lh + ah
        return out

    def feature_row(self, s, box, cells, size=10, border=True):
        """한 상자를 세로선으로 나눈 [아이콘 + 굵은 제목 + 개조식] 칸. cells: [{icon, title, bullets}]"""
        b = self.b
        x, y, w, h = box
        if border:
            self.box(s, box, fill=b["white"], line="D5DAE1", line_w=0.75)
        n = len(cells)
        cw = w / n
        d = min(h * 0.42, 0.44)
        for i, c in enumerate(cells):
            cx = x + i * cw
            if i:
                self.line(s, cx, y + 0.1, cx, y + h - 0.1, "E3E7EC", 0.75)
            self.icon(s, (cx + 0.12, y + 0.12, d, d), c.get("icon"))
            tx = cx + d + 0.24
            self.text(s, (tx, y + 0.08, cw - (tx - cx) - 0.1, h - 0.16),
                      [{"text": c["title"], "bold": True, "size": size + 1, "color": b["primary"], "space_after": 2}] +
                      [{"text": t, "bullet": "•", "size": size - 1, "color": b["text"], "space_after": 1}
                       for t in c.get("bullets", [])])

    def tile_grid(self, s, box, tiles, label=None, cols=None, highlight=None, size=10, dashed=True):
        """점선 영역 안 아이콘 타일 묶음(양부대 '통합 WMS' 단). tiles: [{icon, title, bullets?}]"""
        b = self.b
        x, y, w, h = box
        if label:
            self.box(s, (x, y + 0.11, w, h - 0.11), fill=None, line="AEB6C0", line_w=0.75, dash=dashed)
            tw = 0.25 + len(label) * 0.11
            self.box(s, (x + 0.12, y, tw, 0.24), label, fill=b["white"], size=size + 1, bold=True,
                     color=b["secondary"], margin=0.02)
            x, y, w, h = x + 0.1, y + 0.34, w - 0.2, h - 0.44
        cols = cols or len(tiles)
        rows = -(-len(tiles) // cols)
        cells = [c for r in self.vsplit((x, y, w, h), [1] * rows, 0.12 * self.dg)
                 for c in self.hsplit(r, [1] * cols, 0.1 * self.dg)]
        for i, (t, (cx, cy, cw, ch)) in enumerate(zip(tiles, cells)):
            hl = i == highlight
            self.box(s, (cx, cy, cw, ch), fill=b["white"], line=b["accent"] if hl else "C9CED6",
                     line_w=1.5 if hl else 0.75, shape="round", radius=0.06)
            d = min(ch * 0.38, cw * 0.5, 0.4)
            self.icon(s, (cx + (cw - d) / 2, cy + 0.06, d, d), t.get("icon"))
            ty = cy + d + 0.1
            lines = t["title"].splitlines()
            th = min(0.2 * len(lines) + 0.06, cy + ch - ty)
            self.text(s, (cx + 0.03, ty, cw - 0.06, th),
                      [{"text": ln, "bold": True, "size": size, "color": b["primary"], "space_after": 0, "align": "c"}
                       for ln in lines], align="c")
            rest = cy + ch - (ty + th) - 0.04
            if t.get("bullets") and rest > 0.18:
                self.bullet_block(s, (cx + 0.06, ty + th, cw - 0.12, rest), t["bullets"], size=size - 2, char="-")

    def group_columns(self, s, box, groups, size=10):
        """남색 머리 + 아이콘·개조식 상자 열(양부대 '물류 실행 영역' 단). groups: [{title, icon, bullets}]"""
        b = self.b
        cells = self.hsplit(box, [1] * len(groups), 0.12 * self.dg)
        for g, (x, y, w, h) in zip(groups, cells):
            self.box(s, (x, y, w, 0.3), g["title"], fill=g.get("color", b["primary"]), size=size + 1, bold=True,
                     color=b["white"], margin=0.02)
            self.box(s, (x, y + 0.34, w, h - 0.34), fill=b["white"], line="D5DAE1", line_w=0.75)
            d = min(h * 0.3, w * 0.26, 0.48)
            self.icon(s, (x + 0.09, y + 0.46, d, d), g.get("icon"))
            tx = x + d + 0.16
            self.bullet_block(s, (tx, y + 0.4, x + w - tx - 0.06, h - 0.46), g.get("bullets", []), size=size - 1,
                              char="-")

    def chain(self, s, box, steps, caption=None, size=10):
        """캡션 줄(— 문장 —) + 둥근 상자 ➜ 상자 연결(양부대 '보유기술 ➜ 시스템 구현 ➜ …').
        steps: [(굵은 제목, 괄호 부제)]"""
        b = self.b
        x, y, w, h = box
        if caption:
            tw = min(w * 0.7, len(caption) * 0.1 + 0.4)
            self.line(s, x, y + 0.13, x + (w - tw) / 2 - 0.1, y + 0.13, "9AA3AE", 0.75)
            self.line(s, x + (w + tw) / 2 + 0.1, y + 0.13, x + w, y + 0.13, "9AA3AE", 0.75)
            self.text(s, (x + (w - tw) / 2, y, tw, 0.26), caption, size=size, bold=True, color=b["secondary"],
                      align="c", anchor="m")
            y, h = y + 0.36, h - 0.36
        n = len(steps)
        aw = 0.4
        cw = (w - aw * (n - 1)) / n
        for i, (t, sub) in enumerate(steps):
            cx = x + i * (cw + aw)
            col = self.seq(i) if i < n - 1 else b["positive"]
            self.box(s, (cx, y, cw, h), [{"text": t, "bold": True, "size": size, "color": col, "space_after": 0,
                                          "align": "c"},
                                         {"text": f"({sub})" if sub else "", "bold": True, "size": size - 1,
                                          "color": col, "align": "c"}],
                     fill=b["white"], line="C9CED6", line_w=0.75, shape="round", radius=0.12)
            if i < n - 1:
                self.box(s, (cx + cw + 0.08, y + h / 2 - 0.1, aw - 0.16, 0.2), fill=b["accent"], shape="arrow")

    def photo(self, s, box, path=None, caption="[자산 필요: 현장·화면 사진]"):
        """사진 자리. 경로가 있으면 이미지, 없으면 회색 자리표시와 [자산 필요] 문구."""
        if path and os.path.exists(path):
            return self.image(s, box, path)
        return self.box(s, box, caption, fill="E4E7EB", line="C9CED6", size=10, color="6B7280")

    def _mock_dashboard(self, s, body):
        b = self.b
        gx, gy, gw, gh = body
        ch_h = gh * 0.58
        cw1 = gw * 0.6
        self.box(s, (gx, gy, cw1, ch_h), fill=b["white"], line=b["line"])
        nb = 7
        for k in range(nb):
            bh = ch_h * (0.25 + 0.09 * k)
            self.box(s, (gx + 0.15 + k * (cw1 - 0.3) / nb, gy + ch_h - 0.08 - bh,
                         (cw1 - 0.3) / nb * 0.6, bh), fill=self.seq(1) if k < nb - 1 else b["accent"])
        self.box(s, (gx + cw1 + 0.1, gy, gw - cw1 - 0.1, ch_h), fill=b["white"], line=b["line"])
        dd = min(ch_h * 0.7, (gw - cw1 - 0.1) * 0.8)
        self.box(s, (gx + cw1 + 0.1 + (gw - cw1 - 0.1 - dd) / 2, gy + (ch_h - dd) / 2, dd, dd),
                 fill=None, line=self.seq(2), line_w=6, shape="oval")
        ty = gy + ch_h + 0.12
        self.box(s, (gx, ty, gw, 0.22), fill="E4E7EB")
        for k in range(max(int((gy + gh - ty - 0.22) / 0.2), 1)):
            self.line(s, gx, ty + 0.22 + (k + 1) * 0.2, gx + gw, ty + 0.22 + (k + 1) * 0.2, b["line"], 0.5)

    def _mock_list(self, s, body):
        """관리 화면 목업: 검색 조건 줄 + 목록 그리드(체크박스·상태 배지)."""
        b = self.b
        gx, gy, gw, gh = body
        self.box(s, (gx, gy, gw, 0.34), "  검색 조건 ▾      기간 ▾      상태 ▾", fill=b["light"], line=b["line"],
                 size=8, color=b["subtext"], align="l", margin=0.05)
        self.box(s, (gx + gw - 0.7, gy + 0.05, 0.62, 0.24), "조회", fill=b["accent"], size=8, color=b["white"], margin=0)
        ty = gy + 0.46
        self.box(s, (gx, ty, gw, 0.26), fill="E4E7EB")
        for k in range(max(int((gy + gh - ty - 0.3) / 0.28), 3)):
            ry = ty + 0.26 + 0.28 * k
            self.line(s, gx, ry + 0.28, gx + gw, ry + 0.28, b["line"], 0.5)
            self.box(s, (gx + 0.08, ry + 0.08, 0.12, 0.12), fill=None, line="9AA3AE", line_w=0.75)
            for c, cw in ((0.1, 0.14), (0.3, 0.16), (0.52, 0.12)):
                self.box(s, (gx + gw * c, ry + 0.11, gw * cw, 0.06), fill="D5DAE1")
            self.box(s, (gx + gw * 0.74, ry + 0.06, 0.5, 0.16), "", fill=(b["accent"], self.seq(1), "9AA3AE")[k % 3],
                     shape="round", radius=0.5)

    def _mock_cards(self, s, body):
        """고객 화면 목업: 차량 카드 그리드(이미지 자리·배지·가격·관심 하트)."""
        b = self.b
        gx, gy, gw, gh = body
        ncol, nrow = 3, 2
        cw = (gw - 0.2) / ncol
        chh = (gh - 0.1) / nrow
        tags = [("-5%", b["negative"]), ("00:42", b["primary"]), ("즉시체결", b["accent"]), ("Blind", "6B7280")]
        for r in range(nrow):
            for c in range(ncol):
                x0, y0 = gx + c * (cw + 0.1), gy + r * (chh + 0.1)
                self.box(s, (x0, y0, cw, chh), fill=b["white"], line=b["line"], shape="round", radius=0.06)
                self.box(s, (x0 + 0.06, y0 + 0.06, cw - 0.12, chh * 0.45), fill="DCE3EC")
                t, col = tags[(r * ncol + c) % 4]
                self.box(s, (x0 + 0.1, y0 + 0.1, 0.55, 0.18), t, fill=col, size=7, color=b["white"], margin=0,
                         shape="round", radius=0.5)
                self.box(s, (x0 + 0.08, y0 + chh * 0.56, cw * 0.6, 0.07), fill="C5CCD6")
                self.box(s, (x0 + 0.08, y0 + chh * 0.56 + 0.14, cw * 0.4, 0.07), fill="C5CCD6")
                self.text(s, (x0 + 0.08, y0 + chh - 0.32, cw - 0.4, 0.26), "0,000만원", size=10, bold=True,
                          color=b["primary"])
                self.text(s, (x0 + cw - 0.32, y0 + chh - 0.32, 0.24, 0.26), "♡", size=10, color=b["negative"])

    def cycle_ring(self, s, box, core, nodes, size=11):
        """순환 링: 가운데 코어 원, 12·3·6·9시 방향 카드 4개. 장(章)·제안서의 결론 장표."""
        b = self.b
        x, y, w, h = box
        d = min(h * 0.82, w * 0.4)
        cx, cy = x + w / 2, y + h / 2
        self.box(s, (cx - d / 2, cy - d / 2, d, d), fill=None, line=self.seq(3), line_w=4.5, shape="oval")
        dc = d * 0.52
        t, sub = (core, "") if isinstance(core, str) else core
        self.box(s, (cx - dc / 2, cy - dc / 2, dc, dc),
                 [{"text": t, "bold": True, "size": size + 4, "color": b["white"], "space_after": 1, "align": "c"},
                  {"text": sub, "size": size - 1, "color": "DCE6F5", "align": "c"}], fill=b["primary"], shape="oval")
        cw, chh = min(2.7, w * 0.3), 0.95
        spots = [(cx - cw / 2, cy - d / 2 - chh * 0.45), (cx + d / 2 - cw * 0.25, cy - chh / 2),
                 (cx - cw / 2, cy + d / 2 - chh * 0.55), (cx - d / 2 - cw * 0.75, cy - chh / 2)]
        for i, (nd, (px, py)) in enumerate(zip(nodes, spots)):
            col = nd.get("color", self.seq(i))
            self.box(s, (px, py, cw, chh), fill=b["white"], line=col, line_w=1.5, shape="round", radius=0.12)
            self.icon(s, (px + 0.1, py + (chh - 0.46) / 2, 0.46, 0.46), nd.get("icon"))
            self.text(s, (px + 0.62, py + 0.06, cw - 0.7, chh - 0.12),
                      [{"text": nd["title"], "bold": True, "size": size + 2, "color": col, "space_after": 1},
                       {"text": nd.get("desc", ""), "size": size - 1, "color": b["text"]}], anchor="m")

    def label_desc_rows(self, s, box, rows, label_w=None, size=11):
        """라벨 pill + 회색 설명 바 행 목록(전략1·2·3, 원칙 등). rows: [(라벨, 설명)]"""
        b = self.b
        x, y, w, h = box
        lw = label_w or min(w * 0.3, 1.6)
        n = len(rows)
        rh = min((h - 0.1 * (n - 1)) / n, 0.44)
        gap = min((h - rh * n) / max(n - 1, 1), 0.16) if n > 1 else 0
        y = y + (h - (rh * n + gap * (n - 1))) / 2
        for i, (lab, desc) in enumerate(rows):
            ry = y + i * (rh + gap)
            col = self.seq(i)
            self.box(s, (x, ry, lw, rh), lab, fill=b["white"], line=col, line_w=1.5, shape="round", radius=0.5,
                     size=size, bold=True, color=col)
            self.box(s, (x + lw + 0.1, ry, w - lw - 0.1, rh), desc, fill=b["light"],
                     size=size - 1, color=b["text"], align="l", margin=0.12)

    def step_rail(self, s, box, steps, size=11, color=None):
        """체크 스텝 레일: 옅은 레일 위에 [색 머리 + ✓ 원 + 2줄 설명] 카드, 사이 화살표."""
        b = self.b
        x, y, w, h = box
        col = color or b["primary"]
        n = len(steps)
        lines = max(len(st.get("desc")) if isinstance(st.get("desc"), list) else 1 for st in steps)
        need = 0.78 + lines * size * b["text_scale"] / 72 * 1.6 + 0.12
        if h > need:
            y, h = y + (h - need) / 2, need
            box = (x, y, w, h)
        self.box(s, (x, y + h * 0.42, w, h * 0.16), fill="DCE6F2")
        cells = self.hsplit(box, [1] * n, 0.18)
        for i, (st, (cx, cy, cw, ch)) in enumerate(zip(steps, cells)):
            self.box(s, (cx, cy, cw, ch), fill=b["white"], line=col, line_w=1, shape="round", radius=0.06)
            self.box(s, (cx, cy, cw, 0.3), f"{i + 1:02d}  {st['title']}", fill=col, size=size + 1, bold=True,
                     color=b["white"], margin=0.02)
            self.box(s, (cx + cw / 2 - 0.15, cy + 0.36, 0.3, 0.3), "✓", fill=b["white"], line=col, shape="oval",
                     size=size + 2, bold=True, color=col, margin=0)
            desc = st.get("desc", "")
            paras = [{"text": t, "size": size - 1, "align": "c", "space_after": 2}
                     for t in (desc if isinstance(desc, list) else [desc])]
            self.text(s, (cx + 0.06, cy + 0.72, cw - 0.12, ch - 0.76), paras, color=b["text"], align="c")
            if i < n - 1:
                self.box(s, (cx + cw + 0.02, cy + ch / 2 - 0.08, 0.14, 0.16), fill=col, shape="tri").rotation = 90

    def pyramid(self, s, box, levels, size=11, width_ratio=0.45):
        """등급 피라미드(장애 등급·우선순위): 위가 좁은 단, 오른쪽에 설명과 담당.
        levels: [{name, desc, owner?}] 위에서 아래로."""
        b = self.b
        x, y, w, h = box
        pw = w * width_ratio
        n = len(levels)
        lh = min((h - 0.06 * (n - 1)) / n, 1.2)
        y = y + (h - (lh * n + 0.06 * (n - 1))) / 2
        for i, lv in enumerate(levels):
            ly = y + i * (lh + 0.06)
            ww = pw * (0.35 + 0.65 * (i + 1) / n)
            col = lerp(b["primary"], "8FA9CF", i / max(n - 1, 1))
            self.box(s, (x + (pw - ww) / 2, ly, ww, lh), lv["name"], fill=col, shape="rect", size=size + 1,
                     bold=True, color=b["white"] if i < n - 1 else b["primary"])
            self.line(s, x + (pw + ww) / 2 + 0.05, ly + lh / 2, x + pw + 0.35, ly + lh / 2, b["line"], 1, dash=True)
            paras = [{"text": lv.get("desc", ""), "size": size - 1, "color": b["text"], "space_after": 1}]
            if lv.get("owner"):
                paras.insert(0, {"text": lv["owner"], "bold": True, "size": size, "color": col, "space_after": 1})
            self.text(s, (x + pw + 0.4, ly, w - pw - 0.4, lh), paras, anchor="m")

    # -------------------------------------------------------- 템플릿 원형 복제
    def _find_shape(self, slide, name):
        base, _, nth = name.partition("#")
        hits = [sh for sh in slide.shapes if sh.name == base]
        k = int(nth) - 1 if nth else 0
        return hits[k] if len(hits) > k else None

    def _set_text_keep_style(self, shp, value):
        """도형의 첫 문단·첫 글자 서식을 유지한 채 글자만 바꾼다. value: str | list[str]"""
        lines = value if isinstance(value, list) else str(value).split("\n")
        lines = [ln for ln in lines if ln is not None]
        body = shp.text_frame._txBody
        ps = body.findall(qn("a:p"))
        tmpl = copy.deepcopy(ps[0]) if ps else etree.Element(qn("a:p"))
        for p_ in ps:
            body.remove(p_)
        for ln in lines or [""]:
            np_ = copy.deepcopy(tmpl)
            runs = np_.findall(qn("a:r"))
            for extra in runs[1:] + np_.findall(qn("a:br")) + np_.findall(qn("a:fld")):
                np_.remove(extra)
            if runs:
                r0 = runs[0]
                np_.remove(r0)
                at = list(np_).index(np_.find(qn("a:endParaRPr"))) if np_.find(qn("a:endParaRPr")) is not None else len(np_)
                for part in [x for x in re.split(r"(\*\*.+?\*\*)", ln) if x] or [""]:
                    r = copy.deepcopy(r0)
                    hl = part.startswith("**") and part.endswith("**") and len(part) > 4
                    r.find(qn("a:t")).text = part[2:-2] if hl else part
                    if hl:
                        rpr = r.find(qn("a:rPr"))
                        if rpr is None:
                            rpr = etree.Element(qn("a:rPr"))
                            r.insert(0, rpr)
                        for old in rpr.findall(qn("a:solidFill")):
                            rpr.remove(old)
                        sf = etree.Element(qn("a:solidFill"))
                        etree.SubElement(sf, qn("a:srgbClr")).set("val", self.b["accent"])
                        rpr.insert(0, sf)
                        if self.b.get("font_bold"):
                            for tag in ("a:latin", "a:ea", "a:cs"):
                                el = rpr.find(qn(tag))
                                if el is None:
                                    el = etree.SubElement(rpr, qn(tag))
                                el.set("typeface", self.b["font_bold"])
                        else:
                            rpr.set("b", "1")
                    np_.insert(at, r)
                    at += 1
            else:
                r = etree.SubElement(np_, qn("a:r"))
                etree.SubElement(r, qn("a:t")).text = ln
                end = np_.find(qn("a:endParaRPr"))
                if end is not None:
                    np_.remove(end)
                    np_.append(end)
            body.append(np_)

    def clone(self, kind, texts=None, layout=None, autonum=None):
        """템플릿의 원형 장표(kind)를 복제하고 역할(role)별 글자를 바꾼다.
        spec에 `"keep_only": True`를 주면 roles에 없는 도형(작업 중 표·화면 캡처 등)은 복제하지 않는다."""
        spec = self.proto[kind]
        src = self._protos[kind]
        keep = ({n.split("#")[0] for n in spec.get("roles", {}).values()}
                if spec.get("keep_only") else None)
        lo = self._layout(layout) if layout else src.slide_layout
        s = self.prs.slides.add_slide(lo)
        for ph in list(s.placeholders):
            ph._element.getparent().remove(ph._element)
        tree = s.shapes._spTree
        ext = tree.find(qn("p:extLst"))
        for shp in src.shapes:
            if keep is not None and shp.name not in keep:
                continue
            el = copy.deepcopy(shp._element)
            for blip in el.iter(qn("a:blip")):
                rid = blip.get(qn("r:embed"))
                if rid:
                    blip.set(qn("r:embed"), s.part.relate_to(src.part.related_part(rid), RT.IMAGE))
            if ext is not None:
                ext.addprevious(el)
            else:
                tree.append(el)
            nv = next(el.iter(qn("p:cNvPr")), None)
            if nv is not None:
                self._tpl_shapes.add((s.slide_id, nv.get("id")))     # 템플릿에서 온 도형(줄 간격 고정 제외)
        for role, val in (texts or {}).items():
            name = spec.get("roles", {}).get(role)
            if not name:
                continue
            shp = self._find_shape(s, name)
            if shp is not None and shp.has_text_frame:
                self._set_text_keep_style(shp, val)
                n = (autonum or {}).get(role)
                if n:
                    # 자동 번호(Ⅰ. Ⅱ. …)를 쓰는 글상자는 시작 번호를 장 번호로 맞춘다
                    found = False
                    for bu in shp.text_frame._txBody.iter(qn("a:buAutoNum")):
                        bu.set("startAt", str(n))
                    for p_ in shp.text_frame._txBody.findall(qn("a:p")):  # 문단 자체에 자동 번호가 있을 때만 인정
                        if p_.find(qn("a:pPr")) is not None and p_.find(qn("a:pPr")).find(qn("a:buAutoNum")) is not None:
                            found = True
                    if not found and not re.match(r"^[ⅠⅡⅢⅣⅤⅥ]", str(val).strip()):
                        # 자동 번호가 없는 템플릿이면 장 번호를 글자에 직접 붙인다
                        rn = {v: k for k, v in ROMANS.items()}.get(n, str(n))
                        self._set_text_keep_style(shp, f"{rn}. {val}")
        self.page += 1
        return s

    # ---------------------------------------------------------------- 도식
    def bullets(self, s, box, items, size=14, title=None, char="•", color=None):
        """개조식 본문. items: [str | (str, level)]"""
        paras = []
        if title:
            paras.append({"text": title, "bold": True, "size": size + 1,
                          "color": self.b["primary"], "space_after": 6})
        for it in items:
            t, lv = (it, 0) if isinstance(it, str) else it
            paras.append({"text": t, "bullet": char if lv == 0 else "–", "level": lv,
                          "size": size if lv == 0 else size - 1.5, "space_after": 5})
        return self.text(s, box, paras, color=color)

    def kpi_cards(self, s, box, items, cols=None, value_size=30):
        """숫자 강조 카드. items: [{value, label, note?, color?}]"""
        b = self.b
        cols = cols or len(items)
        rows = (len(items) + cols - 1) // cols
        cells = [c for r in self.vsplit(box, [1] * rows, 0.2) for c in self.hsplit(r, [1] * cols, 0.2)]
        for it, cell in zip(items, cells):
            col = it.get("color", b["secondary"])
            x, y, w, h = cell
            self.box(s, (x, y, w, 0.07), fill=col)
            paras = [{"text": it["value"], "size": value_size, "bold": True, "color": col,
                      "space_after": 2},
                     {"text": it["label"], "size": 13, "bold": True, "color": b["text"]}]
            if it.get("note"):
                paras.append({"text": it["note"], "size": 10.5, "color": b["subtext"]})
            self.box(s, (x, y + 0.07, w, h - 0.07), paras, fill=b["light2"], line=b["line"],
                     align="c", anchor="m", margin=0.12)

    def process(self, s, box, steps, head_h=0.75, numbered=True, colors=None, fill_height=False, desc_size=12,
                head_size=13):
        """단계 흐름. steps: [{title, desc(list|str)?, period?}] — 쉐브론 + 단계별 설명 카드.
        설명 상자는 가장 긴 단계의 내용 높이에 맞춘다(fill_height=True면 영역 끝까지).
        반환값: 사용한 아래쪽 y 좌표. 남는 공간에는 callout·kpi_cards 같은 보조 도식을 둔다."""
        b = self.b
        n = len(steps)
        x, y, w, h = box
        cw = w / n
        c1, c2 = colors or (b["primary"], b["secondary"])
        longest = max(([d] if isinstance(d, str) else d for d in
                       (st.get("desc") or [] for st in steps)), key=len, default=[])
        need = max(self.est_h(longest, cw - 0.45, desc_size * b["text_scale"], extra=0.12), 0.4)
        bottom = y + head_h
        for i, st in enumerate(steps):
            col = lerp(c1, c2, i / max(n - 1, 1))
            label = f"{i + 1:02d}  {st['title']}" if numbered else st["title"]
            shp = "pentagon" if i == 0 else "chevron"
            self.box(s, (x + i * cw - (0 if i == 0 else 0.12), y, cw + (0.1 if i == 0 else 0.12), head_h),
                     label, fill=col, shape=shp, size=head_size, bold=True, color=b["white"], margin=0.25)
            if st.get("period"):
                self.text(s, (x + i * cw, y + head_h + 0.05, cw - 0.1, 0.3), st["period"],
                          size=10, color=b["subtext"], align="c")
            desc = st.get("desc")
            if desc:
                top = y + head_h + (0.38 if st.get("period") else 0.15)
                items = desc if isinstance(desc, list) else [desc]
                paras = [{"text": t, "bullet": "•", "size": desc_size, "space_after": 5} for t in items]
                bh = y + h - top if fill_height else min(y + h - top, need)
                self.box(s, (x + i * cw + 0.04, top, cw - 0.2, bh), paras,
                         fill=b["light2"], line=b["line"], align="l", anchor="t", margin=0.12)
                bottom = max(bottom, top + bh)
        return bottom

    def timeline(self, s, box, periods, rows, label_w=2.3, bands=None, today=None):
        """간트·로드맵.
        periods: ['M', 'M+1', ...]
        rows: [{label, start, end, text?, color?, milestones?: [(pos, 'label')], group?: bool}]
              start/end 는 period 인덱스(소수 가능, end 포함 안 함: 0~2 는 두 칸)
        bands: [{label, start, end, color?}] 상단 단계 띠(예: 1단계 오픈)
        """
        b = self.b
        x, y, w, h = box
        gx, gw = x + label_w, w - label_w
        pw = gw / len(periods)
        top = y
        if bands:
            for bd in bands:
                self.box(s, (gx + bd["start"] * pw, top, (bd["end"] - bd["start"]) * pw - 0.04, 0.32),
                         bd["label"], fill=bd.get("color", b["primary"]), size=10.5, bold=True,
                         color=b["white"])
            top += 0.38
        for i, p in enumerate(periods):
            self.box(s, (gx + i * pw, top, pw - 0.03, 0.34), p, fill=b["light"], size=10.5,
                     bold=True, color=b["primary"])
        top += 0.42
        rh = (y + h - top) / max(len(rows), 1)
        for i in range(len(periods) + 1):
            self.line(s, gx + i * pw - 0.015, top, gx + i * pw - 0.015, y + h, b["line"], 0.5,
                      dash=True)
        for r, row in enumerate(rows):
            ry = top + r * rh
            if row.get("group"):
                self.box(s, (x, ry + 0.04, w, rh - 0.08), row["label"], fill=b["light2"], size=11,
                         bold=True, color=b["primary"], align="l", margin=0.12)
                continue
            self.text(s, (x, ry, label_w - 0.1, rh), row["label"], size=11, bold=True,
                      color=b["text"], anchor="m")
            has_ms = bool(row.get("milestones"))
            by, bh = (ry + rh * 0.42, rh * 0.46) if has_ms else (ry + rh * 0.2, rh * 0.6)
            if "start" in row:
                col = row.get("color", b["secondary"])
                bw = (row["end"] - row["start"]) * pw - 0.06
                self.box(s, (gx + row["start"] * pw + 0.03, by, bw, bh),
                         row.get("text", ""), fill=col, shape="round", radius=0.3, size=10,
                         bold=True, color=b["white"], margin=0.06)
            for pos, lab in row.get("milestones", []):
                mx = gx + pos * pw
                self.box(s, (mx - 0.11, by - 0.11, 0.22, 0.22), fill=b["accent"], line=b["white"],
                         line_w=1, shape="diamond")
                self.text(s, (mx - 0.9, ry + 0.02, 1.8, by - ry - 0.14), lab, size=9.5, bold=True,
                          color=b["accent"], align="c", anchor="b")
            self.line(s, x, ry + rh, x + w, ry + rh, b["line"], 0.5)
        if today is not None:
            tx = gx + today * pw
            self.line(s, tx, top - 0.05, tx, y + h, b["negative"], 1.5)

    def table(self, s, box, headers, rows, col_widths=None, size=11, highlight_col=None,
              highlight_rows=(), first_col_bold=True, align=None, header_fill=None, row_h=None):
        """네이티브 표. highlight_col: 권고안 열 등 강조. align: 열별 'l'|'c'|'r'
        row_h: 데이터 행 높이 목록(인치) — row_heights()로 줄 수에 맞춰 여백을 고르게 준다(없으면 같은 높이)."""
        b = self.b
        x, y, w, h = box
        nr, nc = len(rows) + 1, len(headers)
        gs = s.shapes.add_table(nr, nc, Inches(x), Inches(y), Inches(w), Inches(h))
        tbl = gs.table
        tblpr = tbl._tbl.tblPr
        sid = tblpr.find(qn("a:tableStyleId"))
        if sid is None:
            sid = etree.SubElement(tblpr, qn("a:tableStyleId"))
        sid.text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"  # No Style, No Grid
        widths = col_widths or [1] * nc
        tot = sum(widths)
        for i, cw in enumerate(widths):
            tbl.columns[i].width = Inches(w * cw / tot)
        hh = min(h / nr, 0.42)
        rh = (h - hh) / max(nr - 1, 1)
        for r in range(nr):
            tbl.rows[r].height = Inches(hh if r == 0 else (row_h[r - 1] if row_h else rh))
        for c, hd in enumerate(headers):
            cell = tbl.cell(0, c)
            hf = b["accent"] if c == highlight_col else (header_fill or b["primary"])
            self._cell(cell, hd, hf, size, True, b["white"], "c", top=1.0)
        for r, row in enumerate(rows, start=1):
            for c, val in enumerate(row):
                cell = tbl.cell(r, c)
                if c == highlight_col:
                    fill = lerp(b["accent"], "FFFFFF", 0.88)
                elif (r - 1) in highlight_rows:
                    fill = "E8F0FC"
                else:
                    fill = b["white"] if r % 2 else b["light2"]
                bold = (c == 0 and first_col_bold) or c == highlight_col
                al = (align[c] if align else ("l" if c == 0 else "c"))
                self._cell(cell, val, fill, size, bold, b["text"], al)
        return gs

    def _cell(self, cell, val, fill, size, bold, color, align, top=0.25):
        cell.fill.solid()
        cell.fill.fore_color.rgb = rgb(fill)
        cell.margin_left = cell.margin_right = Inches(0.08)
        cell.margin_top = cell.margin_bottom = Inches(0.04)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf = cell.text_frame
        tf.text = ""
        self._fill_text(tf, str(val) if str(val) != "" else " ", size, bold, color, align, "m", 0.08, space_after=0)
        tcpr = cell._tc.get_or_add_tcPr()
        # 회사 가이드: 좌우 선 없음, 위 선 1pt(머리글), 안쪽 가로선 0.25pt
        spec = {"a:lnL": None, "a:lnR": None, "a:lnT": (top, self.b["primary"] if top >= 1 else self.b["line"]),
                "a:lnB": (0.25, self.b["line"])}
        for i, tag in enumerate(("a:lnL", "a:lnR", "a:lnT", "a:lnB")):
            old = tcpr.find(qn(tag))
            if old is not None:
                tcpr.remove(old)
            ln = etree.Element(qn(tag))
            if spec[tag] is None:
                ln.set("w", "0")
                etree.SubElement(ln, qn("a:noFill"))
            else:
                ln.set("w", str(int(Pt(spec[tag][0]))))
                sf = etree.SubElement(ln, qn("a:solidFill"))
                etree.SubElement(sf, qn("a:srgbClr")).set("val", spec[tag][1])
            tcpr.insert(i, ln)  # 스키마 순서: lnL, lnR, lnT, lnB 가 채우기보다 앞

    def compare(self, s, box, left, right, arrow_text="", left_color=None, right_color=None):
        """As-Is → To-Be. left/right: {title, items, sub?}"""
        b = self.b
        L, M, R = self.hsplit(box, [1, 0.28, 1], 0.12)
        for (bx, d, col, txt) in ((L, left, left_color or "8A94A6", b["subtext"]),
                                  (R, right, right_color or b["secondary"], b["text"])):
            x, y, w, h = bx
            self.box(s, (x, y, w, 0.55), d["title"], fill=col, size=15, bold=True,
                     color=b["white"])
            paras = []
            if d.get("sub"):
                paras.append({"text": d["sub"], "bold": True, "size": 14, "color": col,
                              "space_after": 12})
            fs = 14 if len(d["items"]) <= 4 else (13 if len(d["items"]) <= 6 else 12)
            paras += [{"text": t, "bullet": "•", "size": fs, "space_after": 10, "color": txt}
                      for t in d["items"]]
            self.box(s, (x, y + 0.55, w, h - 0.55), paras, fill=b["light2"], line=b["line"],
                     align="l", anchor="t", margin=0.25)
        x, y, w, h = M
        self.box(s, (x, y + h / 2 - 0.55, w, 1.1), arrow_text, fill=b["accent"], shape="arrow",
                 size=12, bold=True, color=b["white"], margin=0.02)

    def layers(self, s, box, layers, side=None, label_w=1.6, side_w=1.5):
        """계층형 아키텍처. layers: [{name, items: [str], color?}] 위에서 아래로.
        side: [{name, items}] 오른쪽 세로 패널(보안·운영관리 등)"""
        b = self.b
        x, y, w, h = box
        main_w = w - (len(side) * (side_w + 0.15) if side else 0)
        lh = (h - 0.12 * (len(layers) - 1)) / len(layers)
        for i, ly in enumerate(layers):
            ly_y = y + i * (lh + 0.12)
            col = ly.get("color", lerp(b["primary"], b["secondary"], i / max(len(layers) - 1, 1)))
            self.box(s, (x, ly_y, label_w, lh), ly["name"], fill=col, size=12, bold=True,
                     color=b["white"])
            self.box(s, (x + label_w, ly_y, main_w - label_w, lh), fill=b["light"])
            items = ly["items"]
            n = len(items)
            iw = (main_w - label_w - 0.2 - 0.12 * (n - 1)) / max(n, 1)
            for j, it in enumerate(items):
                d = it if isinstance(it, dict) else {"text": it}
                hl = d.get("new")
                self.box(s, (x + label_w + 0.1 + j * (iw + 0.12), ly_y + 0.12, iw, lh - 0.24),
                         d["text"], fill=b["white"], line=b["accent"] if hl else col,
                         line_w=1.75 if hl else 0.75, shape="round", radius=0.12, size=10.5,
                         bold=bool(hl), color=b["text"])
        if side:
            sx = x + main_w + 0.15
            for sd in side:
                paras = [{"text": sd["name"], "bold": True, "size": 12, "color": b["white"],
                          "space_after": 8}]
                paras += [{"text": t, "size": 10.5, "color": b["white"], "space_after": 5}
                          for t in sd["items"]]
                self.box(s, (sx, y, side_w, h), paras, fill=sd.get("color", "5B6B82"),
                         anchor="t", margin=0.12)
                sx += side_w + 0.15

    def cards(self, s, box, items, cols=None, numbered=True, head_color=None):
        """카드 그리드. items: [{title, body(list|str), tag?}]"""
        b = self.b
        cols = cols or min(len(items), 4)
        rows = (len(items) + cols - 1) // cols
        cells = [c for r in self.vsplit(box, [1] * rows, 0.25) for c in self.hsplit(r, [1] * cols, 0.25)]
        for i, (it, cell) in enumerate(zip(items, cells)):
            x, y, w, h = cell
            col = head_color or lerp(b["primary"], b["secondary"], i / max(len(items) - 1, 1))
            body = it.get("body", [])
            body = body if isinstance(body, list) else [body]
            paras = [{"text": it["title"], "bold": True, "size": 14, "color": col,
                      "space_after": 8}]
            paras += [{"text": t, "bullet": "•", "size": 11.5, "space_after": 4} for t in body]
            self.box(s, (x, y, w, h), paras, fill=b["white"], line=b["line"], shape="round",
                     radius=0.05, align="l", anchor="t", margin=0.2 if not numbered else 0.2)
            self.box(s, (x, y, w, 0.07), fill=col)
            if numbered:
                self.box(s, (x + w - 0.62, y + 0.2, 0.42, 0.42), f"{i + 1:02d}", fill=col,
                         shape="oval", size=11, bold=True, color=b["white"], margin=0)
            if it.get("tag"):
                self.box(s, (x + 0.2, y + h - 0.45, min(w - 0.4, 2.2), 0.3), it["tag"],
                         fill=b["light"], shape="round", radius=0.5, size=9.5, bold=True,
                         color=col)

    def matrix(self, s, box, items, x_labels=("낮음", "중간", "높음"),
               y_labels=("낮음", "중간", "높음"), x_title="발생 가능성", y_title="영향도"):
        """3x3 리스크 매트릭스. items: [{label, x:0-2, y:0-2}]"""
        b = self.b
        x, y, w, h = box
        gx, gy = x + 0.9, y
        gw, gh = w - 0.9, h - 0.75
        cw, ch = gw / 3, gh / 3
        heat = {0: "E6F4EA", 1: "FFF4DE", 2: "FDE2DC", 3: "F9C9BE", 4: "F4A898"}
        for xi in range(3):
            for yi in range(3):
                labels = [it["label"] for it in items if it["x"] == xi and it["y"] == yi]
                paras = [{"text": t, "bullet": "•", "size": 10, "space_after": 2} for t in labels]
                self.box(s, (gx + xi * cw, gy + (2 - yi) * ch, cw - 0.04, ch - 0.04),
                         paras if paras else "", fill=heat[xi + yi], align="l", anchor="m",
                         margin=0.12)
        for i, lab in enumerate(x_labels):
            self.text(s, (gx + i * cw, gy + gh + 0.02, cw, 0.3), lab, size=10.5,
                      color=b["subtext"], align="c")
        for i, lab in enumerate(y_labels):
            self.text(s, (x + 0.3, gy + (2 - i) * ch, 0.6, ch), lab, size=10.5,
                      color=b["subtext"], align="c", anchor="m")
        self.text(s, (gx, gy + gh + 0.35, gw, 0.35), f"{x_title} →", size=11.5, bold=True,
                  color=b["primary"], align="c")
        self.text(s, (x - 0.05, gy, 0.35, gh), "↑ " + y_title, size=11.5, bold=True,
                  color=b["primary"], align="c", anchor="m")

    def funnel(self, s, box, stages, show_rate=True):
        """퍼널. stages: [{label, value(str), note?}] 위가 가장 넓다."""
        b = self.b
        x, y, w, h = box
        n = len(stages)
        sh = (h - 0.1 * (n - 1)) / n
        bar_w = w * 0.62
        for i, st in enumerate(stages):
            bw = bar_w * (1 - 0.45 * i / max(n - 1, 1))
            bx = x + (bar_w - bw) / 2
            col = lerp(b["secondary"], b["primary"], i / max(n - 1, 1))
            self.box(s, (bx, y + i * (sh + 0.1), bw, sh), [
                {"text": st["value"], "size": 18, "bold": True, "color": b["white"], "space_after": 0},
                {"text": st["label"], "size": 11, "color": "DCE6F5"}], fill=col, shape="round",
                radius=0.1)
            if st.get("note"):
                self.text(s, (x + bar_w + 0.3, y + i * (sh + 0.1), w - bar_w - 0.3, sh), st["note"],
                          size=11.5, color=b["text"], anchor="m")

    def swimlane(self, s, box, lanes, steps, cols, label_w=1.5, arrows=True):
        """스윔레인. steps: [{lane, col, text, kind?('decision'|'end'|'hl')}] 순서대로 화살표 연결"""
        b = self.b
        x, y, w, h = box
        lh = h / len(lanes)
        gx, gw = x + label_w, w - label_w
        cw = gw / cols
        for i, ln in enumerate(lanes):
            ly = y + i * lh
            self.box(s, (x, ly, label_w - 0.05, lh - 0.05), ln, fill=b["primary"], size=11.5,
                     bold=True, color=b["white"])
            self.box(s, (gx, ly, gw, lh - 0.05), fill=b["light2"] if i % 2 else b["light"])
        centers = []
        for st in steps:
            cx = gx + st["col"] * cw + cw / 2
            cy = y + st["lane"] * lh + lh / 2 - 0.025
            bw, bh = min(cw - 0.35, 1.9), min(lh - 0.3, 0.75)
            kind = st.get("kind")
            shp = "diamond" if kind == "decision" else ("oval" if kind == "end" else "round")
            fill = b["accent"] if kind == "hl" else (b["white"] if kind != "end" else b["primary"])
            color = b["white"] if kind in ("hl", "end") else b["text"]
            self.box(s, (cx - bw / 2, cy - bh / 2, bw, bh), st["text"], fill=fill,
                     line=b["secondary"], shape=shp, radius=0.2, size=10, bold=kind == "hl",
                     color=color, margin=0.05)
            centers.append((cx, cy, bw, bh))
        if arrows:
            for (x1, y1, w1, h1), (x2, y2, w2, h2) in zip(centers, centers[1:]):
                if abs(y1 - y2) < 0.01:
                    self.line(s, x1 + w1 / 2, y1, x2 - w2 / 2, y2, b["subtext"], 1.25, arrow=True)
                else:
                    down = y2 > y1
                    self.line(s, x1, y1 + (h1 / 2 if down else -h1 / 2), x2 - w2 / 2 if x2 > x1 else x2,
                              y2, b["subtext"], 1.25, arrow=True, elbow=True)

    def org(self, s, box, root, teams, side=None, fill_height=False):
        """조직도. root: {title, name?}; teams: [{title, members: [str]}]; side: [{title, note}] 점선 박스.
        반환값: 사용한 아래쪽 y 좌표(남는 공간에 R&R 표·callout을 둔다)."""
        b = self.b
        x, y, w, h = box
        rw, rh = 2.6, 0.75
        rx = x + (w - rw) / 2
        self.box(s, (rx, y, rw, rh), [{"text": root["title"], "bold": True, "size": 14,
                                        "color": b["white"], "space_after": 0},
                                       {"text": root.get("name", ""), "size": 11, "color": "DCE6F5"}],
                 fill=b["primary"], shape="round", radius=0.12)
        if side:
            for i, sd in enumerate(side):
                sx = x + (0 if i % 2 == 0 else w - 2.4)
                self.box(s, (sx, y, 2.4, rh), [{"text": sd["title"], "bold": True, "size": 12,
                                                 "space_after": 0},
                                                {"text": sd.get("note", ""), "size": 10,
                                                 "color": b["subtext"]}],
                         fill=b["white"], line=b["subtext"], dash=True, shape="round")
                ax = sx + 2.4 if i % 2 == 0 else sx
                self.line(s, ax, y + rh / 2, rx if i % 2 == 0 else rx + rw, y + rh / 2,
                          b["subtext"], 1, dash=True)
        n = len(teams)
        ty = y + rh + 0.55
        bus = y + rh + 0.27
        self.line(s, rx + rw / 2, y + rh, rx + rw / 2, bus, b["subtext"], 1.25)
        cells = self.hsplit((x, ty, w, h - (ty - y)), [1] * n, 0.2)
        self.line(s, cells[0][0] + cells[0][2] / 2, bus, cells[-1][0] + cells[-1][2] / 2, bus,
                  b["subtext"], 1.25)
        for i, (tm, c) in enumerate(zip(teams, cells)):
            cx, cy, cw, ch = c
            self.line(s, cx + cw / 2, bus, cx + cw / 2, cy, b["subtext"], 1.25)
            col = lerp(b["secondary"], b["primary"], i / max(n - 1, 1))
            self.box(s, (cx, cy, cw, 0.5), tm["title"], fill=col, size=12, bold=True,
                     color=b["white"])
            paras = [{"text": m, "size": 11, "space_after": 4} for m in tm.get("members", [])]
            maxm = max(len(t.get("members", [])) for t in teams)
            mh = ch - 0.5 if fill_height else min(ch - 0.5, maxm * 0.3 + 0.35)
            self.box(s, (cx, cy + 0.5, cw, mh), paras or "", fill=b["light2"],
                     line=b["line"], align="c", anchor="m", margin=0.1)
            bottom = cy + 0.5 + mh
        return bottom

    def chart(self, s, box, kind, categories, series, title=None, number_format="#,##0",
              colors=None, show_values=True, legend=None, font_size=11, value_axis=False):
        """네이티브 차트. kind: column|bar|stacked|line|pie|doughnut. series: {이름: [값...]}"""
        b = self.b
        types = {"column": XL_CHART_TYPE.COLUMN_CLUSTERED, "bar": XL_CHART_TYPE.BAR_CLUSTERED,
                 "stacked": XL_CHART_TYPE.COLUMN_STACKED, "line": XL_CHART_TYPE.LINE_MARKERS,
                 "pie": XL_CHART_TYPE.PIE, "doughnut": XL_CHART_TYPE.DOUGHNUT}
        cd = CategoryChartData()
        cd.categories = categories
        for name, vals in series.items():
            cd.add_series(name, vals)
        x, y, w, h = box
        gf = s.shapes.add_chart(types[kind], Inches(x), Inches(y), Inches(w), Inches(h), cd)
        ch = gf.chart
        ch.font.size = Pt(font_size)
        ch.font.name = b["font"]
        ch.font.color.rgb = rgb(b["text"])
        palette = colors or [b["secondary"], b["accent"], b["positive"], b["primary"], "8A94A6"]
        plot = ch.plots[0]
        if title:
            ch.has_title = True
            ch.chart_title.text_frame.text = title
            r = ch.chart_title.text_frame.paragraphs[0].runs[0]
            self._font(r, font_size + 2, True, b["primary"])
        else:
            ch.has_title = False
        multi = len(series) > 1
        ch.has_legend = legend if legend is not None else (multi or kind in ("pie", "doughnut"))
        if ch.has_legend:
            ch.legend.position = XL_LEGEND_POSITION.BOTTOM
            ch.legend.include_in_layout = False
        if kind in ("pie", "doughnut"):
            s0 = plot.series[0]
            for i, pt in enumerate(s0.points):
                pt.format.fill.solid()
                pt.format.fill.fore_color.rgb = rgb(palette[i % len(palette)])
        else:
            for i, se in enumerate(plot.series):
                col = rgb(palette[i % len(palette)])
                if kind == "line":
                    se.format.line.color.rgb = col
                    se.format.line.width = Pt(2.25)
                    se.smooth = False
                    se.marker.format.fill.solid()
                    se.marker.format.fill.fore_color.rgb = col
                else:
                    se.format.fill.solid()
                    se.format.fill.fore_color.rgb = col
            if kind in ("column", "bar", "stacked"):
                plot.gap_width = 70
                if kind == "stacked":
                    plot.overlap = 100
            va = ch.value_axis
            va.has_major_gridlines = value_axis
            if value_axis:
                va.major_gridlines.format.line.color.rgb = rgb(b["line"])
            va.visible = value_axis
            ca = ch.category_axis
            ca.format.line.color.rgb = rgb(b["line"])
            ca.tick_labels.font.size = Pt(font_size)
        if show_values:
            plot.has_data_labels = True
            dl = plot.data_labels
            dl.number_format = number_format
            dl.number_format_is_linked = False
            dl.font.size = Pt(font_size)
            dl.font.bold = True
            if kind in ("column", "bar", "line"):
                dl.position = XL_LABEL_POSITION.OUTSIDE_END if kind != "line" else XL_LABEL_POSITION.ABOVE
            elif kind == "stacked":
                dl.position = XL_LABEL_POSITION.CENTER
                dl.font.color.rgb = rgb(b["white"])
        return gf

    def callout(self, s, box, text, kind="key"):
        """핵심 메시지 상자. kind: key(주황) | good(청록) | risk(빨강) | info(파랑)"""
        b = self.b
        col = {"key": b["accent"], "good": b["positive"], "risk": b["negative"],
               "info": b["secondary"]}[kind]
        bg = {"key": "FDF0E3", "good": "E3F4F1", "risk": "FBE6E6", "info": "E8F0FC"}[kind]
        x, y, w, h = box
        self.box(s, (x, y, 0.08, h), fill=col)
        self.box(s, (x + 0.08, y, w - 0.08, h), text, fill=bg, size=13, bold=True,
                 color=b["text"], align="l", margin=0.18, hl_color=col)

    def steps_vertical(self, s, box, steps):
        """세로 번호 단계(절차·체크포인트). steps: [{title, desc}]"""
        b = self.b
        x, y, w, h = box
        n = len(steps)
        sh = h / n
        for i, st in enumerate(steps):
            sy = y + i * sh
            self.box(s, (x, sy + 0.08, 0.55, 0.55), str(i + 1), fill=b["secondary"], shape="oval",
                     size=14, bold=True, color=b["white"], margin=0)
            if i < n - 1:
                self.line(s, x + 0.275, sy + 0.63, x + 0.275, sy + sh + 0.08, b["line"], 1.5)
            self.text(s, (x + 0.75, sy + 0.05, w - 0.75, sh - 0.1),
                      [{"text": st["title"], "bold": True, "size": 13, "color": b["primary"],
                        "space_after": 2},
                       {"text": st.get("desc", ""), "size": 11, "color": b["text"]}])


def inspect(path):
    """템플릿 구조 출력: 장표별 도형 이름·위치·글자·서식, 레이아웃 목록. proto 설정을 만들 때 쓴다."""
    prs = Presentation(path)
    print(f"size {prs.slide_width / 914400:.2f} x {prs.slide_height / 914400:.2f} in")
    for mi, m in enumerate(prs.slide_masters):
        for lo in m.slide_layouts:
            txt = [sh.text_frame.text.strip()[:12] for sh in lo.shapes if sh.has_text_frame and sh.text_frame.text.strip()]
            print(f"  master{mi} layout {lo.name!r} {txt[:3]}")
    for i, sl in enumerate(prs.slides, 1):
        print(f"--- slide {i} layout={sl.slide_layout.name!r}")
        for sh in sl.shapes:
            info = ""
            if sh.has_text_frame and sh.text_frame.paragraphs and sh.text_frame.paragraphs[0].runs:
                r = sh.text_frame.paragraphs[0].runs[0]
                info = f"{r.font.name} {r.font.size.pt if r.font.size else ''} b={r.font.bold}"
            t = sh.text_frame.text.strip().replace("\n", " / ")[:40] if sh.has_text_frame else ""
            print(f"  {sh.name!r:28} ({sh.left / 914400:.2f},{sh.top / 914400:.2f},{sh.width / 914400:.2f}x{sh.height / 914400:.2f}) {info} | {t}")


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 2 and sys.argv[1] == "--inspect":
        sys.stdout.reconfigure(encoding="utf-8")
        inspect(sys.argv[2])
