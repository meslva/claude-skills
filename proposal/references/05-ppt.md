# 5. PPT 제작

## 원칙

- **관문 C(원고 승인) 후에 만든다.** 원고가 바뀌면 PPT를 다시 만들어야 하므로, 확정 원고를 기준으로 한다. 급하면 승인된 장부터 만든다.
- **장표마다 도식을 먼저 고른다.** 원고의 `도식` 필드가 기준이다. 패턴은 `references/visual-patterns.md`.
- **편집 가능하게 만든다.** 표·차트·프로세스·간트·아키텍처는 `scripts/deck.py`의 네이티브 도형·차트로 만든다. 담당자가 PPT에서 바로 고칠 수 있어야 한다. 네이티브로 표현하기 어려운 복잡 도식만 이미지로 넣고 원본을 `_도식원본/`에 둔다.
- **만들고 나서 반드시 눈으로 본다.** 코드만 보고 완성이라고 판단하지 않는다.

## 제작 절차

1. `project.md`에서 템플릿 경로와 브랜드 설정을 읽는다.
   - **사업 폴더에 `디자인.md`(템플릿 실측 + 회사 가이드 병합본)가 있으면 먼저 읽고 그 값을 따른다.** 판형·좌표·색·글꼴 크기·금지 사항이 그 문서에 있다. 없으면 템플릿을 `--inspect`로 실측해 새로 만든다(글로비즈 예: `0. 탬플릿\디자인.md`).
2. **템플릿이 있으면 원형 장표 복제 방식으로 만든다.** 템플릿 안의 표지·목차·간지·본문·마지막 장을 "원형"으로 삼아 복제하고 글자만 바꾸므로, 글꼴·색·위치가 템플릿과 똑같이 나온다.
   - 템플릿을 이미지로 본다: `render_pptx.ps1 -Path <템플릿> -OutDir 05_PPT/_template`
   - 구조를 출력한다: `python <스킬>/scripts/deck.py --inspect <템플릿>` → 슬라이드 마스터·레이아웃(간지 번호별 레이아웃 포함), 장표별 도형 이름·위치·글꼴.
   - `PROTO`를 만든다. 원형 장표 번호와, 역할별 도형 이름을 적는다. 같은 이름 도형이 여러 개면 `이름#2`처럼 순서를 붙인다.
     ```python
     PROTO = {
       "cover":   {"slide": 1, "roles": {"title": "제목 1", "slogan": "Text 0#1", "date": "Text 0#2"}},
       "toc":     {"slide": 2},                       # 목차는 그대로 복제
       "section": {"slide": 3, "roles": {"title": "TextBox 4", "items": "TextBox 2"},
                   "layouts": {1: "사용자 지정 레이아웃", 2: "6_사용자 지정 레이아웃", 3: "8_사용자 지정 레이아웃"}},
       "content": {"slide": 4, "roles": {"title": "제목 2", "headline": "TextBox 9",
                                         "chapter": "TextBox 154", "section": "TextBox 155"},
                   "area": [0.42, 1.3, 9.99, 5.75]},   # 본문 패널 안쪽 좌표(인치)
       "end":     {"slide": 5},
     }
     d = Deck(TEMPLATE, BRAND, PROTO)
     ```
     (위 값은 글로비즈 템플릿 기준 실제 예시다.)
   - 본문 영역 `area`는 템플릿의 본문 패널(흰 상자 등) 안쪽 좌표를 재서 넣는다.
   - 장 번호 자동 매기기(Ⅰ. Ⅱ.)를 쓰는 글상자는 `section(번호, …)`를 부른 뒤의 `content_slide`가 번호를 맞춘다. `section="Ⅱ. 구축 방안 > 1. 시스템 구축 전략"`처럼 써도 된다.
   - 템플릿에 작업 중 도형이 남아 있으면 `PROTO[<원형>]["keep_only"] = True`로 역할 도형만 복제한다.
   - 템플릿 글꼴을 `BRAND["font"]`에, 굵은 글꼴이 따로 있으면 `font_bold`에 넣는다. 템플릿의 대표 색을 `primary`·`secondary`에 넣는다(추천값: `python <스킬>/scripts/reskin.py brand <템플릿>`).
   - 결정한 `TEMPLATE`·`BRAND`·`PROTO`를 `project.md`의 "브랜드 설정"에 기록한다(다른 담당자도 같은 값을 쓰게).
3. **템플릿이 없으면** 회사 가이드 서식(`references/company-guide.md`)을 반영한 기본 디자인으로 만들고 "임시 디자인"이라고 알린다.
4. **밀도를 정한다(1~5, 기본 3).** 사용자가 `밀도=4`처럼 주면 그대로, 없으면 `project.md`의 "기본 밀도", 그것도 없으면 **3**으로 만든다. 밀도는 **내용의 양**이다. 높을수록 내용을 더해 본문 영역의 빈 공간을 최소화하고, 낮을수록 중요하지 않은 내용을 빼서 만든다. 처음부터 글자를 작게 해서 채우지 않는다 — 기본 크기로 만들고, 넘칠 때만 6번 순서로 줄인다. 예전 표기 상/중/하는 4/3/2로 읽는다.
   `Deck(TEMPLATE, BRAND, PROTO, density=4)` · 빌드 스크립트는 `python build_deck.py <파일명> 4`로 받는다. 개수는 `d.pick(7,6,5,4,3)`·`d.cut(목록)`으로 맞춘다. `deck.py`를 쓰지 않고 직접 그리는 장표(시안, 기존 PPT 틀에 넣는 장표)도 같은 기준표를 따른다(`visual-patterns.md` 7장).
   렌더링 확인 때 **본문 영역 채움 비율**이 기준표의 목표에 못 미치면(빈 띠가 남으면) 근거 있는 항목·보조 블록을 더해 채우고(글자를 키우거나 도형만 늘려 채우지 않는다), 넘치면 중요도가 낮은 항목부터 빼거나 장을 나눈다. 보고할 때 적용한 밀도를 함께 적는다.
5. **장표마다 레시피를 고른다**(`visual-patterns.md` 6.1). 원고의 `도식` 필드 → 레시피(R1~R8) → 컴포넌트 조합. 한 컴포넌트로 본문을 다 채우지 않는다: 주 도식 + 보조 블록 + 하단 요약 바. 원고에 `슬로건`·`키워드`가 있으면 `content_slide(slogan=, keywords=)`에 넣는다. 아이콘은 뜻이 있는 자리에만 쓴다(장표당 개수 목표 없음, `visual-patterns.md` 8장).
6. **글자 크기는 회사 가이드를 따른다.** 본문 영역 글자는 10~8pt가 기준이고, 모든 장표를 이 크기로 시작한다. 원본(필수) 내용이 넘치면 ① 배치·간격·도형 크기를 조정하고 ② 그래도 넘치는 블록만 글자를 0.5pt씩 줄이되 **8pt 아래로 내리지 않으며** ③ 8pt에서도 넘치면 문구를 다듬거나 장을 나눈다. 밀도용으로 더한 내용 때문에 넘치면 글자를 줄이지 말고 그 내용을 줄인다. 보고할 때 기본 크기보다 줄인 장표를 적는다. `BRAND["text_scale"]`(기본 0.75)로 도식 안 글자를 줄인다. 강조 문구는 13pt까지. 기존 덱에 장을 추가할 때는 회사 가이드 기준보다 그 덱의 역할별 크기 체계를 우선한다. 흐름 라벨·글자 크기 혼용·줄 간격·덱 크기 체계 규칙은 `visual-patterns.md` 8장.
7. `05_PPT/build_deck.py`를 작성한다. 장별 함수로 나눈다(담당자가 자기 장만 다시 만들 수 있게).
   ```python
   import sys
   sys.path.insert(0, r"<스킬 폴더>/scripts")
   from deck import Deck

   TEMPLATE = None        # project.md의 템플릿 경로
   BRAND = {...}          # project.md의 브랜드 설정
   PROTO = None           # 템플릿이 있으면 원형 장표 설정

   def ch2(d):
       s = d.content_slide("무중단 인프라 전환", headline="… **2회 리허설**로 …합니다.",
                           section="Ⅱ. 구축 방안 > 1. 시스템 구축 전략", source="RFP p8",
                           slogan="미리 검증하는 리허설, 경매 없는 시간의 전환",
                           keywords="Rehearsal · Cut-over · Rollback")
       kpi, body = d.vsplit(d.area(), [0.16, 0.84], gap=0.15)          # 레시피 R2
       d.kpi_strip(s, kpi, [{"value": "0분", "label": "경매 중단", "icon": "순환"}, ...])
       left, right = d.hsplit(body, [0.68, 0.32], gap=0.2)
       d.timeline(s, d.tab_panel(s, left, "전환 일정"), periods=[...], rows=[...])
       d.label_desc_rows(s, d.tab_panel(s, right, "롤백 기준"), [("응답시간", "…"), ...])

   if __name__ == "__main__":
       d = Deck(TEMPLATE, BRAND, PROTO)
       d.cover(...); d.toc()
       d.section(2, "구축 방안", items=["1. 시스템 구축 전략", "2. 상세 구축 방안"])
       ch2(d)
       d.end()
       d.save("제안서_v1.pptx")
   ```
   - 원고의 헤드메시지·본문·근거를 그대로 옮긴다. PPT 단계에서 내용을 새로 지어내지 않는다.
   - 원고의 [근거 필요]·[자산 필요]는 장표에 그대로 남겨 눈에 띄게 한다(나중에 사람이 채운다).
   - 발표자 메모가 필요하면 `content_slide(notes=...)`.
8. 실행해서 `제안서_vN.pptx`를 만든다. 버전은 올려 가며 저장한다(덮어쓰지 않는다). 사용자가 직접 고친 PPT를 받아 작업하는 경우에는 빌드 스크립트로 다시 만들어 덮어쓰지 않고, 그 파일을 고쳐 `{원본 파일명}_001.pptx`처럼 번호를 붙여 저장한다(SKILL.md "수정 요청을 다룰 때").
9. **완료 조건 검사**: `python <스킬>/scripts/verify_pptx.py 05_PPT/제안서_v1.pptx` — 넘침·이탈·겹침·8pt 미만·튀는 색·빽빽한 표·두 줄 라벨이 모두 0이 될 때까지 고치고 다시 돌린다(SKILL.md "자기 검증 루프"). "크기혼용"은 의도한 강조인지 확인한다. 빽빽한 표는 표만 늘리지 말고 위 구역을 줄여 장표 전체 균형을 맞춘다(`visual-patterns.md` 7장 5번).
10. **렌더링 확인** (맥은 `python <스킬>/scripts/render_pptx_mac.py <pptx> 05_PPT/_render --slides 3,5`)
   ```
   powershell -ExecutionPolicy Bypass -File <스킬>/scripts/render_pptx.ps1 -Path 05_PPT/제안서_v1.pptx -OutDir 05_PPT/_render
   ```
   - 이미지를 직접 열어 본다. 확인할 것: 글자 넘침·잘림, 요소 겹침, **상자 안 빈 공간**(보조 블록을 붙이거나 비율 조정), **구역 간 밀도 차이**(위는 헐렁하고 아래만 빽빽하지 않은지), 글자 과다, 대비 부족, 연속 동일 레시피, 번호 마커가 내용을 가리는지.
   - `visual-patterns.md` 3장의 "AI스러운 장표 경고 신호"에 걸리면 도식을 바꾼다.
   - 고친 뒤 해당 장표만 다시 렌더링한다(`-Slides "3,5"`).
   - PowerPoint가 없는 환경이면 LibreOffice(`soffice --headless --convert-to pdf`) 후 PDF 페이지를 이미지로 본다. 단 맥의 Claude 실행 환경에서는 LibreOffice가 시스템 글꼴을 못 불러와 **한글이 빈칸으로 나온다.** headless 모드가 fontconfig를 쓰는데, 설정에 맥 글꼴 폴더가 없기 때문이다. 꼭 써야 하면 `/System/Library/Fonts`·`/Library/Fonts`·`~/Library/Fonts`를 `<dir>`로 적은 fonts.conf를 만들고, `FONTCONFIG_FILE=<절대 경로>`를 붙여 샌드박스 밖에서 실행한다. 상대 경로는 무시된다. 결과 PDF의 `pdffonts`에 템플릿 글꼴이 보이면 정상이다. 둘 다 안 되면 사용자에게 렌더링을 확인할 수 없다고 알리고 `verify_pptx.py`·python-pptx 구조 점검(도형 수·글꼴·좌표)으로 대신한다.
   - 사용자가 PowerPoint로 다른 파일을 편집 중이면 렌더링하지 않고 먼저 묻는다(렌더링이 PowerPoint에서 사본을 열고 닫아 작업을 방해한다). 허락이 없으면 파일만 저장해 넘기고 구조 점검으로 대신한다.

## 기존 PPT의 XML을 직접 고칠 때

장을 복제하거나 도형을 지우는 작업을 XML로 할 때 지킨다. 2026-11-05 MES 소개서 작업에서 실제로 PowerPoint 복구 창이 떴던 원인들이다.

- **장을 비울 때 `p:spTree`의 `p:nvGrpSpPr`·`p:grpSpPr`는 남긴다.** 이 둘을 지우면 `validate.py`는 통과해도 PowerPoint가 "내용에 문제가 있습니다… [복구]" 창을 띄운다. 도형을 지우는 반복문에서 이 두 요소는 건너뛴다.
- **복제한 장의 `p14:creationId`(장)와 `a16:creationId`(도형)는 새 값으로 바꾼다.** 원본 장과 같은 값이 남지 않게 한다.
- **`slideN.xml` 파일 번호는 장 순서와 다를 수 있다.** 사용자가 저장한 파일은 번호가 다시 매겨진다. `ppt/presentation.xml`의 `sldIdLst`와 `_rels`로 "몇 번째 장 = 어느 파일"을 먼저 확인하고 고친다.
- **XML 검사 통과가 곧 PowerPoint에서 열린다는 뜻은 아니다.** 고친 파일은 PowerPoint로 열어 복구 창이 없는지 확인한다(`render_pptx_mac.py`가 열 때 함께 확인된다). 확인용 사본은 고유 이름으로 열고 그 이름으로만 닫는다. "열린 문서 모두 닫기"는 사용자 문서까지 닫으므로 쓰지 않는다.

## deck.py 요약

| 메서드 | 용도 |
|---|---|
| `cover`, `toc`, `section(번호, 제목, items=[...])`, `end` | 표지, 목차, 간지, 마지막 장(템플릿 원형 모드면 원형 복제) |
| `clone(kind, texts)` | 템플릿 원형 장표 복제 후 역할별 글자 교체 |
| `python deck.py --inspect <템플릿>` | 템플릿 마스터·레이아웃·도형 이름 출력(PROTO 작성용) |
| `content_slide(title, headline, section, source, notes)` | 본문 장표 뼈대(제목·헤드메시지 띠·바닥글) |
| `area()`, `hsplit(box, ratios)`, `vsplit(box, ratios)`, `inset` | 영역 나누기 |
| `bullets`, `text`, `box`, `line`, `image` | 기본 요소 |
| `step_label(s, box, 번호, 이름, 색, size)` | 흐름 도형 라벨: 번호+이름 한 줄·같은 크기(안 들어가면 오류) |
| `Deck.row_heights(줄 수 목록, 크기, total=표 높이)` · `table(…, row_h=)` | 표 행 높이를 줄 수에 맞추고 남는 높이를 같은 여백으로 나눔 |
| `fix_line_spacing(k)` | 본문 줄 간격 고정 pt(`save()`가 `BRAND["line_pt"]`로 자동 호출) |
| `kpi_cards`, `process`, `timeline`, `table`, `compare`, `layers`, `cards`, `matrix`, `funnel`, `swimlane`, `org`, `chart`, `callout`, `steps_vertical` | 도식 |
| `lane_stack`, `feature_row`, `tile_grid`, `group_columns` | 3단 레인 통합 구조(R14) |
| `pillar_matrix`, `summary_bar`, `kpi_strip`, `tab_panel`, `header_card`, `numbered_cards`, `hub_quadrant`, `screen_callout`, `cycle_ring`, `pyramid`, `label_desc_rows`, `step_rail`, `section_bar`, `stamp`, `icon` | 회사 수주 제안서 스타일 컴포넌트(`visual-patterns.md` 6장) |
| `content_slide(…, slogan, keywords)` | 슬로건·영문 키워드 줄. 이후 `area()`는 그 아래부터 |

- 텍스트에 `**굵게**`를 쓰면 강조색 굵은 글씨가 된다(템플릿 원형의 헤드라인 자리에서도 동작).
- 템플릿이 없을 때의 기본 틀과 색·글꼴은 회사 공통 가이드(`company-guide.md`)를 따른다: 나눔스퀘어, `#002060`/`#123E8C`, 18pt 제목, Governing Message 상자, 표 좌우선 없음.
- `process`와 `org`는 사용한 아래쪽 y를 돌려준다. 남는 공간에 보조 도식을 둔다.
- 새 도식이 필요하면 `deck.py`에 메서드를 추가한다. 규칙: 글자는 도형 자체의 텍스트 프레임에 넣고(도형 위에 글상자를 겹치지 않는다), `_strip_style`로 테마 효과를 없애고 색을 명시한다.
- pptxgenjs를 써야 하는 상황이면 사용자 스킬 `pptx-coding-rule`의 규칙(도형+글상자 겹치기 금지, 단일 오브젝트 패턴)을 따른다.

## 결과 보고

- PPT 파일 링크, 장표 수, 렌더링으로 고친 점, 남은 [근거 필요]·[자산 필요] 위치.
- 템플릿이 임시 디자인이면 그 사실과, 템플릿을 주면 브랜드 설정만 바꿔 다시 만들 수 있다는 점.
