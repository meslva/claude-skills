"""deck.py 컴포넌트 예시. 실행하면 현재 폴더에 example_deck.pptx를 만든다.

    python example_build_deck.py
    powershell -ExecutionPolicy Bypass -File ../scripts/render_pptx.ps1 -Path example_deck.pptx -OutDir _render
"""
import os
import sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "scripts"))
from deck import Deck

d = Deck(brand={"company": "제안사", "doc_title": "오토비즈 핵심 서비스 고도화"})

d.cover("현대글로비스 오토비즈\n핵심 서비스 고도화 프로젝트", "기술 제안서", "2026. 10", "제안사")

d.toc([("Ⅰ", "제안 개요", "배경·범위·추진전략·기대효과"),
       ("Ⅱ", "구축 방안", "인프라 전환·플랫폼 경매·정보보호·회원 통합"),
       ("Ⅲ", "사업 관리", "방법론·일정·조직·품질"),
       ("Ⅳ", "지원 부문", "기술이전·교육·비상대책"),
       ("Ⅴ", "제안사 일반", "회사 소개·유사 실적")])

# 1. 컷오버 타임라인 + KPI
s = d.content_slide("무중단 인프라 전환 전략",
                    headline="경매가 없는 시간대 일괄 전환과 2회 리허설로 실시간 경매 중단 없이 Java 17 전환을 완료합니다.",
                    section="Ⅱ. 구축 방안 > 1. 인프라 최신화", source="RFP p8, 요구사항 1-1 #1 (Hazelcast 3.12.2 → 5.6.x, 사용 파일 35개)")
left, right = d.hsplit(d.area(), [0.64, 0.36])
d.timeline(s, left, periods=["D-14", "D-10", "D-7", "D-3", "D-Day", "D+3", "D+7"],
           bands=[{"label": "준비", "start": 0, "end": 3}, {"label": "전환·안정화", "start": 3, "end": 7, "color": "2E6BD6"}],
           rows=[{"label": "인벤토리·자동 변환", "start": 0, "end": 2, "text": "OpenRewrite 변환"},
                 {"label": "모의경매 리허설", "start": 1.5, "end": 3.5, "text": "리허설 2회", "milestones": [(2, "1차"), (3.2, "2차")]},
                 {"label": "서버 일괄 전환", "start": 4, "end": 4.6, "text": "", "color": "F28C28", "milestones": [(4.3, "컷오버")]},
                 {"label": "CS 클라이언트 강제 업데이트", "start": 3.5, "end": 5, "text": "배포·버전 체크"},
                 {"label": "안정화 모니터링", "start": 4.6, "end": 7, "text": "응답시간·TPS 비교", "color": "1F9E89"}])
d.kpi_cards(s, right, [{"value": "0분", "label": "경매 중단 시간", "note": "경매 없는 시간대 컷오버"},
                       {"value": "2회", "label": "전환 리허설", "note": "롤백 절차 포함"}], cols=1)

# 2. As-Is / To-Be
s = d.content_slide("회원 체계 통합",
                    headline="3개로 나뉜 회원 유형을 2개로 통합해 947개 미가입 사업자의 경매 전환 장벽을 없앱니다.",
                    section="Ⅱ. 구축 방안 > 4. 회원·정산 체계", source="RFP p6, 요구사항 3-1 #1~6")
d.compare(s, d.area(),
          {"title": "AS-IS", "sub": "유형별 계정 최대 3개 · 도메인 분리", "items": [
              "오토벨 일반 / 오토벨 딜러 / 스마트옥션 경매 별도 가입",
              "경매 가입은 스마트옥션 도메인에서만 가능",
              "TBC20 한 테이블에 회원·비회원 혼재, 동일 고객 중복 적재",
              "상사 단위 공동계정으로 실사용자 식별 불가"]},
          {"title": "TO-BE", "sub": "1인 2유형 · 오토벨 단일 가입", "items": [
              "**오토벨 일반** / **딜러&경매** 2개 유형, 권한으로 메인·서브·사무장 구분",
              "오토벨 단일 URL에서 모든 유형 가입",
              "회원 테이블과 고객 마스터 테이블 분리",
              "1인 1단말 인증으로 실사용자 기반 운영"]},
          arrow_text="통합")

# 3. 레이어 아키텍처
s = d.content_slide("목표 시스템 아키텍처",
                    headline="채널·서비스·데이터 계층을 표준화하고 플랫폼 경매 엔진을 독립 모듈로 분리합니다.",
                    section="Ⅱ. 구축 방안 > 2. 시스템 아키텍처")
d.layers(s, d.area(), [
    {"name": "채널", "items": ["오토벨 Web/App", "스마트옥션 Web/App", "CS 응찰(OpenJFX)", "GAOS(Nexacro)"]},
    {"name": "서비스", "items": [{"text": "플랫폼 경매 엔진", "new": True}, "회원·권한", "정산·PG", "알림"]},
    {"name": "기반", "items": ["Nginx / Tomcat 10", "Java 17 · Spring 6", "Netty 4.2 · Hazelcast 5"]},
    {"name": "데이터", "items": ["Tibero", "Redis", "통합로그 연계"]}],
    side=[{"name": "정보보호", "items": ["마스킹", "접속기록 2년", "파기 배치"]},
          {"name": "운영", "items": ["Git-Runner", "모니터링"]}])

# 4. 차트 + 퍼널
s = d.content_slide("경매 회원 전환 목표",
                    headline="플랫폼 가입 매매사업자의 경매 회원 비율 25%를 단계적으로 끌어올립니다.",
                    section="Ⅰ. 제안 개요 > 5. 기대효과", source="RFP p6 (’26.05 기준 1,267개 중 320개) · 목표치는 [가정]")
l, r = d.hsplit(d.area(), [0.5, 0.5])
d.funnel(s, l, [{"label": "플랫폼 가입 매매사업자", "value": "1,267개", "note": "오토벨 가입 딜러"},
                {"label": "경매 회원", "value": "320개", "note": "현재 25%"},
                {"label": "목표 경매 회원 [가정]", "value": "600개+", "note": "통합 가입·권한 추가 유도"}])
d.chart(s, r, "column", ["현재", "1차 오픈", "2차 오픈", "1년 후"], {"경매 회원 비율(%)": [25, 30, 40, 50]},
        title="경매 회원 비율 목표 [가정]", number_format='0"%"')

# 5. 스윔레인
s = d.content_slide("경매(서브) 회원 가입 절차",
                    headline="서브 회원 가입은 메인 회원의 전자서명 동의로 완료되어 공동계정 사용을 줄입니다.",
                    section="Ⅱ. 구축 방안 > 4. 회원·정산 체계", source="요구사항 3-1 #3, #13, #18")
d.swimlane(s, d.area(), ["신청자", "메인 회원", "고객센터"],
           [{"lane": 0, "col": 0, "text": "오토벨 가입\n(딜러&경매)"},
            {"lane": 0, "col": 1, "text": "서브 권한 선택\n메인 ID 입력"},
            {"lane": 1, "col": 2, "text": "동의 알림톡 수신\n(24시간 유효)", "kind": "hl"},
            {"lane": 1, "col": 3, "text": "본인인증\n전자서명"},
            {"lane": 2, "col": 4, "text": "서류 검토\n승인", "kind": "decision"},
            {"lane": 0, "col": 5, "text": "가입 완료", "kind": "end"}], cols=6)

# 6. 표 (권고안 강조)
s = d.content_slide("Tiles 처리 방안 비교",
                    headline="화면 149개의 재작성 위험을 피하기 위해 JSP 유지·Tiles 제거 방식을 권고합니다.",
                    section="Ⅱ. 구축 방안 > 1. 인프라 최신화", source="요구사항 1-1 #2")
d.table(s, d.vsplit(d.area(), [0.72, 0.28])[0], ["구분", "A. JSP 유지 + Tiles 제거 (권고)", "B. Thymeleaf 전환", "C. Tiles 유지·패치"],
        [["작업 범위", "레이아웃 include 전환", "화면 템플릿 전면 재작성", "비공식 패치 적용"],
         ["화면 영향", "149개 화면 로직 유지", "149개 화면 재작성", "영향 적음"],
         ["일정 위험", "낮음", "높음", "중간(장기 유지보수 불가)"],
         ["지원 지속성", "표준 JSP", "표준", "지원 종료"]],
        col_widths=[1, 2, 2, 2], highlight_col=1)
d.callout(s, d.vsplit(d.area(), [0.72, 0.28])[1], "권고: **A안** — 1단계 3.5개월 일정 안에서 화면 로직을 건드리지 않고 Spring 6 호환을 확보합니다.")

# 7. 리스크 매트릭스 + 카드
s = d.content_slide("주요 리스크와 대응",
                    headline="일정·범위 리스크 14건을 발생 가능성과 영향으로 관리하고 상위 3건은 착수 전 대응합니다.",
                    section="Ⅲ. 사업 관리 > 1. 위험 관리")
l, r = d.hsplit(d.area(), [0.45, 0.55])
d.matrix(s, l, [{"label": "1단계 일정 초과", "x": 2, "y": 2}, {"label": "able 의존 범위", "x": 1, "y": 2},
                {"label": "TBC20 영향 과다", "x": 2, "y": 2}, {"label": "알림톡 심사 지연", "x": 1, "y": 1},
                {"label": "핵심 인력 교체", "x": 0, "y": 2}])
d.cards(s, r, [{"title": "1단계 분할 오픈", "body": ["1-A: EOS·ISMS-P 직결 항목", "1-B: 회원 DB 개편 묶음 오픈"], "tag": "R1 일정"},
               {"title": "착수 3주 인벤토리", "body": ["able·JEUS 의존 지점 전수 식별", "재기준선 후 일정 확정"], "tag": "R2 범위"},
               {"title": "호환 뷰 단계 전환", "body": ["TBC20 호환 View 제공", "서비스별 순차 전환"], "tag": "R4 데이터"}], cols=1)

# 8. 프로세스 + 조직도
s = d.content_slide("단계별 추진 절차",
                    headline="요건정의부터 안정화까지 2단계로 나눠 오픈하고 단계마다 품질 관문을 둡니다.",
                    section="Ⅲ. 사업 관리 > 2. 추진 일정")
bot = d.process(s, d.area(), [
    {"title": "요건정의", "period": "M ~ M+1", "desc": ["현업 인터뷰", "미정의 규칙 13건 확정"]},
    {"title": "분석·설계", "period": "M+1 ~ M+2", "desc": ["아키텍처 설계", "데이터 이행 설계"]},
    {"title": "개발", "period": "M+1 ~ M+5", "desc": ["자동 변환·수동 보정", "플랫폼 경매 엔진"]},
    {"title": "테스트", "period": "M+3 ~ M+6", "desc": ["회귀 자동화", "모의해킹·부하 테스트"]},
    {"title": "이행·안정화", "period": "M+3.5 / M+7", "desc": ["리허설 2회", "ISMS-P 보완 대응"]}])
d.kpi_cards(s, (0.5, bot + 0.3, 12.33, 6.8 - bot - 0.3), [
    {"value": "M+3.5", "label": "1차 오픈", "note": "인프라·ISMS-P·DB 구조"},
    {"value": "M+7", "label": "2차 오픈", "note": "플랫폼 경매·회원·정산"},
    {"value": "4회", "label": "품질 관문", "note": "설계·개발·테스트·이행"}], value_size=26)

s = d.content_slide("프로젝트 수행 조직",
                    headline="AA·DBA를 PM 직속으로 두어 인프라 전환과 DB 개편의 기술 판단을 빠르게 내립니다.",
                    section="Ⅲ. 사업 관리 > 3. 수행 조직")
bot = d.org(s, d.area(), {"title": "PM", "name": "[자산 필요: PM 성명]"},
      [{"title": "인프라 전환", "members": ["AA", "CS·WEB 전환", "iOS"]},
       {"title": "플랫폼 경매", "members": ["엔진 BE", "GAOS(Nexacro)", "스마트옥션 FO"]},
       {"title": "회원·정산", "members": ["회원·권한", "오토벨 FO·Admin", "PG 연동"]},
       {"title": "정보보호", "members": ["마스킹·로그", "파기 배치"]},
       {"title": "품질·보안", "members": ["QA", "보안 점검 대응"]}],
      side=[{"title": "발주처", "note": "현대글로비스·오토에버"}, {"title": "PMO/QA", "note": "품질 관문"}])
d.table(s, (0.5, bot + 0.3, 12.33, 6.8 - bot - 0.3), ["역할", "책임", "투입 시점", "비고"],
        [["AA", "Java·Spring 전환 기술 리드", "착수~1차 오픈", "AI 도구 활용 경험"],
         ["DBA", "TBC20 분리·이행·암호화 설계", "착수~2차 오픈", "필수 기재"]], col_widths=[1, 3, 2, 2])

out = os.path.join(os.getcwd(), "example_deck.pptx")
d.save(out)
print("saved", out, d.page, "slides")
