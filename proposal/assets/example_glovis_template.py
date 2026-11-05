"""회사 스타일 레시피(R1~R7) 예시 — 사업 템플릿 원형 복제 모드(글로비즈 템플릿). visual-patterns.md 6장 참조."""
import glob
import os
import sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "scripts"))
from deck import Deck

TEMPLATE = glob.glob(r"C:\workFISH\글로비즈\0. 탬플릿\*.pptx")[0]
BRAND = {"font": "페이퍼로지 4 Regular", "font_bold": "페이퍼로지 7 Bold",
         "primary": "030C39", "secondary": "004A8D", "accent": "0072CE",
         "light": "EEF1F5", "light2": "F7F8FA", "line": "C8CED6", "text": "333333", "subtext": "6B7280",
         "seq": ["0B2A5B", "004A8D", "0072CE", "2E9BD6"],
         "text_scale": 0.75}
PROTO = {
    "cover": {"slide": 1, "roles": {"title": "제목 1", "slogan": "Text 0#1", "date": "Text 0#2"}},
    "toc": {"slide": 2},
    "section": {"slide": 3, "roles": {"title": "TextBox 4", "items": "TextBox 2"},
                "layouts": {1: "사용자 지정 레이아웃", 2: "6_사용자 지정 레이아웃", 3: "8_사용자 지정 레이아웃",
                            4: "9_사용자 지정 레이아웃", 5: "10_사용자 지정 레이아웃"}},
    "content": {"slide": 4, "roles": {"title": "제목 2", "headline": "TextBox 9",
                                      "chapter": "TextBox 154", "section": "TextBox 155"},
                "area": [0.42, 1.3, 9.99, 5.75]},
    "end": {"slide": 5},
}
SEC1 = "구축 방안 > 1. 시스템 구축 전략"
SEC2 = "구축 방안 > 2. 상세 구축 방안"

d = Deck(TEMPLATE, BRAND, PROTO)
d.cover("오토비즈 핵심 서비스 고도화 제안서", "“ 차량 거래 경험을 연결에서 완결로 ”", "2026. 10. 00", "㈜엠투엠글로벌")
d.toc()
d.section(2, "구축 방안", items=["1. 시스템 구축 전략", "2. 상세 구축 방안", "3. 시스템 설계/구축 방안",
                                 "4. 디자인(UX/UI) 설계 전략", "5. 인터페이스 구축 방안"])

# 1. 추진 전략 — 기둥 매트릭스 + 하단 요약 바
s = d.content_slide("1. 추진 전략",
                    headline="4대 추진 전략으로 **경매 중단 없는 전환**과 **회원·데이터 통합**을 8개월 안에 완수합니다.",
                    section=SEC1, slogan="끊김 없는 전환, 하나로 잇는 거래 경험",
                    keywords="Zero-Downtime · Unified Member · Data Integrity · Operational Stability")
x, y, w, h = d.area()
body, bar = d.vsplit((x, y, w, h), [0.8, 0.2], gap=0.15)
d.pillar_matrix(s, body, [
    {"title": "무중단 전환", "en": "Zero-Downtime", "rows": [
        ["Java 17·Spring 6 전환", "OpenRewrite 자동 변환", "경매 없는 시간대 컷오버"], ["전환 계획서", "리허설 결과서", "롤백 절차서"], ["경매 중단 0분", "롤백 30분 이내"]]},
    {"title": "회원 통합", "en": "Unified Member", "rows": [
        ["3개 유형 → 2개 유형", "오토벨 단일 가입", "1인 1단말 인증"], ["회원 모델 설계서", "전환 매핑표", "안내·재동의 계획"], ["중복 계정 0건", "가입 단계 축소"]]},
    {"title": "데이터 정합", "en": "Data Integrity", "rows": [
        ["회원·고객 마스터 분리", "이관 검증 자동화", "TBC20 정제"], ["데이터 이관 계획서", "검증 리포트", "데이터 사전"], ["건수·금액 대사 100%", "이관 오류 0건"]]},
    {"title": "안정 운영", "en": "Stable Operation", "rows": [
        ["성능 기준선 비교", "장애 등급별 대응", "운영 인수인계"], ["운영 전환 계획서", "모니터링 대시보드", "운영 매뉴얼"], ["응답시간 전환 전 수준", "1등급 장애 0건"]]},
], ["핵심 업무", "산출물", "검증 포인트"])
d.summary_bar(s, bar, [
    {"title": "0분", "desc": "실시간 경매 중단 시간", "icon": "i231"},
    {"title": "2개 유형", "desc": "회원 유형 통합(3→2)", "icon": "i236"},
    {"title": "100%", "desc": "이관 데이터 대사율", "icon": "i239"},
    {"title": "2회", "desc": "운영 전환 리허설", "icon": "i235"},
])
d.stamp(s, x + w - 0.7, bar[1] - 0.12, "목표치")

# 2. 무중단 전환 — 타임라인 + KPI 띠 + 탭 패널
s = d.content_slide("2. 무중단 인프라 전환",
                    headline="경매가 없는 시간대 일괄 전환과 **2회 리허설**로 실시간 경매 중단 없이 Java 17 전환을 끝냅니다.",
                    section=SEC1, source="RFP p8, 요구사항 1-1 #1",
                    slogan="리허설로 검증하고, 경매 없는 시간에 전환한다", keywords="Rehearsal · Cut-over · Rollback · Monitoring")
x, y, w, h = d.area()
kpi, body = d.vsplit((x, y, w, h), [0.16, 0.84], gap=0.15)
d.kpi_strip(s, kpi, [{"value": "0분", "label": "경매 중단 시간", "icon": "i231"},
                     {"value": "2회", "label": "전환 리허설", "icon": "i235"},
                     {"value": "30분", "label": "롤백 완료 목표", "icon": "i220"},
                     {"value": "D+7", "label": "안정화 모니터링", "icon": "i221"}])
left, right = d.hsplit(body, [0.68, 0.32], gap=0.2)
inner = d.tab_panel(s, left, "전환 일정(D-Day 기준)")
d.timeline(s, inner, periods=["D-14", "D-10", "D-7", "D-3", "D-Day", "D+3", "D+7"], label_w=1.5,
           bands=[{"label": "준비", "start": 0, "end": 3}, {"label": "전환·안정화", "start": 3, "end": 7, "color": "004A8D"}],
           rows=[{"label": "인벤토리·자동 변환", "start": 0, "end": 2, "text": "OpenRewrite"},
                 {"label": "모의경매 리허설", "start": 1.5, "end": 3.5, "text": "2회", "milestones": [(2, "1차"), (3.2, "2차")]},
                 {"label": "서버 일괄 전환", "start": 4, "end": 4.6, "color": "0072CE", "milestones": [(4.3, "컷오버")]},
                 {"label": "CS 강제 업데이트", "start": 3.5, "end": 5, "text": "버전 체크"},
                 {"label": "안정화 모니터링", "start": 4.6, "end": 7, "text": "응답시간·TPS"}])
inner = d.tab_panel(s, right, "롤백 기준", color="0072CE")
d.label_desc_rows(s, inner, [("응답시간", "전환 전 대비 20% 이상 지연 시"), ("오류율", "5xx 오류 1% 초과 시"),
                             ("경매 체결", "체결 실패 1건 발생 시"), ("판단", "PM·고객 담당자 합동 결정")], label_w=0.8)

# 3. 회원 체계 통합 — 비교 + 스텝 레일
s = d.content_slide("3. 회원 체계 통합",
                    headline="3개로 나뉜 회원 유형을 2개로 통합해 **947개 미가입 사업자**의 경매 전환 장벽을 없앱니다.",
                    section=SEC2, source="RFP p6, 요구사항 3-1 #1~6",
                    slogan="한 번 가입으로 조회부터 경매까지", keywords="Single Sign-up · Customer Master · 1 User 1 Device")
x, y, w, h = d.area()
top, bot = d.vsplit((x, y, w, h), [0.68, 0.32], gap=0.2)
d.compare(s, top,
          {"title": "AS-IS", "sub": "유형별 계정 최대 3개 · 도메인 분리", "items": [
              "오토벨 일반 / 딜러 / 스마트옥션 경매 별도 가입", "경매 가입은 스마트옥션 도메인에서만 가능",
              "TBC20 한 테이블에 회원·비회원 혼재", "상사 공동계정으로 실사용자 식별 불가"]},
          {"title": "TO-BE", "sub": "1인 2유형 · 오토벨 단일 가입", "items": [
              "**오토벨 일반** / **딜러&경매** 2개 유형", "오토벨 단일 URL에서 모든 유형 가입",
              "회원 테이블과 고객 마스터 테이블 분리", "1인 1단말 인증으로 실사용자 기반 운영"]}, arrow_text="통합")
d.step_rail(s, bot, [{"title": "현황 분석", "desc": ["유형별 계정 집계", "중복·휴면 계정 식별"]},
                     {"title": "모델 설계", "desc": ["회원·고객 마스터 분리", "유형별 권한 정의"]},
                     {"title": "매핑·이관", "desc": ["계정 병합 규칙 적용", "이관 리허설 2회"]},
                     {"title": "전환 안내", "desc": ["기존 회원 재동의", "1인 1단말 인증 등록"]},
                     {"title": "검증", "desc": ["건수·권한 대사", "로그인 시나리오 점검"]}])

# 4. 플랫폼 구조 — 허브 + 4사분면
s = d.content_slide("4. 통합 플랫폼 구성",
                    headline="회원·경매·차량·정산을 하나의 **오토비즈 공통 기반** 위에 올려 서비스 간 단절을 없앱니다.",
                    section=SEC2, slogan="따로 있던 서비스를 하나의 기반으로",
                    keywords="Common Platform · Member · Auction · Settlement")
hub, bar = d.vsplit(d.area(), [0.8, 0.2], gap=0.15)
d.hub_quadrant(s, hub, ("오토비즈", "공통 기반"), [
    {"title": "회원·인증", "sub": "Member & Auth", "bullets": ["2개 유형 통합 가입", "1인 1단말 인증", "공통 권한 관리"], "icon": "i245"},
    {"title": "경매", "sub": "Auction", "bullets": ["실시간 경매 무중단", "CS 클라이언트 연계", "체결 이력 통합"], "icon": "i225"},
    {"title": "차량·매물", "sub": "Vehicle", "bullets": ["오토벨 매물 연계", "차량 이력 조회", "시세 데이터 공유"], "icon": "i227"},
    {"title": "정산·데이터", "sub": "Settlement & Data", "bullets": ["고객 마스터 통합", "정산 데이터 대사", "통계 리포트"], "icon": "i241"},
])
d.summary_bar(s, bar, [{"title": "단일 가입", "desc": "모든 서비스 한 계정", "icon": "i236"},
                       {"title": "공통 데이터", "desc": "고객 마스터 하나로", "icon": "i239"},
                       {"title": "무중단 운영", "desc": "경매 중단 0분", "icon": "i231"}])

# 5. 관리자 화면 — 화면 콜아웃
s = d.content_slide("5. 통합 관리자 화면",
                    headline="회원 유형·경매 현황·전환 지표를 한 화면에서 보도록 **통합 관리자 화면**을 제공합니다.",
                    section=SEC2, source="요구사항 3-2 #4",
                    slogan="보고, 판단하고, 바로 처리한다", keywords="Dashboard · Member Admin · Alert")
d.screen_callout(s, d.area(), [
    {"title": "핵심 지표 카드", "desc": ["신규 가입·진행 경매·응답시간", "전일 대비 증감 표시"]},
    {"title": "가입 추이 차트", "desc": ["일반·딜러&경매 유형별 추이", "전환율 목표선 표시"]},
    {"title": "회원 유형 비중", "desc": ["유형별 비중 도넛 차트", "미가입 사업자 전환 현황"]},
    {"title": "처리 대기 목록", "desc": ["승인 대기 회원·문의", "목록에서 바로 처리"]},
], markers=[0.215, 0.45, 0.85, 0.5], title="오토비즈 통합 관리자")

# 6. 장애 대응 — 피라미드 + 번호 카드
s = d.content_slide("6. 장애 대응 체계",
                    headline="장애를 **3개 등급**으로 나눠 등급별 담당·시간 기준으로 대응해 경매 영향을 최소화합니다.",
                    section=SEC2, slogan="등급으로 나누고, 시간으로 약속한다", keywords="Severity · Escalation · SLA")
x, y, w, h = d.area()
left, right = d.hsplit((x, y, w, h), [0.5, 0.5], gap=0.25)
inner = d.tab_panel(s, left, "장애 등급")
d.pyramid(s, inner, [
    {"name": "1등급", "owner": "PM·고객 담당자", "desc": "경매 중단·체결 불가 / 30분 내 복구"},
    {"name": "2등급", "owner": "모듈 리더", "desc": "일부 기능 장애 / 4시간 내 조치"},
    {"name": "3등급", "owner": "담당 개발자", "desc": "화면·문구 오류 / 익일 반영"},
], width_ratio=0.42)
rt, rb = d.vsplit(right, [0.55, 0.45], gap=0.15)
inner = d.tab_panel(s, rt, "대응 절차", color="0072CE")
d.numbered_cards(s, inner, [
    {"title": "감지", "bullets": ["모니터링 경보", "5분 내 등급 판정"], "icon": "i221"},
    {"title": "조치", "bullets": ["담당 즉시 투입", "필요 시 롤백"], "icon": "i220"},
    {"title": "재발 방지", "bullets": ["원인 분석 보고", "점검 항목 추가"], "icon": "i235"},
], gap=0.1, style="border", size=10)
d.table(s, rb, ["등급", "보고", "복구 목표", "사후 조치"],
        [["1등급", "즉시(유선)", "30분", "원인 보고서 24시간 내"], ["2등급", "30분 내", "4시간", "주간 보고 반영"],
         ["3등급", "일일 보고", "익일", "배포 이력 관리"]], col_widths=[0.8, 1, 1, 1.8], size=11)

# 7. 기대 효과 — 순환 링
s = d.content_slide("7. 기대 효과",
                    headline="전환·통합·운영이 맞물려 **오토비즈 거래 경험 전체**가 한 단계 올라갑니다.",
                    section=SEC1, slogan="연결에서 완결로", keywords="Continuity · Simplicity · Reliability · Growth")
ring, kpi = d.vsplit(d.area(), [0.82, 0.18], gap=0.15)
d.kpi_strip(s, kpi, [{"value": "0분", "label": "경매 중단", "icon": "i231"}, {"value": "3→2", "label": "회원 유형", "icon": "i236"},
                     {"value": "947개", "label": "경매 전환 대상 사업자", "icon": "i226"}, {"value": "100%", "label": "데이터 대사", "icon": "i239"}])
d.cycle_ring(s, ring, ("오토비즈", "고도화"), [
    {"title": "끊김 없는 경매", "desc": "전환 중에도 경매 중단 0분", "icon": "i231"},
    {"title": "쉬운 가입", "desc": "단일 가입으로 경매까지", "icon": "i236"},
    {"title": "믿을 수 있는 데이터", "desc": "회원·고객 마스터 분리", "icon": "i239"},
    {"title": "지속 성장", "desc": "미가입 사업자 경매 유입", "icon": "i222"},
])
d.end()
out = "제안서_예시_글로비즈템플릿.pptx"
d.save(out)
print("saved", out)
