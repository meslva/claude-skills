# claude-skills

> Claude Code에서 쓰는 개인 스킬 모음입니다.
> 제안서 작성과 이전 제안서 아카이브 관리를 위한 스킬이 들어 있습니다.

| 스킬 | 부르는 법 | 하는 일 |
|---|---|---|
| [`proposal`](proposal/) | `/proposal` | RFP로 기술 제안서를 만드는 전 과정: 착수 → 분석 → 질의 → 전략 → 원고 → PPT → 검수 |
| [`제안서`](제안서/) | `/제안서` | `proposal`을 부르는 한글 명령 |
| [`pdf2md`](pdf2md/) | 자동 (PDF를 주며 "넣어줘", "md로 변환해줘" 등) | 제안서 PDF를 `~/Desktop/이전제안서`에 등록하고 페이지별 Markdown으로 변환 |

---

## 설치

**플러그인으로 설치 (권장)**

Claude Code에서 아래 두 명령을 실행합니다. 세 스킬이 한꺼번에 설치됩니다.

```
/plugin marketplace add meslva/claude-skills
/plugin install proposal-skills@meslva-skills
```

스킬을 고쳐 push한 뒤 다른 컴퓨터에서 받을 때는 `/plugin marketplace update meslva-skills`를 실행합니다.

**git clone으로 설치**

스킬은 `~/.claude/skills/` 아래에 있어야 Claude Code가 인식합니다.

**스킬 폴더가 비어 있는 새 컴퓨터**

```bash
git clone https://github.com/meslva/claude-skills.git ~/.claude/skills
```

**이미 다른 스킬이 있는 컴퓨터**

```bash
git clone https://github.com/meslva/claude-skills.git ~/claude-skills
ln -s ~/claude-skills/proposal ~/.claude/skills/proposal
ln -s ~/claude-skills/제안서  ~/.claude/skills/제안서
ln -s ~/claude-skills/pdf2md  ~/.claude/skills/pdf2md
```

설치한 뒤에는 **새 세션**을 열어야 명령이 목록에 나타납니다.

### 필요한 도구

| 스킬 | 필요한 것 |
|---|---|
| `proposal` | Python 패키지 `python-pptx` `openpyxl` `pdfplumber` `pypdfium2`, 렌더링 확인용 Microsoft PowerPoint |
| `pdf2md` | poppler의 `pdftotext` `pdfinfo` (`brew install poppler`) |

---

## proposal · 제안서

한국 SI·IT 구축 사업의 RFP와 요구사항 자료로 기술 제안서를 만듭니다. 한 번에 끝내지 않고, 네 번의 관문(분석·전략·원고·최종)에서 멈춰 확인을 받으며 진행합니다. 진행 상태는 작업 폴더의 `제안작업/project.md`에 저장되므로, 다음 날 `/제안서`만 입력해도 이어서 진행합니다.

| 명령 | 하는 일 |
|---|---|
| `/제안서` | 처음이면 착수, 아니면 진행 상태를 보고 다음 할 일을 제안 |
| `/제안서 착수` | 사업 설정 파일과 제출일 기준 역산 일정표 |
| `/제안서 분석` | 분석보고서, 요구사항 매트릭스(엑셀), 용어사전 → **관문 A** |
| `/제안서 질의` | 질의서 엑셀 |
| `/제안서 전략` | Win theme, 차별화, 목차·배점 배분 → **관문 B** |
| `/제안서 원고 [N장]` | 장표별 원고 → **관문 C** |
| `/제안서 ppt [N장] [밀도=1~5]` | 확정 원고로 PPT를 만들고 PowerPoint로 렌더링해 확인 |
| `/제안서 테마 <템플릿> [PPT]` | 템플릿의 색·글꼴을 입힘 (구조·배치는 그대로) |
| `/제안서 검수` | 누락·모호 표현·수치 불일치 점검 → **관문 D** |
| `/제안서 일정` · `상태` | 일정 재계산, 진행 상황 보고 |

자세한 사용법은 [`proposal/사용법.md`](proposal/사용법.md)를 보세요.

<details>
<summary><b>폴더 구성</b></summary>

```
proposal/
├── SKILL.md              스킬 본체 (단계별 진행 규칙)
├── 사용법.md              사용자용 안내
├── references/           단계별 지침
│   ├── 00-kickoff.md … 07-theme.md
│   ├── visual-patterns.md    도식 패턴 + 장표 품질 규칙(사용자 지적 모음)
│   ├── style-guide.md
│   ├── company-guide.md
│   └── company-assets.md
├── scripts/
│   ├── deck.py               PPT 도식 컴포넌트 (python-pptx)
│   ├── verify_pptx.py        넘침·겹침·글자 크기 등 완료 조건 검사
│   ├── render_pptx_mac.py    맥 PowerPoint 렌더링
│   ├── render_pptx.ps1       윈도우 PowerPoint 렌더링
│   ├── reskin.py             템플릿 색·글꼴 적용
│   ├── extract_inputs.py     PDF·엑셀·Word 입력 추출
│   └── review_storyboard.py
├── assets/               빌드 예시, 템플릿, 아이콘
└── evals/                스킬 평가 케이스
```

</details>

---

## pdf2md

이전 제안서 PDF를 아카이브에 쌓아 두고, 검색할 수 있게 텍스트로 바꿉니다.

```
~/Desktop/이전제안서/
├── pdf/   원본 PDF
└── md/    같은 이름의 .md (페이지마다 <!-- p.N --> 표시)
```

PDF 파일을 주면서 "넣어줘", "등록해줘", "md로 변환해줘"라고 하면 자동으로 동작합니다. 직접 실행할 수도 있습니다.

```bash
python3 ~/.claude/skills/pdf2md/scripts/import_pdf.py "<PDF 경로>" ["<PDF 경로>" ...]
```

| 옵션 | 뜻 |
|---|---|
| `--root DIR` | 아카이브 위치가 다를 때 |
| `--force` | 같은 이름이 있으면 덮어쓰기 (사람이 손본 md가 있을 수 있으니 확인 후에만) |

---

## 고치고 올리기

스킬을 고친 뒤에는 커밋해서 이력을 남깁니다.

```bash
cd ~/.claude/skills
git add -A
git commit -m "proposal: 장표 품질 규칙 추가"
git push
```

다른 컴퓨터에서 최신 내용을 받을 때는 `git pull`을 실행합니다.

## 이 저장소에 없는 것

- `resume` 스킬 — 개인 이력 작업용이라 올리지 않습니다 (`.gitignore`).
- `pptx`, `docx` 같은 Claude 앱 기본 스킬 — 앱이 따로 관리합니다.
