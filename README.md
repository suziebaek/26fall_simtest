# 워드 문제 → 엑셀 문제은행 변환기 (AI 없이, 규칙 기반)

**문제 문서 하나만 업로드**하면, 내장된 공식 스키마(원고템플릿)와 챕터/CELL
참고자료를 자동 적용하고, 정답/유형 등은 **규칙 기반으로 초안을 채운 뒤**
표에서 직접 검수/수정해서 엑셀로 내보내는 Streamlit 앱입니다.

**AI(LLM) 호출이 없습니다.** 이전 버전에 있던 Claude/Gemini 자동 채우기
단계는 제거되었고, 대신:
1. 문서 파싱(지문/보기/빨간색 정답 표시 추출)
2. (선택) 정답지 엑셀 병합 — 챕터/문항번호로 매칭해 정답·난이도를 가져옴
3. 규칙 기반 필드 채우기 (`rule_fill.py`) — q_type/tp_type 추정, Answer 포맷팅 등
4. 사람이 표에서 직접 검수/수정

의 순서로 동작합니다.

## 설치 및 실행

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 스키마 (내장 · 두 레벨 공용)

`resources/H_sample.xlsx` (공식 원고템플릿)을 기준으로 아래 26개 컬럼을 사용합니다.
`explanation`(해설) 컬럼은 없습니다.

```
service_code, track_code, top_cors_id, component_code, book_code, act_code,
level_code, lesson_order_seq, lesson_title, page_order_seq, cell_id, cell_title,
q_type, tp_type, difficulty, text_passage, image_passage, question, text_prompt,
image_prompt, text_example_1~5, Answer
```

## 레벨별 차이

| 항목 | H레벨 | E레벨 |
|---|---|---|
| 스키마(컬럼 구성) | 26개 컬럼 | H와 완전히 동일 |
| 공통 메타데이터 | `level_code=TO_G_H_AU`, `cell_id` 접두사 `GR-H-CH..` | `level_code=TO_G_E_AU`, `cell_id` 접두사 `GR-E-CH..` |
| 챕터/CELL 참고문서 | `resources/H_chapter_ref.docx` (문단 형식, 9개 챕터) | `resources/E_chapter_ref.docx` (표 형식, 9개 챕터) |
| 정답 표시 | 문서에 빨간 글씨로 정답이 표시되어 있는 경우가 많음 | 정답 표시가 없는 문서일 수 있음 — 정답지 업로드 또는 검수 단계에서 직접 입력 |

`resources/E_sample.xlsx`는 실제 E 커리큘럼에 맞춰 `lesson_title="Chapter 09/관계사"`,
`cell_id` 접두사 `GR-E-CH09-*`로 맞춰져 있습니다 (E레벨 9번째 챕터가 관계사).

## 규칙 기반 채우기가 하는 일 / 못 하는 일 (`rule_fill.py`)

**자동으로 채우는 것:**
- `q_type`/`tp_type`: 보기(①~⑤) 유무와 지시문 키워드(배열/조건/고치시오 등)로 추정
- `text_example_1~5`: 파싱된 보기를 그대로 배정
- `Answer`:
  - 정답지가 있으면 그 값을 최우선 사용
  - 없으면 원문의 빨간색 표시를 사용 — 객관식은 동그라미 숫자만 추출해 번호로 변환
    (다른 강조 텍스트가 섞여 있어도 무시), 서술형은 스키마 표기(`{...}` 등)로 감싸는
    기본 포맷을 시도
- `text_passage` / `text_prompt`: 밑줄 오류형 객관식(문장 속에 번호가 박힌 경우)은
  `text_prompt`로, 그 외에는 `text_passage`로 배치 (조건/보기 등 보조 문구까지
  완벽히 분리하지는 못하므로 검수 필요)

**자동으로 채우지 않는 것 (검수 단계에서 직접 입력 필요):**
- `cell_id` / `cell_title` — 문제 내용과 문법 포인트를 판단해야 하므로 항상 빈 값으로
  둡니다. '고급 설정'에 표시되는 챕터별 CELL 후보 목록을 참고해 직접 채워주세요.
- `difficulty` — 정답지에 값이 없으면 빈 값
- `image_passage` / `image_prompt` — 표가 있으면 `text_prompt`에 원문을 그대로
  덧붙이고 `[검수필요-표를 이미지로 대체 필요]` 표시만 남깁니다. 실제 이미지 제작은
  별도 프로세스입니다.
- 여러 개의 빈칸이 있는 서술형(SA)의 개별 `{}` 분리, 어법수정(IC)의 기호/수정문
  자동 인식 실패 시 — 이런 경우 `Answer` 필드가 `[검수필요-...]` 형태로 표시되니
  직접 다듬어주세요.

## 정답지 엑셀 형식 (선택 업로드)

컬럼명 `Chapter`, `문제`, `난이도`, `정답`(, `해설`)이 헤더에 있으면 자동 인식합니다
(열 순서 무관, 다른 컬럼은 무시). `Chapter` 값은 챕터가 바뀌는 행에만 채워져 있어도
됩니다 (그 아래 문항들은 다음 Chapter 값이 나오기 전까지 같은 챕터로 처리).

## 사용 순서

1. **레벨 선택** — H레벨 / E레벨
2. **(선택) 고급 설정** — 챕터 번호 지정, 챕터참고문서/스키마 엑셀 직접 교체
3. **문제 문서 업로드** — 유일한 필수 업로드. `.docx` 파일 하나
4. **(선택) 정답지 엑셀 업로드**
5. **검수 및 편집** — 표에서 직접 수정 (특히 `cell_id`/`cell_title`은 항상 직접 입력)
6. **엑셀 다운로드**

## 파일 구성

| 파일 | 역할 |
|---|---|
| `app.py` | Streamlit UI 메인 (AI 호출 없음) |
| `doc_parser.py` | 워드 문서 파싱 — 챕터/CELL 목차(문단형·표형 자동인식), 문제모음집 + 빨간색 정답 감지, 밑줄오류 지문과 실제 보기목록 구분 |
| `rule_fill.py` | AI 없이 규칙 기반으로 q_type/tp_type/Answer 등을 채우는 로직 |
| `answer_key.py` | 별도 정답지 엑셀 파싱 (챕터/문항번호로 정답·난이도 매칭) |
| `sample_xlsx.py` | 샘플 엑셀에서 스키마/공통메타/챕터-셀 후보 추출 |
| `cell_utils.py` | cell_id 자동 생성(챕터/셀 번호 치환) 유틸리티 |
| `resources/H_chapter_ref.docx` | H레벨 9개 챕터 × 3개 CELL 참고문서 |
| `resources/E_chapter_ref.docx` | E레벨 9개 챕터 × 3개 CELL 참고문서 |
| `resources/H_sample.xlsx` | 공식 원고템플릿 — 스키마 기준 |
| `resources/E_sample.xlsx` | E레벨 공통 메타데이터(level_code/cell_id 접두사) 기준 |
| `lesson_meta_H.json` / `lesson_meta_E.json` | 사이드바 "저장" 시 레벨별로 생성되는 공통 메타데이터 캐시 |

## 알려진 한계

- 문항 경계 인식은 "두 자리 이하 숫자로 시작하는 문단"을 기준으로 하므로,
  본문 안에 우연히 숫자로 시작하는 줄이 있으면 오탐할 수 있습니다.
- 빨간색 판정은 `EE0000`/`FF0000` 계열 hex 색상을 기준으로 하며, 문서마다
  강조 색상이 다르면 `doc_parser.py`의 `RED_HEXES`/`_is_red()`를 조정하세요.
- AI가 없으므로 `cell_id`/`cell_title`, 정답 표시가 없는 문항의 `Answer`,
  복잡한 서술형의 정확한 `{}` 분리는 사람이 직접 채워야 합니다.
