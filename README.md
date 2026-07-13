# 워드 문제 → 엑셀 문제은행 자동 변환기

**문제 문서 하나만 업로드**하면, 내장된 공식 스키마(원고템플릿)와 챕터/CELL
참고자료를 자동 적용해서 타겟 엑셀 포맷으로 변환해주는 Streamlit 앱입니다.

## 설치 및 실행

```bash
pip install -r requirements.txt
streamlit run app.py
```

## API 키 설정 (둘 중 하나)

1. 앱 실행 후 왼쪽 사이드바에 직접 입력
2. `.streamlit/secrets.toml`:
   ```toml
   ANTHROPIC_API_KEY = "sk-ant-..."
   ```

## 스키마 (내장 · 두 레벨 공용)

`resources/H_sample.xlsx` (공식 원고템플릿, `RP_H_AU_Test_G_L01_원고템플릿.xlsx`)을
기준으로 아래 26개 컬럼을 사용합니다. **`explanation`(해설) 컬럼은 없습니다.**

```
service_code, track_code, top_cors_id, component_code, book_code, act_code,
level_code, lesson_order_seq, lesson_title, page_order_seq, cell_id, cell_title,
q_type, tp_type, difficulty, text_passage, image_passage, question, text_prompt,
image_prompt, text_example_1~5, Answer
```

필드 관례:
- `text_passage`: 순수 독해 지문/대화문만. **밑줄 어법오류 지문은 여기 안 들어감.**
- `text_prompt`: 어법오류 밑줄 지문, `<보기>`, `<조건>`, 우리말 해석 등 보조 텍스트.
- `image_passage` / `image_prompt`: 지문 또는 표를 이미지로 대체해야 할 때의 파일명.
- `Answer`: 유형별 표기(예: 객관식 `"2"`, 복수정답 `"2,4"`, 단어배열 `{a/b/c/.}`,
  빈칸 `{ans}`, 어법수정 `기호: {(a)} 수정 후: {..}`) — `ai_enrich.py`의 few-shot과
  `resources/H_sample.xlsx` 실제 행을 참고해 AI가 그대로 학습합니다.

## 레벨별 차이

| 항목 | H레벨 | E레벨 |
|---|---|---|
| 스키마(컬럼 구성) | 26개 컬럼 | **H와 완전히 동일** |
| 공통 메타데이터 | `level_code=TO_G_H_AU`, `cell_id` 접두사 `GR-H-CH..` | `level_code=TO_G_E_AU`, `cell_id` 접두사 `GR-E-CH..` (레벨 식별자만 치환) |
| 챕터/CELL 참고문서 | `resources/H_chapter_ref.docx` 내장 (9개 챕터 전체) | 없음 — `resources/E_sample.xlsx`에서 챕터1의 cell_id/cell_title만 자동 추출해 사용 |
| 정답 표시 | 문서에 **빨간 글씨**로 정답이 이미 표시되어 있음 → AI는 그대로 옮겨 포맷팅만 함 | 문서에 정답 표시가 **없음** → AI가 문제를 **직접 풀어서** 정답을 판단 (검수 필수) |
| 챕터 범위 | 한 파일에 여러 챕터가 있어도 `Chapter` 헤더로 자동 구분 | 보통 한 파일 = 한 챕터 (챕터 헤더 없어도 '고급 설정'에서 지정한 챕터 번호로 처리) |

`resources/E_sample.xlsx`는 H 공식템플릿과 동일한 26개 컬럼 구조로 재구성했고,
`level_code`/`cell_id`만 H→E로 치환해 생성했습니다 (원본 E 참고자료에 있던
`explanation` 컬럼 등 스키마 차이는 모두 제거됨). E레벨도 스키마 자체는 H와
동일하다고 확인됨. 다만 전용 챕터참고문서가 아직 없으므로, 여러 챕터를 다루게
되면 **고급 설정에서 챕터/CELL 참고문서(.docx, H와 동일한 `CHAPTER 0N 제목` /
`CELL n 제목` 형식)를 올려주시면 정확도가 올라갑니다.**

## 사용 순서

1. **레벨 선택** — H레벨 / E레벨 라디오 버튼
2. **(선택) 고급 설정** — 챕터 번호 지정, 챕터참고문서/스키마 엑셀 직접 교체
3. **문제 문서 업로드** — 유일한 필수 업로드. `.docx` 파일 하나
4. **AI 자동 채우기 실행** — Claude API 호출 (사이드바에 API 키 필요)
5. **검수 및 편집** — 표에서 직접 수정 가능 (특히 E레벨은 AI가 스스로 정답을
   판단하므로 꼼꼼히 확인 권장)
6. **엑셀 다운로드**

## 파일 구성

| 파일 | 역할 |
|---|---|
| `app.py` | Streamlit UI 메인 |
| `doc_parser.py` | 워드 문서 파싱 (챕터/CELL 목차, 문제모음집 + 빨간색 정답 감지, 챕터헤더 없는 단일챕터 문서도 지원) |
| `sample_xlsx.py` | 샘플 엑셀에서 스키마/공통메타/few-shot/챕터-셀 후보 추출 |
| `ai_enrich.py` | Claude API 프롬프트 및 스키마 매핑 로직 (레벨별 정답판단 방식 분기) |
| `resources/H_chapter_ref.docx` | H레벨 9개 챕터 × 3개 CELL 참고문서 (내장) |
| `resources/H_sample.xlsx` | 공식 원고템플릿 — 스키마/포맷 기준 (내장, 두 레벨 공용) |
| `resources/H_guide.xlsx` | 원고 작성 가이드 (참고용) |
| `resources/E_sample.xlsx` | E레벨 챕터1 cell_id/cell_title 추출용 (내장) |
| `lesson_meta.json` | 사이드바 "저장" 시 생성되는 공통 메타데이터 캐시 |

## 알려진 한계

- 문항 경계 인식은 "두 자리 이하 숫자로 시작하는 문단"을 기준으로 하므로,
  본문 안에 우연히 숫자로 시작하는 줄이 있으면 오탐할 수 있습니다.
- 빨간색 판정은 `EE0000`/`FF0000` 계열 hex 색상을 기준으로 하며, 문서마다
  강조 색상이 다르면 `doc_parser.py`의 `RED_HEXES`/`_is_red()`를 조정하세요.
- E레벨은 AI가 직접 정답을 판단하므로 H레벨보다 오류 가능성이 높습니다 —
  검수 단계를 건너뛰지 마세요.
- `image_passage`/`image_prompt`에 제안된 파일명은 실제 이미지 생성 여부와
  무관한 "제안"일 뿐이며, 실제 이미지 제작/업로드는 별도 프로세스가 필요합니다.
