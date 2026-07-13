"""
ai_enrich.py
파싱된 원시 문항(dict)을 받아 Claude API로:
  - q_type / tp_type 판별
  - difficulty(H/M/L) 산정
  - cell_id / cell_title 배정 (해당 챕터의 CELL 후보 중 선택)
  - Answer 필드를 타겟 포맷 규칙에 맞게 정리
  - text_passage / image_passage / question / text_prompt / image_prompt /
    text_example_1~5 재구성
을 수행해 최종 스키마에 맞는 dict를 반환한다.

스키마는 공식 원고템플릿(RP_H_AU_Test_G_L01_원고템플릿.xlsx) 기준:
service_code, track_code, top_cors_id, component_code, book_code, act_code,
level_code, lesson_order_seq, lesson_title, page_order_seq, cell_id, cell_title,
q_type, tp_type, difficulty, text_passage, image_passage, question, text_prompt,
image_prompt, text_example_1~5, Answer
(explanation 컬럼은 없음)
"""

import json
import re
from anthropic import Anthropic

MODEL_ID = "claude-sonnet-5"  # 필요 시 앱에서 override 가능

BASE_RULES = """당신은 한국 중고등 영어 내신/서술형 문제은행 데이터를 정해진 엑셀 스키마로 정리하는 어시스턴트입니다.
아래 [일반 규칙]과 [이 데이터셋의 실제 예시 행(few-shot)]을 참고해서, 요청된 JSON 키만 정확히 채워 출력하세요.
다른 설명, 마크다운, 코드펜스는 절대 포함하지 마세요. 예시 행에 나타난 포맷 관례(특히 Answer 필드의 중괄호/구분자 표기법)를 최우선으로 따르세요.

[일반 규칙]
- q_type: "객관식"(보기 ①~⑤ 중 선택) / "서술형"(주관식/영작/어법수정) / "단어배열"(주어진 단어 배열) 등 문항 성격에 맞게 선택 (few-shot 예시에 나타난 값들 중에서 고르는 것을 우선)
- tp_type 대략적 의미:
  - "MC": 객관식
  - "US": 주어진 단어를 모두 사용한 영작(단어 배열) — Answer는 슬래시로 구분된 단어열을 중괄호로 감싸는 표기
  - "SA": 빈칸 채우기형 서술형 — 각 빈칸을 {정답} 형태로, 여러 개면 순서대로 표기. (A),(B) 라벨이 있으면 "(A):{..} (B):{..}" 형태
  - "IC": 밑줄 (a)~(e) 중 어법 오류를 찾아 고치는 유형 — Answer는 "기호: {(x)} 수정 후: {고친표현}" 형식
  - "ES": <조건>에 맞춰 완전한 문장을 작문하는 유형 — Answer는 중괄호로 감싼 완성 문장 전체
  (정확한 표기 규칙은 few-shot 예시 행을 최우선으로 따를 것)
- difficulty: "H"(상)/"M"(중)/"L"(하). 문법 난이도, 어휘 수준, 추론 난이도를 종합 판단
- cell_id / cell_title: 제공된 해당 챕터의 CELL 후보 목록 중에서만 선택 (후보가 없으면 문제 내용에 가장 부합하는 문법 포인트를 자유롭게 판단)
- text_passage: 순수 독해 지문/대화문/글(단락)만 해당. **밑줄 친 어법 오류를 찾는 문제의 지문은 여기 넣지 않음.**
- image_passage: text_passage 내용을 스크린샷 이미지로도 제공해야 하는 경우 파일명 제안 (예: GR-H-CH01-C2_Q01_PS.png), 보통은 빈 문자열로 둠
- question: 문항 지시문 한 문장. 그룹 문항이면 공통 지시문을 사용
- text_prompt: **어법 오류를 찾는 문제의 밑줄 친 지문**, <보기>, <조건>, 요약문 빈칸, 우리말 해석 등 지문 이외의 모든 보조 텍스트. 없으면 빈 문자열, 줄바꿈은 <br/>
- image_prompt: 표(데이터 테이블 등) 형태의 자료를 이미지로 대체해야 하는 경우 파일명 제안 (예: GR-H-CH01-C3_Q19_PR.png), 아니면 빈 문자열
- text_example_1~5: 객관식 보기(있으면), 없으면 빈 문자열
- Answer: 객관식이면 정답 번호("2", 복수면 "2,4"), 그 외는 few-shot 예시의 표기 관례를 따름
"""

ANSWER_MARKED_NOTE = """
[정답 판단 방식]
원문에서 빨간색으로 표시된 텍스트가 정답으로 이미 확인되어 있습니다. 그 표시를 그대로 신뢰해서
Answer 필드 포맷팅에 반영하세요 (직접 문제를 다시 풀 필요 없음)."""

ANSWER_SOLVE_NOTE = """
[정답 판단 방식]
이 문서에는 정답이 별도로 표시되어 있지 않습니다. 문제를 직접 정확히 풀어서 정답을 판단한 뒤
Answer 필드를 규칙에 맞게 채우세요. 신중하게 검토하고, 확신이 서지 않으면 가장 근거가 명확한
답을 선택하세요."""

DEFAULT_FEW_SHOT = """[예시 1 - 객관식]
지문: Many people think money brings happiness, but research shows the opposite...
보기: excited / lonely / proud / healthy / curious
출력 예: {"q_type":"객관식","tp_type":"MC","difficulty":"H","cell_id":"GR-H-CH01-C2","cell_title":"관계대명사 that","text_passage":"Many people think money brings happiness, but research shows the opposite. Even rich people can feel ______ when they have no close friends or family to share their lives with.","image_passage":"","question":"다음 글의 빈칸에 들어갈 단어로 가장 알맞은 것은?","text_prompt":"","image_prompt":"","text_example_1":"excited","text_example_2":"lonely","text_example_3":"proud","text_example_4":"healthy","text_example_5":"curious","Answer":"2"}

[예시 2 - 단어배열(US)]
정답 문장: The movie was so interesting that I watched it twice.
출력 예 Answer 필드: "{The movie/was/so/interesting/that/I/watched/it/twice/.}"

[예시 3 - 빈칸채우기(SA)]
정답: Does he like music
출력 예 Answer 필드: "{Does} {he} {like} music?"

[예시 4 - 어법오류 수정(IC), 밑줄 지문은 text_prompt에 위치]
출력 예 일부: {"text_passage":"","text_prompt":"In the analog era, people (a) <u>struggled</u> to preserve their records...","Answer":"기호: {(a)} 수정 후: {had relied}"}

[예시 5 - 조건영작(ES)]
정답 문장: They have lost important data due to unexpected errors.
출력 예 Answer 필드: "{They have lost important data due to unexpected errors.}"
"""

DEFAULT_OUTPUT_FIELDS = [
    "cell_id", "cell_title", "q_type", "tp_type", "difficulty",
    "text_passage", "image_passage", "question", "text_prompt", "image_prompt",
    "text_example_1", "text_example_2", "text_example_3", "text_example_4",
    "text_example_5", "Answer",
]


def build_cell_id(template_id: str, target_chapter: int, target_cell: int) -> str:
    """샘플에 나온 cell_id 표기 패턴(예: 'GR-H-CH01-C2')을 이용해
    다른 챕터/셀 번호에 대한 cell_id를 유추 생성. 숫자 그룹이 2개 미만이면 원본 반환."""
    runs = list(re.finditer(r"\d+", template_id))
    if len(runs) < 2:
        return template_id
    chapter_run, cell_run = runs[-2], runs[-1]
    width = len(chapter_run.group())
    s = template_id
    s = s[: cell_run.start()] + str(target_cell) + s[cell_run.end():]
    s = s[: chapter_run.start()] + str(target_chapter).zfill(width) + s[chapter_run.end():]
    return s


def _build_system_prompt(output_fields, level):
    keys_template = ", ".join(f'"{k}": ""' for k in output_fields)
    note = ANSWER_MARKED_NOTE if level == "H" else ANSWER_SOLVE_NOTE
    return (
        BASE_RULES
        + note
        + f"\n\n[출력 형식]\n반드시 아래 키를 모두 포함한 단일 JSON 객체만 출력:\n{{{keys_template}}}\n"
    )


def _build_user_prompt(raw_q, cell_candidates):
    cand_lines = "\n".join(f"  - {cid}: {title}" for cid, title in cell_candidates) or "  (후보 없음 - 자유 판단)"
    passage = "\n".join(raw_q.get("raw_lines", []))
    options = "\n".join(raw_q.get("options", []))
    tables = "\n---\n".join(raw_q.get("tables", []))
    instruction = raw_q.get("instruction") or raw_q.get("shared_instruction") or ""
    answer_marks = " | ".join(raw_q.get("answer_marks", []))

    return f"""[이 챕터의 CELL 후보 - 후보가 있으면 이 중에서만 선택]
{cand_lines}

[문항 지시문]
{instruction}

[본문/보기 외 원문 라인 (지문, 조건, 우리말 해석 등이 섞여 있을 수 있음)]
{passage}

[객관식 보기 (①~⑤, 있는 경우)]
{options}

[문서 내 표 데이터 (있는 경우 - 원문 그대로, 필요시 image_prompt 판단에 참고)]
{tables}

[원문에서 빨간색으로 표시된 텍스트 - 있으면 정답 표시, 없으면 빈 상태]
{answer_marks}

위 정보를 바탕으로 스키마 규칙에 맞는 JSON 하나만 출력하세요."""


def enrich_question(
    client: Anthropic,
    raw_q: dict,
    cell_candidates: list,
    output_fields: list = None,
    few_shot_text: str = None,
    level: str = "H",
    model: str = MODEL_ID,
) -> dict:
    """단일 문항을 AI로 채워서 최종 스키마 dict 반환.
    output_fields: 채워야 할 JSON 키 목록 (샘플 엑셀의 '문항별' 컬럼 목록). 없으면 기본값 사용.
    few_shot_text: 샘플 엑셀 실제 행에서 뽑은 few-shot 텍스트. 없으면 기본 하드코딩 예시 사용.
    level: "H" (정답이 빨간색으로 표시됨) 또는 "E" (정답 표시 없음, AI가 직접 풀어야 함)
    """
    output_fields = output_fields or DEFAULT_OUTPUT_FIELDS
    few_shot_text = few_shot_text or DEFAULT_FEW_SHOT
    system_prompt = _build_system_prompt(output_fields, level)
    user_prompt = few_shot_text + "\n\n" + _build_user_prompt(raw_q, cell_candidates)
    resp = client.messages.create(
        model=model,
        max_tokens=1500,
        system=system_prompt,
        messages=[{"role": "user", "content": user_prompt}],
    )
    text = "".join(b.text for b in resp.content if getattr(b, "type", None) == "text")
    text = text.strip()
    text = re.sub(r"^```(json)?", "", text).strip()
    text = re.sub(r"```$", "", text).strip()
    try:
        data = json.loads(text)
    except json.JSONDecodeError:
        m = re.search(r"\{.*\}", text, re.DOTALL)
        if not m:
            raise
        data = json.loads(m.group(0))
    return data


def enrich_all(
    client: Anthropic,
    raw_questions: list,
    cell_candidates_by_chapter: dict,
    output_fields: list = None,
    few_shot_text: str = None,
    level: str = "H",
    model: str = MODEL_ID,
    progress_cb=None,
):
    """raw_questions 전체를 순회하며 enrich.
    cell_candidates_by_chapter: {chapter_num: [(cell_id, cell_title), ...]}
    """
    results = []
    total = len(raw_questions)
    for idx, rq in enumerate(raw_questions):
        ch = rq.get("chapter_num")
        cand = cell_candidates_by_chapter.get(ch, [])
        data = enrich_question(
            client, rq, cand,
            output_fields=output_fields, few_shot_text=few_shot_text,
            level=level, model=model,
        )
        data["chapter_num"] = ch
        data["q_num"] = rq.get("q_num")
        results.append(data)
        if progress_cb:
            progress_cb(idx + 1, total)
    return results


def build_candidates_for_all_chapters(chapter_ref: dict, cell_id_template: str):
    """chapter_ref(doc_parser.parse_chapter_reference 결과)와 샘플에서 뽑은
    cell_id_template을 이용해 모든 챕터의 cell_id 후보를 생성."""
    result = {}
    for ch, info in chapter_ref.items():
        cells = info.get("cells", {})
        if cell_id_template:
            result[ch] = [
                (build_cell_id(cell_id_template, ch, c), title) for c, title in sorted(cells.items())
            ]
        else:
            result[ch] = [(f"CH{ch}-C{c}", title) for c, title in sorted(cells.items())]
    return result
