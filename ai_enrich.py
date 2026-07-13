"""
ai_enrich.py
파싱된 원시 문항(dict)을 받아 Claude API로:
  - q_type / tp_type 판별
  - difficulty(H/M/L) 산정
  - explanation(해설) 생성  (단, 해설/난이도 참고문서가 있으면 그걸 우선 사용 — app.py에서 처리)
  - cell_id / cell_title 배정 (해당 챕터의 3개 CELL 후보 중 선택)
  - Answer 필드를 타겟 포맷 규칙에 맞게 정리
  - text_passage / question / text_prompt / text_example_1~5 / image_promt 재구성
을 수행해 최종 스키마에 맞는 dict를 반환한다.
"""

import json
import re
from anthropic import Anthropic

MODEL_ID = "claude-sonnet-5"  # 필요 시 앱에서 override 가능

SYSTEM_PROMPT = """당신은 한국 중고등 영어 내신/서술형 문제은행 데이터를 정해진 엑셀 스키마로 정리하는 어시스턴트입니다.
아래 [스키마 규칙]과 [예시]를 정확히 따라 JSON만 출력하세요. 다른 설명, 마크다운, 코드펜스는 절대 포함하지 마세요.

[스키마 규칙]
- q_type: "객관식" (보기 ①~⑤ 중 선택) 또는 "서술형"(주관식/영작/어법수정 등) 또는 "단어배열"(주어진 단어를 배열)
- tp_type:
  - "MC": 객관식 (text_prompt는 보통 비움, 지문이 있으면 text_passage에, 밑줄 어법 오류 지문은 text_passage에 넣고 question은 지시문)
  - "US": 주어진 단어를 모두 사용한 영작(단어 배열) — Answer는 반드시 {단어1/단어2/.../.}처럼 슬래시로 구분된 단어열을 중괄호로 감쌈. 예: {The movie/was/so/interesting/that/I/watched/it/twice/.}
  - "SA": 빈칸 채우기형 서술형 — 각 빈칸을 {정답} 형태로 표기. 빈칸이 여러 개면 순서대로 이어붙임. 예: {Does} {he} {like} music?  / 요약형처럼 (A),(B) 라벨이 있으면 (A):{ans1} (B):{ans2} 형태
  - "IC": 밑줄 (a)~(e) 중 어법 오류를 찾아 고치는 유형 — Answer는 "기호: {(x)} 수정 후: {고친표현}" 형식
  - "ES": <조건>에 맞춰 완전한 문장을 작문하는 유형 — Answer는 {완성된 문장 전체}
- difficulty: "H"(상), "M"(중), "L"(하) 중 하나. 문법 난이도, 어휘 수준, 추론 난이도를 종합 판단.
- explanation: 정답의 근거를 한국어로 2~3문장, 간결하고 핵심 단서 위주로 작성 (예시 문체 참고)
- cell_id / cell_title: 제공된 해당 챕터의 CELL 후보 목록 중 문제가 다루는 문법 포인트와 가장 가까운 것을 선택. 후보 목록 밖의 값을 만들어내지 말 것.
- text_passage: 지문(대화, 단락, 밑줄 어법 지문 등). 줄바꿈은 <br/> 사용. 지문이 없으면 빈 문자열.
- question: 문항 지시문 한 문장 (예: "다음 빈칸에 들어갈 말로 가장 알맞은 것은?"). 그룹 문항(예: [08-09])이면 공통 지시문을 사용.
- text_prompt: <보기>, <조건>, 우리말 해석 등 추가 지시/조건 텍스트. 필요 없으면 빈 문자열. 줄바꿈은 <br/> 사용.
- text_example_1~5: 객관식 보기 5개(있으면), 없으면 모두 빈 문자열
- image_promt: 문제에 표(데이터 테이블 등) 이미지가 필요한 경우 파일명을 제안 (예: GR-H-CH01-C3_Q19.png), 없으면 빈 문자열
- Answer: 객관식이면 정답 보기 번호(예: "2", 복수 정답이면 "2,4"), 그 외에는 위 tp_type 규칙에 따른 포맷

[출력 형식]
반드시 아래 키를 모두 포함한 단일 JSON 객체만 출력:
{"q_type":"", "tp_type":"", "difficulty":"", "cell_id":"", "cell_title":"", "text_passage":"", "question":"", "text_prompt":"", "image_promt":"", "text_example_1":"", "text_example_2":"", "text_example_3":"", "text_example_4":"", "text_example_5":"", "Answer":"", "explanation":""}
"""

FEW_SHOT = """[예시 1 - 객관식]
입력 지시문: 다음 글의 빈칸에 들어갈 단어로 가장 알맞은 것은?
보기: excited / lonely / proud / healthy / curious
지문: Many people think money brings happiness... feel ______ when they have no close friends...
빨간색 표시(정답 후보): ②
출력 예: {"q_type":"객관식","tp_type":"MC","difficulty":"H","cell_id":"GR-H-CH01-C2","cell_title":"관계대명사 that","text_passage":"Many people think money brings happiness, but research shows the opposite. Even rich people can feel ______ when they have no close friends or family to share their lives with.","question":"다음 글의 빈칸에 들어갈 단어로 가장 알맞은 것은?","text_prompt":"","image_promt":"","text_example_1":"excited","text_example_2":"lonely","text_example_3":"proud","text_example_4":"healthy","text_example_5":"curious","Answer":"2","explanation":"'no close friends or family to share' 부분이 결정적 단서. '외로움'을 의미하는 lonely가 정답."}

[예시 2 - 단어배열(US)]
입력 지시문: 다음 우리말과 일치하도록 주어진 단어를 모두 사용하여 영작하시오. (필요시 어형 변화 가능)
우리말: 그 영화는 너무 재미있어서 나는 두 번 봤다.
주어진 단어 배열 순서(빨간색 정답 문장): The movie was so interesting that I watched it twice.
출력 예 Answer 필드: "{The movie/was/so/interesting/that/I/watched/it/twice/.}"

[예시 3 - 빈칸채우기(SA)]
빨간색 정답: Does he like music
출력 예 Answer 필드: "{Does} {he} {like} music?"

[예시 4 - 어법오류 수정(IC)]
빨간색 표시: 기호 (a), 수정 후 had relied
출력 예 Answer 필드: "기호: {(a)} 수정 후: {had relied}"

[예시 5 - 조건영작(ES)]
빨간색 정답 문장: They have lost important data due to unexpected errors.
출력 예 Answer 필드: "{They have lost important data due to unexpected errors.}"
"""


def _build_user_prompt(raw_q, cell_candidates):
    cand_lines = "\n".join(f"  - {cid}: {title}" for cid, title in cell_candidates)
    passage = "\n".join(raw_q.get("raw_lines", []))
    options = "\n".join(raw_q.get("options", []))
    tables = "\n---\n".join(raw_q.get("tables", []))
    instruction = raw_q.get("instruction") or raw_q.get("shared_instruction") or ""
    answer_marks = " | ".join(raw_q.get("answer_marks", []))

    return f"""[이 챕터의 CELL 후보 - 이 중에서만 선택]
{cand_lines}

[문항 지시문]
{instruction}

[본문/보기 외 원문 라인 (지문, 조건, 우리말 해석 등이 섞여 있을 수 있음)]
{passage}

[객관식 보기 (①~⑤, 있는 경우)]
{options}

[문서 내 표 데이터 (있는 경우 - 원문 그대로, 필요시 image_promt 판단에 참고)]
{tables}

[원문에서 빨간색으로 표시된 텍스트 = 정답 단서]
{answer_marks}

위 정보를 바탕으로 스키마 규칙에 맞는 JSON 하나만 출력하세요."""


def enrich_question(client: Anthropic, raw_q: dict, cell_candidates: list, model: str = MODEL_ID) -> dict:
    """단일 문항을 AI로 채워서 최종 스키마 dict 반환"""
    user_prompt = FEW_SHOT + "\n\n" + _build_user_prompt(raw_q, cell_candidates)
    resp = client.messages.create(
        model=model,
        max_tokens=1500,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_prompt}],
    )
    text = "".join(b.text for b in resp.content if getattr(b, "type", None) == "text")
    text = text.strip()
    text = re.sub(r"^```(json)?", "", text).strip()
    text = re.sub(r"```$", "", text).strip()
    try:
        data = json.loads(text)
    except json.JSONDecodeError:
        # 모델이 JSON 앞뒤로 잡담을 붙였을 경우 { ... } 블록만 추출
        m = re.search(r"\{.*\}", text, re.DOTALL)
        if not m:
            raise
        data = json.loads(m.group(0))
    return data


def enrich_all(client: Anthropic, raw_questions: list, chapter_ref: dict, model: str = MODEL_ID, progress_cb=None):
    """raw_questions 전체를 순회하며 enrich. chapter_ref는 doc_parser.parse_chapter_reference 결과."""
    results = []
    total = len(raw_questions)
    for idx, rq in enumerate(raw_questions):
        ch = rq.get("chapter_num")
        cells = chapter_ref.get(ch, {}).get("cells", {})
        cand = [
            (f"GR-H-CH{ch:02d}-C{c}", title) for c, title in sorted(cells.items())
        ]
        data = enrich_question(client, rq, cand, model=model)
        data["chapter_num"] = ch
        data["q_num"] = rq.get("q_num")
        results.append(data)
        if progress_cb:
            progress_cb(idx + 1, total)
    return results
