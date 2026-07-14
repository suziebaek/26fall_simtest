"""
rule_fill.py
AI 호출 없이, 파싱된 원시 문항(doc_parser 결과)과 (있으면) 정답지(answer_key)만으로
타겟 스키마 필드를 규칙 기반으로 채운다.

채우는 필드: q_type, tp_type, difficulty, text_passage, image_passage, question,
            text_prompt, image_prompt, text_example_1~5, Answer
(cell_id/cell_title은 자동 판단하지 않음 — 후보 목록을 참고해 검수 단계에서 직접 채워야 함)

원칙:
- 100% 정확을 보장하지 않는 휴리스틱이므로, 결과는 반드시 사람이 검수/수정하는 것을 전제로 한다.
- 정답지(answer_key)에 값이 있으면 그 값을 최우선으로 사용한다.
- 정답지가 없으면 원문의 빨간색 표시(answer_marks)를 정리해서 사용한다.
"""

import re

from answer_key import CIRCLED_MAP

UNDERLINE_ERROR_KEYWORDS = ["어법상 틀린", "어법상 옳지", "어법상 어색한", "밑줄 친 부분 중"]
IC_KEYWORDS = ["바르게 고치", "찾아 고치", "옳게 고치"]
US_KEYWORDS = ["단어를 모두 사용", "단어들을 배열", "단어를 배열", "바르게 배열", "순서대로 배열"]
ES_KEYWORDS = ["<조건>", "〈조건〉", "＜조건＞", "조건에 맞"]

_WORD_OR_PUNCT_RE = re.compile(r"[A-Za-z가-힣0-9'’]+|[.,!?]")


def _clean(s):
    return (s or "").strip()


def _strip_option_marker(opt):
    return re.sub(r"^[①②③④⑤⑥⑦⑧⑨⑩]\s*", "", opt).strip()


def guess_q_type_tp_type(raw_q):
    """옵션 존재 여부와 지시문 키워드로 q_type/tp_type을 추정."""
    instruction = _clean(raw_q.get("instruction") or raw_q.get("shared_instruction"))
    all_text = instruction + " " + " ".join(raw_q.get("raw_lines", []))

    if raw_q.get("options"):
        return "객관식", "MC"
    if raw_q.get("has_inline_markers") and any(k in all_text for k in UNDERLINE_ERROR_KEYWORDS):
        # 보기 목록 없이 밑줄 오류 지문 안의 번호 하나를 답으로 고르는 객관식
        return "객관식", "MC"
    if any(k in all_text for k in IC_KEYWORDS):
        return "서술형", "IC"
    if any(k in all_text for k in US_KEYWORDS):
        return "서술형", "US"
    if any(k in all_text for k in ES_KEYWORDS):
        return "서술형", "ES"
    return "서술형", "SA"


def _format_answer_text(text, tp_type):
    """정답지/빨간색표시에서 얻은 '순수 텍스트' 정답을 스키마의 중괄호 표기로 감싼다.
    (완벽하지 않은 휴리스틱 - 특히 여러 개의 빈칸(SA)이나 (A)/(B) 라벨은 검수 필요)"""
    text = _clean(text)
    if not text:
        return ""

    # 이미 스키마 표기(중괄호, '기호:', '(A):')가 되어 있으면 그대로 둠
    if "{" in text or text.startswith("기호"):
        return text

    if tp_type == "US":
        tokens = _WORD_OR_PUNCT_RE.findall(text)
        return "{" + "/".join(tokens) + "}" if tokens else "{" + text + "}"

    if tp_type == "IC":
        m = re.search(r"\(([a-eA-E])\)\s*(.*)", text)
        if m:
            return f"기호: {{({m.group(1).lower()})}} 수정 후: {{{m.group(2).strip()}}}"
        return f"[검수필요-기호/수정문 확인] {{{text}}}"

    # SA, ES, 그 외 서술형: 통째로 하나의 빈칸으로 감쌈 (빈칸이 여러 개면 수동 분리 필요)
    return "{" + text + "}"


def _format_answer_mc(marks_or_text):
    """빨간색 표시 안에 정답 원문자 외의 다른 강조 텍스트(예: 근거로 같이 빨갛게 표시된
    보기 문장의 앞글자 등)가 섞여 있어도, 실제 동그라미 숫자만 뽑아 정답으로 사용."""
    circled_found = [ch for ch in marks_or_text if ch in CIRCLED_MAP]
    if circled_found:
        seen = []
        for ch in circled_found:
            n = CIRCLED_MAP[ch]
            if n not in seen:
                seen.append(n)
        return ",".join(seen)
    return _clean(marks_or_text)


def build_row_fields(raw_q, answer_key_entry=None, cell_candidates=None):
    """단일 문항에 대해 VARYING_FIELDS(cell_id 제외 나머지)를 규칙 기반으로 채워 dict 반환.
    answer_key_entry: {"difficulty","answer","explanation"} 또는 None
    """
    q_type, tp_type = guess_q_type_tp_type(raw_q)
    instruction = _clean(raw_q.get("instruction") or raw_q.get("shared_instruction"))

    # text_passage vs text_prompt 배치:
    # - MC + 밑줄오류(문장 속에 번호가 박힌 경우) -> 지문은 text_prompt
    # - 그 외 -> 지문은 text_passage, <조건>/우리말 등 보조문구는 text_prompt로 남을 수 있음
    #   (규칙만으로 완벽히 분리하기 어려워 raw_lines를 우선 text_passage에 몰아넣고 검수에 맡김)
    joined_lines = "<br/>".join(raw_q.get("raw_lines", []))
    if tp_type == "MC" and raw_q.get("has_inline_markers"):
        text_passage = ""
        text_prompt = joined_lines
    else:
        text_passage = joined_lines
        text_prompt = ""

    tables = raw_q.get("tables") or []
    image_prompt = ""
    if tables:
        # 표 데이터는 실제 이미지 파일을 자동 생성할 수 없으므로, 값이 유실되지 않도록
        # text_prompt 뒤에 원문 그대로 덧붙여 검수 시 참고/수동 이미지화할 수 있게 함
        table_note = "<br/>".join(f"[표{i+1}] {t}" for i, t in enumerate(tables))
        text_prompt = (text_prompt + "<br/>" + table_note).strip("<br/>") if text_prompt else table_note
        image_prompt = "[검수필요-표를 이미지로 대체 필요]"

    # Answer: 정답지 우선, 없으면 원문 빨간색 표시
    if answer_key_entry and answer_key_entry.get("answer"):
        raw_answer = answer_key_entry["answer"]
    else:
        raw_answer = " ".join(m.strip() for m in raw_q.get("answer_marks", []) if m.strip())

    if tp_type == "MC":
        answer = _format_answer_mc(raw_answer)
    else:
        answer = _format_answer_text(raw_answer, tp_type)

    difficulty = ""
    if answer_key_entry and answer_key_entry.get("difficulty"):
        difficulty = answer_key_entry["difficulty"]

    options = [_strip_option_marker(o) for o in raw_q.get("options", [])]
    text_examples = {f"text_example_{i+1}": (options[i] if i < len(options) else "") for i in range(5)}

    return {
        "cell_id": "",
        "cell_title": "",
        "q_type": q_type,
        "tp_type": tp_type,
        "difficulty": difficulty,
        "text_passage": text_passage,
        "image_passage": "",
        "question": instruction,
        "text_prompt": text_prompt,
        "image_prompt": image_prompt,
        **text_examples,
        "Answer": answer,
    }


def build_all(raw_questions, answer_key=None):
    """raw_questions 전체에 규칙 기반 필드 채우기를 적용."""
    answer_key = answer_key or {}
    results = []
    for rq in raw_questions:
        ch = rq.get("chapter_num")
        entry = answer_key.get(ch, {}).get(rq.get("q_num"))
        data = build_row_fields(rq, answer_key_entry=entry)
        data["chapter_num"] = ch
        data["q_num"] = rq.get("q_num")
        results.append(data)
    return results
