"""
answer_key.py
챕터별로 문항번호/난이도/정답/해설이 정리된 '정답지' 엑셀을 파싱한다.

기대하는 형태 (열 순서는 상관없음, 헤더 행에 아래 이름들이 존재하면 인식):
  학기 | 교재명 | Chapter | 문제 | 난이도 | 정답 | 해설
  (학기/교재명처럼 모르는 컬럼은 무시함)

- 'Chapter' 컬럼은 챕터 경계에서만 값이 채워짐(예: 'Ch01') -> 그 아래 문항들은
  다음 Chapter 값이 나오기 전까지 같은 챕터로 간주.
- 챕터 경계 행에는 종종 헤더가 다시 한 번(난이도/정답/해설 텍스트 그대로) 찍혀
  있는데, 이런 행은 '문제' 컬럼이 숫자가 아니므로 자동으로 건너뜀.

반환 형태: {chapter_num(int): {q_num(int): {"difficulty": str, "answer": str, "explanation": str}}}
"""

import re

import openpyxl

CIRCLED_MAP = {
    "①": "1", "②": "2", "③": "3", "④": "4", "⑤": "5",
    "⑥": "6", "⑦": "7", "⑧": "8", "⑨": "9", "⑩": "10",
}


def _clean(v):
    if v is None:
        return ""
    return str(v).replace("\xa0", " ").strip()


def normalize_answer(raw):
    """'②\xa0' 같은 원문자 정답 표기를 '2'로, '④ ⑤'는 '4,5'로 정리.
    원문자가 아닌 자유 텍스트(서술형/어법수정 등)는 공백만 정리해서 그대로 반환."""
    s = _clean(raw)
    if not s:
        return ""
    core = s.replace(" ", "")
    if core and all(ch in CIRCLED_MAP for ch in core):
        return ",".join(CIRCLED_MAP[ch] for ch in core)
    return s


def load_answer_key(path, sheet_name=0):
    """정답지 엑셀을 파싱해 {chapter_num: {q_num: {"difficulty","answer","explanation"}}} 반환.
    예상 형식과 다르면(헤더를 못 찾으면) 빈 dict를 반환한다."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet_name] if isinstance(sheet_name, str) else wb.worksheets[sheet_name]

    col_map = None
    cur_chapter = None
    result = {}

    for row in ws.iter_rows(values_only=True):
        if col_map is None:
            names = [_clean(v) for v in row]
            if "문제" in names and "정답" in names:
                col_map = {name: idx for idx, name in enumerate(row) if _clean(name)}
            continue

        chapter_idx = col_map.get("Chapter")
        qnum_idx = col_map.get("문제")
        diff_idx = col_map.get("난이도")
        ans_idx = col_map.get("정답")
        exp_idx = col_map.get("해설")

        if qnum_idx is None or ans_idx is None:
            continue

        if chapter_idx is not None and chapter_idx < len(row) and row[chapter_idx]:
            m = re.search(r"(\d+)", str(row[chapter_idx]))
            if m:
                cur_chapter = int(m.group(1))

        qnum_val = row[qnum_idx] if qnum_idx < len(row) else None
        if not isinstance(qnum_val, (int, float)):
            # 챕터 경계에 다시 찍힌 헤더 행('문제' 칸이 비어있음) 등은 건너뜀
            continue
        if cur_chapter is None:
            continue

        qnum = int(qnum_val)
        entry = {
            "difficulty": _clean(row[diff_idx]) if diff_idx is not None and diff_idx < len(row) else "",
            "answer": normalize_answer(row[ans_idx]) if ans_idx < len(row) else "",
            "explanation": _clean(row[exp_idx]) if exp_idx is not None and exp_idx < len(row) else "",
        }
        result.setdefault(cur_chapter, {})[qnum] = entry

    return result


def count_entries(answer_key):
    return sum(len(v) for v in answer_key.values())
