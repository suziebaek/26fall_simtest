"""
doc_parser.py
워드 문서(챕터/셀 참고문서, 문제모음집)를 파싱하여 구조화된 dict 리스트로 변환.

핵심 아이디어
-------------
1. 챕터 참고문서(FG_H_챕터.docx 형태)
   "CHAPTER 0N  제목" / "CELL n   제목" 패턴을 순서대로 읽어
   {chapter_num: {"title": ..., "cells": {1: title1, 2: title2, 3: title3}}} 형태로 반환.

2. 문제모음집(FG_H_Sim_Test.docx 형태)
   - "Chapter" 단독 문단 다음에 "01  동사의 종류" 같은 챕터 타이틀이 옴 -> 챕터 경계
   - "01 다음 빈칸에...", "12", "[08-09] ..." 같은 문항 번호로 문항 경계를 잡음
   - 빨간색(EE0000/FF0000 계열) 텍스트 = 정답 표시
       - 동그라미 숫자(①~⑤) 단독이면 객관식 정답 번호
       - 문장/구(자유 텍스트)이면 서술형 정답
   - 문서 안에 섞여 있는 표(Table)는 등장 순서 그대로 문항에 붙임(참고용 원문 보존, 이미지 문항일 경우 image_promt 후보로 사용)
"""

import re
from dataclasses import dataclass, field
from docx import Document
from docx.oxml.ns import qn

CIRCLED = "①②③④⑤⑥⑦⑧⑨⑩"
RED_HEXES = {"EE0000", "FF0000", "CC0000", "C00000", "FF0033"}


def _is_red(rgb):
    if rgb is None:
        return False
    s = str(rgb).upper()
    if s in RED_HEXES:
        return True
    # 일반화: R값이 매우 높고 G,B가 낮으면 붉은색 계열로 간주
    try:
        r, g, b = int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16)
        return r >= 180 and g <= 80 and b <= 80
    except Exception:
        return False


def _para_red_spans(paragraph):
    """문단 내에서 빨간색으로 표시된 run 텍스트들을 순서대로 반환"""
    spans = []
    for r in paragraph.runs:
        color = r.font.color.rgb if (r.font.color and r.font.color.rgb) else None
        if _is_red(color) and r.text.strip():
            spans.append(r.text)
    return spans


CHAPTER_TITLE_RE = re.compile(r"^(\d{2})\s+(.*)$")
CELL_RE = re.compile(r"^CELL\s*(\d+)\s+(.*)$", re.IGNORECASE)


def parse_chapter_reference(path):
    """FG_H_챕터.docx 같은 '챕터/셀 목차' 문서를 파싱.
    반환: {chapter_num(int): {"title": str, "cells": {1: str, 2: str, 3: str}}}
    """
    doc = Document(path)
    result = {}
    cur_chapter = None
    pending_chapter_word = False  # 'CHAPTER'만 있고 번호/제목이 다음 줄에 오는 경우 대비

    for p in doc.paragraphs:
        t = p.text.strip()
        if not t:
            continue
        m_chapter = re.match(r"^CHAPTER\s+(\d{2})\s+(.*)$", t, re.IGNORECASE)
        if m_chapter:
            cur_chapter = int(m_chapter.group(1))
            result[cur_chapter] = {"title": m_chapter.group(2).strip(), "cells": {}}
            continue
        m_cell = CELL_RE.match(t)
        if m_cell and cur_chapter is not None:
            cell_num = int(m_cell.group(1))
            result[cur_chapter]["cells"][cell_num] = m_cell.group(2).strip()
            continue
    return result


def _iter_block_items(doc):
    """문서 본문(paragraph, table)을 실제 등장 순서대로 순회"""
    body = doc.element.body
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            from docx.text.paragraph import Paragraph
            yield "p", Paragraph(child, doc)
        elif child.tag == qn("w:tbl"):
            from docx.table import Table
            yield "tbl", Table(child, doc)


def _table_to_tsv(table):
    rows = []
    for row in table.rows:
        rows.append(" | ".join(c.text.strip().replace("\n", " ") for c in row.cells))
    return "\n".join(rows)


QNUM_RE = re.compile(r"^(\d{1,2})(?:\s+(.*))?$")
QRANGE_RE = re.compile(r"^\[(\d{1,2})-(\d{1,2})\]\s*(.*)$")
CHAPTER_MARK_RE = re.compile(r"^(\d{2})\s+(.+)$")


def parse_question_bank(path):
    """
    문제모음집 워드 문서를 파싱해서 챕터별 문항 리스트를 반환.
    반환: list of dict, 각 dict:
      {
        "chapter_num": int,
        "chapter_title": str,
        "q_num": int,
        "group_range": (start,end) or None,
        "shared_instruction": str or "",   # [08-09] 같은 그룹 공통 지시문
        "instruction": str,                # 문항 지시문(질문)
        "raw_lines": [str, ...],           # 지문/보기/조건 등 원문 라인(순서 보존)
        "tables": [tsv_str, ...],
        "options": [str, ...],             # ①~⑤ 보기 (있는 경우)
        "answer_marks": [str, ...],        # 빨간색으로 표시된 원문(정답 후보)
      }
    파싱은 휴리스틱이므로 반드시 Streamlit UI에서 사람이 검수해야 함.
    """
    doc = Document(path)
    items = list(_iter_block_items(doc))

    questions = []
    cur_chapter_num = None
    cur_chapter_title = None
    expecting_chapter_title = False

    cur_q = None
    cur_group_range = None
    cur_shared_instruction = ""

    def flush():
        if cur_q is not None:
            questions.append(cur_q)

    i = 0
    n = len(items)
    while i < n:
        kind, obj = items[i]
        if kind == "tbl":
            tsv = _table_to_tsv(obj)
            if cur_q is not None:
                cur_q["tables"].append(tsv)
            i += 1
            continue

        text = obj.text.strip()
        if not text:
            i += 1
            continue

        # 챕터 헤더: 'Chapter' 단독 문단 다음에 '01  제목'
        if text.lower() == "chapter":
            expecting_chapter_title = True
            i += 1
            continue
        if expecting_chapter_title:
            m = CHAPTER_MARK_RE.match(text)
            if m:
                flush()
                cur_q = None
                cur_chapter_num = int(m.group(1))
                cur_chapter_title = m.group(2).strip()
                expecting_chapter_title = False
                cur_shared_instruction = ""
                i += 1
                continue
            expecting_chapter_title = False  # 예상과 다르면 무시하고 계속 진행

        # 그룹 헤더: [08-09] 다음 중 ...
        m_range = QRANGE_RE.match(text)
        if m_range:
            flush()
            cur_q = None
            cur_group_range = (int(m_range.group(1)), int(m_range.group(2)))
            cur_shared_instruction = m_range.group(3).strip()
            i += 1
            continue

        # 문항 번호: '01 다음 ~' 또는 단독 '12'
        m_qnum = QNUM_RE.match(text)
        if m_qnum and cur_chapter_num is not None:
            qnum = int(m_qnum.group(1))
            rest = (m_qnum.group(2) or "").strip()
            # 그룹 범위 안에 있는 번호인지 확인
            in_group = cur_group_range and cur_group_range[0] <= qnum <= cur_group_range[1]
            if in_group or rest or qnum <= 30:
                # 새 문항 시작으로 판단 (오탐 방지를 위해 30 이하 숫자만 문항번호로 인정)
                flush()
                cur_q = {
                    "chapter_num": cur_chapter_num,
                    "chapter_title": cur_chapter_title,
                    "q_num": qnum,
                    "group_range": cur_group_range if in_group else None,
                    "shared_instruction": cur_shared_instruction if in_group else "",
                    "instruction": rest,
                    "raw_lines": [],
                    "tables": [],
                    "options": [],
                    "answer_marks": [],
                }
                # 그룹 범위가 끝났으면 초기화(다음 그룹 헤더 전까지는 유지)
                i += 1
                continue

        # 그 외 줄: 현재 문항에 귀속
        if cur_q is not None:
            reds = _para_red_spans(obj)
            if reds:
                cur_q["answer_marks"].extend(reds)
            # 보기(①~⑤)가 한 줄에 여러 개 있을 수 있으므로 분리
            if any(c in text for c in CIRCLED):
                parts = re.split(r"(?=[①②③④⑤⑥⑦⑧⑨⑩])", text)
                for part in parts:
                    part = part.strip()
                    if part:
                        cur_q["options"].append(part)
            else:
                cur_q["raw_lines"].append(text)
        i += 1

    flush()
    return questions
