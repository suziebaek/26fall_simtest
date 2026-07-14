"""
cell_utils.py
cell_id 자동 생성 및 챕터별 CELL 후보 목록 구성 유틸리티.
(예전 ai_enrich.py에서 AI 호출 로직을 걷어내고 순수 유틸 함수만 남긴 모듈)
"""

import re


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


def build_candidates_for_all_chapters(chapter_ref: dict, cell_id_template: str):
    """chapter_ref(doc_parser.parse_chapter_reference 결과)와 샘플에서 뽑은
    cell_id_template을 이용해 모든 챕터의 cell_id 후보를 생성.
    반환: {chapter_num: [(cell_id, cell_title), ...]}"""
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
