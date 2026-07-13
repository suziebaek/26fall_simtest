"""
sample_xlsx.py
샘플/템플릿 엑셀(예: 공식 원고템플릿)을 스키마 기준으로 사용하기 위한 모듈.

- 헤더(컬럼 순서)를 그대로 타겟 스키마로 사용 -> 코드 수정 없이 다른 포맷의
  엑셀도 업로드만 하면 그 형식에 맞춰 동작.
- 모든 행에서 값이 (거의) 동일한 컬럼은 '공통 메타데이터'로 자동 판별.
- 나머지(행마다 달라지는) 컬럼은 AI가 채워야 하는 '문항별 필드'.
- 샘플 행 자체를 few-shot 예시로 재사용해서 AI가 이 데이터셋 고유의
  포맷 관례(Answer 표기 등)를 그대로 학습하도록 함.
- 챕터/CELL 참고문서가 없는 레벨(예: E)을 위해, 샘플 자체에서 챕터 번호와
  cell_id/cell_title 후보쌍을 추출하는 기능도 제공.
"""

import re
import pandas as pd


def load_sample(path, sheet_name=0):
    df = pd.read_excel(path, sheet_name=sheet_name, dtype=str).fillna("")
    columns = list(df.columns)

    n_rows = max(len(df), 1)
    constant_cols = {}
    varying_cols = []
    for col in columns:
        non_empty = [v for v in df[col].tolist() if str(v).strip() != ""]
        vals = set(non_empty)
        coverage = len(non_empty) / n_rows
        # '공통 메타데이터'로 볼 조건: 값이 하나뿐이고, 대부분의 행에 실제로 채워져 있을 것
        # (한두 행에만 우연히 값이 있는 image_prompt 같은 문항별 필드는 제외)
        if len(vals) <= 1 and coverage >= 0.9 and len(non_empty) > 0:
            constant_cols[col] = next(iter(vals))
        else:
            varying_cols.append(col)

    rows = df.to_dict(orient="records")
    return {
        "columns": columns,
        "constant_cols": constant_cols,
        "varying_cols": varying_cols,
        "rows": rows,
        "df": df,
    }


def build_few_shot_from_samples(sample_rows, varying_cols, max_examples=6):
    """행을 다양하게(q_type/tp_type 기준) 뽑아 few-shot 텍스트 생성"""
    if not sample_rows:
        return ""

    seen_keys = set()
    picked = []
    for row in sample_rows:
        key = (row.get("q_type", ""), row.get("tp_type", ""))
        if key in seen_keys:
            continue
        seen_keys.add(key)
        picked.append(row)
        if len(picked) >= max_examples:
            break
    if len(picked) < max_examples:
        for row in sample_rows:
            if row not in picked:
                picked.append(row)
            if len(picked) >= max_examples:
                break

    blocks = []
    for row in picked:
        fields = {c: row.get(c, "") for c in varying_cols}
        lines = "\n".join(f'  "{k}": {fields[k]!r}' for k in fields)
        blocks.append(f"[예시]\n{lines}")
    return "\n\n".join(blocks)


def extract_chapter_cells_from_sample(sample_info):
    """샘플 엑셀 자체에서 (챕터 참고문서가 없는 레벨을 위해) 챕터 번호와
    cell_id/cell_title 후보쌍을 추출.
    반환: (chapter_num:int|None, [(cell_id, cell_title), ...])
    """
    const = sample_info.get("constant_cols", {})
    lesson_title = const.get("lesson_title", "")
    m = re.search(r"(\d+)", lesson_title)
    chapter_num = int(m.group(1)) if m else None

    pairs = []
    seen = set()
    for row in sample_info.get("rows", []):
        cid, ctitle = row.get("cell_id", ""), row.get("cell_title", "")
        if cid and (cid, ctitle) not in seen:
            seen.add((cid, ctitle))
            pairs.append((cid, ctitle))
    return chapter_num, pairs
