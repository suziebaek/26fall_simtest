import io
import json
import os
from collections import Counter

import pandas as pd
import streamlit as st

from doc_parser import parse_chapter_reference, parse_question_bank
from cell_utils import build_candidates_for_all_chapters
from sample_xlsx import load_sample, extract_chapter_cells_from_sample
from answer_key import load_answer_key, count_entries
from rule_fill import build_all

st.set_page_config(page_title="문제은행 자동 변환기", layout="wide")

RESOURCES_DIR = "resources"
H_CHAPTER_REF = os.path.join(RESOURCES_DIR, "H_chapter_ref.docx")
E_CHAPTER_REF = os.path.join(RESOURCES_DIR, "E_chapter_ref.docx")
H_SAMPLE = os.path.join(RESOURCES_DIR, "H_sample.xlsx")   # 공식 원고템플릿 (스키마 기준)
E_SAMPLE = os.path.join(RESOURCES_DIR, "E_sample.xlsx")   # E레벨 템플릿 (H와 동일 스키마, level_code/cell_id만 E)


@st.cache_data
def _load_sample_cached(path):
    return load_sample(path)


@st.cache_data
def _load_h_chapter_ref():
    return parse_chapter_reference(H_CHAPTER_REF)


@st.cache_data
def _load_e_chapter_ref():
    return parse_chapter_reference(E_CHAPTER_REF)


# ---------- 세션 상태 초기화 ----------
for key, default in [
    ("raw_questions", None),
    ("lesson_meta_by_level", {}),
    ("answer_key", {}),
]:
    if key not in st.session_state:
        st.session_state[key] = default

st.title("📘 워드 문제 → 엑셀 문제은행 변환기")
st.caption("문제 문서 하나만 올리면, 내장된 공식 템플릿 스키마 + 챕터 참고자료를 자동 적용해 변환합니다. (AI 호출 없음 — 규칙 기반 자동 채움 + 직접 검수/편집)")

# ================= STEP 1: 레벨 선택 =================
st.header("1️⃣ 레벨 선택")
level = st.radio("레벨을 선택하세요", ["H레벨", "E레벨"], horizontal=True)
level_code = "H" if level.startswith("H") else "E"

# 두 레벨 모두 스키마(컬럼 구성/포맷 관례)는 동일 -> H 템플릿을 스키마 기준으로 사용.
# 공통 메타데이터(level_code, cell_id 접두사 등)는 레벨별 파일에서 각각 로드.
schema_sample = _load_sample_cached(H_SAMPLE)
level_sample = schema_sample if level_code == "H" else _load_sample_cached(E_SAMPLE)

TARGET_COLUMNS = schema_sample["columns"]
VARYING_FIELDS = [c for c in schema_sample["varying_cols"] if c != "page_order_seq"]
CELL_ID_TEMPLATE = next((r["cell_id"] for r in level_sample["rows"] if r.get("cell_id")), None)

CONFIG_PATH = f"lesson_meta_{level_code}.json"
if level_code not in st.session_state.lesson_meta_by_level:
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, encoding="utf-8") as f:
            st.session_state.lesson_meta_by_level[level_code] = json.load(f)
    else:
        st.session_state.lesson_meta_by_level[level_code] = dict(level_sample["constant_cols"])

if level_code == "H":
    st.success("H레벨: 챕터/CELL 참고문서(9개 챕터)와 정답 색상표시 인식이 내장되어 있습니다. 문제 문서만 올리면 됩니다.")
    chapter_ref = _load_h_chapter_ref()
    candidates_by_chapter = build_candidates_for_all_chapters(chapter_ref, CELL_ID_TEMPLATE)
    default_chapter_num = 1
else:
    if os.path.exists(E_CHAPTER_REF):
        st.success("E레벨: 챕터/CELL 참고문서(9개 챕터)가 내장되어 있습니다. 문제 문서만 올리면 됩니다.")
        chapter_ref = _load_e_chapter_ref()
        candidates_by_chapter = build_candidates_for_all_chapters(chapter_ref, CELL_ID_TEMPLATE)
        default_chapter_num = 1
    else:
        st.warning(
            "E레벨: 전용 챕터/CELL 참고문서를 찾지 못해, 내장된 E 템플릿에서 챕터의 cell_id/cell_title만 "
            "자동 추출해 후보로 사용합니다. 아래 '고급 설정'에서 챕터 참고문서를 직접 올릴 수 있습니다."
        )
        e_chapter_num, e_cells = extract_chapter_cells_from_sample(level_sample)
        chapter_ref = {}
        candidates_by_chapter = {e_chapter_num: e_cells} if e_chapter_num else {}
        default_chapter_num = e_chapter_num or 1

st.info(
    "🔎 **cell_id/cell_title, 정답(Answer), 지문/보기 배치는 규칙 기반으로 초안만 채워집니다.** "
    "AI가 문제를 풀거나 판단하지 않으므로, 아래 4️⃣ 검수 단계에서 반드시 직접 확인/수정해주세요."
)

with st.expander("⚙️ 고급 설정 (필요할 때만) — 스키마/챕터자료 직접 교체, 챕터 번호 지정"):
    st.write("**챕터 번호**: 업로드하는 문서에 'Chapter' 헤더가 없는 단일 챕터 문서일 경우 사용할 기본값")
    default_chapter_num = st.number_input("기본 챕터 번호", min_value=1, max_value=99, value=default_chapter_num)

    override_ref = st.file_uploader("챕터/CELL 참고문서 교체 (.docx)", type=["docx"], key="ref_override")
    if override_ref is not None:
        with open("_ref_override.docx", "wb") as f:
            f.write(override_ref.read())
        chapter_ref = parse_chapter_reference("_ref_override.docx")
        candidates_by_chapter = build_candidates_for_all_chapters(chapter_ref, CELL_ID_TEMPLATE)
        st.success("챕터 참고문서를 교체했습니다.")

    override_sample = st.file_uploader("샘플/스키마 엑셀 교체 (.xlsx)", type=["xlsx"], key="sample_override")
    if override_sample is not None:
        with open("_sample_override.xlsx", "wb") as f:
            f.write(override_sample.read())
        schema_sample = load_sample("_sample_override.xlsx")
        TARGET_COLUMNS = schema_sample["columns"]
        VARYING_FIELDS = [c for c in schema_sample["varying_cols"] if c != "page_order_seq"]
        CELL_ID_TEMPLATE = next((r["cell_id"] for r in schema_sample["rows"] if r.get("cell_id")), None)
        st.session_state.lesson_meta_by_level[level_code] = dict(schema_sample["constant_cols"])
        st.success("샘플/스키마를 교체했습니다.")

    if chapter_ref:
        rows = []
        for ch, chinfo in sorted(chapter_ref.items()):
            for c, title in sorted(chinfo["cells"].items()):
                rows.append({"chapter": ch, "chapter_title": chinfo["title"], "cell": c, "cell_title": title})
        st.dataframe(pd.DataFrame(rows), width="stretch", height=200)
    else:
        st.caption("현재 챕터별 CELL 후보: " + str(candidates_by_chapter))

# ================= 사이드바: 공통 메타데이터 =================
with st.sidebar:
    st.header("⚙️ 설정")
    st.subheader("공통 메타데이터 (내장 템플릿에서 자동 채워짐, 수정 가능)")
    meta = st.session_state.lesson_meta_by_level[level_code]
    meta_keys = [c for c in TARGET_COLUMNS if c not in VARYING_FIELDS and c != "page_order_seq"]
    for k in meta_keys:
        meta[k] = st.text_input(k, value=str(meta.get(k, "")))
    if st.button("💾 이 값 저장(재사용)"):
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)
        st.success(f"{level} 설정을 저장했습니다. 다음에 {level}을 선택하면 자동으로 불러옵니다.")

# ================= STEP 2: 문제 문서 업로드 (유일한 필수 업로드) =================
st.header("2️⃣ 문제 문서 업로드")
st.caption(
    "H레벨: 정답이 빨간 글씨로 표시된 문제모음집 워드 문서를 올리세요 (여러 챕터가 한 파일에 있어도 자동 인식).\n"
    "E레벨: 정답 표시가 없는 문제 문서를 올릴 수도 있습니다 — 이 경우 아래 정답지를 함께 올리거나, "
    "검수 단계에서 Answer를 직접 채워주세요."
)
qbank_file = st.file_uploader("문제 문서 (.docx)", type=["docx"], key="qbank_upload")
if qbank_file is not None:
    with open("_qbank_tmp.docx", "wb") as f:
        f.write(qbank_file.read())
    st.session_state.raw_questions = parse_question_bank(
        "_qbank_tmp.docx", default_chapter_num=default_chapter_num
    )

if st.session_state.raw_questions:
    st.success(f"{len(st.session_state.raw_questions)}개 문항을 파싱했습니다.")
    cnt = Counter(q["chapter_num"] for q in st.session_state.raw_questions)
    st.write({f"Chapter {k}": v for k, v in sorted(cnt.items())})

    with st.expander("🔍 원본 파싱 결과 미리보기 (검수용)"):
        preview_rows = []
        for q in st.session_state.raw_questions:
            preview_rows.append({
                "chapter": q["chapter_num"],
                "q_num": q["q_num"],
                "instruction": q["instruction"] or q["shared_instruction"],
                "options": " / ".join(q["options"]),
                "answer_marks(빨간글씨)": " | ".join(q["answer_marks"]),
                "raw_lines": " | ".join(q["raw_lines"]),
            })
        st.dataframe(pd.DataFrame(preview_rows), width="stretch", height=350)
        st.caption("⚠️ 파싱은 휴리스틱입니다. H레벨인데 정답 표시가 비어있다면 원본 문서의 빨간색 서식을 확인해주세요.")

st.divider()
st.subheader("📑 정답지 엑셀 업로드 (선택)")
st.caption(
    "챕터/문항번호별로 난이도·정답·해설이 이미 정리된 정답지 엑셀이 있으면 올려주세요. "
    "일치하는 문항은 정답지의 정답/난이도를 그대로 사용합니다. "
    "형식: 학기 | 교재명 | Chapter | 문제 | 난이도 | 정답 | 해설 (열 순서 무관, 이 이름들만 있으면 인식)"
)
answer_key_file = st.file_uploader("정답지 엑셀 (.xlsx, 선택)", type=["xlsx"], key="answer_key_upload")
if answer_key_file is not None:
    with open("_answer_key_tmp.xlsx", "wb") as f:
        f.write(answer_key_file.read())
    try:
        parsed_key = load_answer_key("_answer_key_tmp.xlsx")
        n = count_entries(parsed_key)
        if n == 0:
            st.warning(
                "정답지에서 인식된 문항이 없습니다. 헤더 행에 'Chapter', '문제', '난이도', '정답' 컬럼명이 "
                "정확히 있는지 확인해주세요."
            )
        else:
            st.session_state.answer_key = parsed_key
            st.success(f"정답지에서 {n}개 문항의 정답/난이도를 읽었습니다 (챕터: {sorted(parsed_key.keys())}).")
    except Exception as e:
        st.error(f"정답지 파싱 중 오류: {e}")

if st.session_state.answer_key:
    with st.expander(f"📑 인식된 정답지 미리보기 ({count_entries(st.session_state.answer_key)}개 문항)"):
        key_rows = []
        for ch, qs in sorted(st.session_state.answer_key.items()):
            for qn, v in sorted(qs.items()):
                key_rows.append({"chapter": ch, "q_num": qn, **v})
        st.dataframe(pd.DataFrame(key_rows), width="stretch", height=250)
        if st.button("🗑️ 정답지 제거"):
            st.session_state.answer_key = {}
            st.rerun()

# ================= STEP 3: 검수 & 편집 =================
st.header("3️⃣ 검수 및 편집")
st.caption(
    "⚠️ cell_id/cell_title은 문서 안에 문항별 CELL 표시가 없어 자동 판단하지 않습니다. "
    "아래 표의 'cell_id 후보(참고용)' 컬럼에서 해당 챕터의 CELL 후보를 확인하고, cell_id/cell_title 칸에 "
    "직접 골라 입력해주세요."
)
if st.session_state.raw_questions:
    enriched = build_all(st.session_state.raw_questions, answer_key=st.session_state.answer_key)

    final_rows = []
    for e in enriched:
        row = dict(meta)
        row["page_order_seq"] = e.get("q_num")
        for col in VARYING_FIELDS:
            row[col] = e.get(col, "")
        cands = candidates_by_chapter.get(e.get("chapter_num"), [])
        row["_cell_id_candidates"] = (
            " | ".join(f"{cid} ({title})" for cid, title in cands) if cands else ""
        )
        final_rows.append(row)

    df = pd.DataFrame(final_rows)
    for col in TARGET_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    display_columns = list(TARGET_COLUMNS)
    insert_at = display_columns.index("cell_id") + 1 if "cell_id" in display_columns else 0
    display_columns.insert(insert_at, "_cell_id_candidates")
    df = df[display_columns]

    edited_df = st.data_editor(
        df, width="stretch", height=500, num_rows="dynamic",
        column_config={
            "_cell_id_candidates": st.column_config.TextColumn(
                "cell_id 후보 (참고용 - 엑셀에는 포함되지 않음)", disabled=True, width="large",
            ),
        },
    )

    st.header("4️⃣ 엑셀로 내보내기")
    export_df = edited_df[TARGET_COLUMNS]
    buf = io.BytesIO()
    export_df.to_excel(buf, index=False, engine="openpyxl")
    buf.seek(0)
    st.download_button(
        "⬇️ 엑셀 다운로드",
        data=buf,
        file_name="question_bank_export.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("먼저 2️⃣ 문제 문서를 업로드하세요.")
