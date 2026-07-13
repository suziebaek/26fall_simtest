import io
import json
import os

import pandas as pd
import streamlit as st
from anthropic import Anthropic

from doc_parser import parse_chapter_reference, parse_question_bank
from ai_enrich import enrich_all, MODEL_ID

st.set_page_config(page_title="문제은행 자동 변환기", layout="wide")

TARGET_COLUMNS = [
    "service_code", "track_code", "top_cors_id", "component_code", "book_code",
    "act_code", "level_code", "lesson_order_seq", "lesson_title", "page_order_seq",
    "cell_id", "cell_title", "q_type", "tp_type", "difficulty", "text_passage",
    "question", "text_prompt", "image_promt", "text_example_1", "text_example_2",
    "text_example_3", "text_example_4", "text_example_5", "Answer", "explanation",
]

CONFIG_PATH = "lesson_meta.json"

DEFAULT_META = {
    "service_code": "SVC177",
    "track_code": "ST_TRK01",
    "top_cors_id": "1880",
    "component_code": "SVC177",
    "book_code": "SVC177",
    "act_code": "ST_ACT01",
    "level_code": "TO_G_H_AU",
    "lesson_order_seq": "1",
    "lesson_title": "Chapter 01/",
}

# ---------- 세션 상태 초기화 ----------
for key, default in [
    ("chapter_ref", None),
    ("raw_questions", None),
    ("enriched", None),
    ("lesson_meta", None),
]:
    if key not in st.session_state:
        st.session_state[key] = default

if st.session_state.lesson_meta is None:
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, encoding="utf-8") as f:
            st.session_state.lesson_meta = json.load(f)
    else:
        st.session_state.lesson_meta = dict(DEFAULT_META)


# ================= 사이드바: 공통 메타데이터 & API 설정 =================
with st.sidebar:
    st.header("⚙️ 설정")

    st.subheader("공통 메타데이터 (교재/챕터 단위 고정값)")
    meta = st.session_state.lesson_meta
    for k in DEFAULT_META:
        meta[k] = st.text_input(k, value=str(meta.get(k, "")))
    if st.button("💾 이 값 저장(재사용)"):
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)
        st.success("저장했습니다. 다음에 앱을 열면 자동으로 불러옵니다.")

    st.divider()
    st.subheader("Claude API")
    try:
        default_key = st.secrets.get("ANTHROPIC_API_KEY", "")
    except Exception:
        default_key = ""
    api_key = st.text_input(
        "Anthropic API Key",
        value=default_key,
        type="password",
        help="st.secrets['ANTHROPIC_API_KEY']로도 설정 가능합니다 (.streamlit/secrets.toml).",
    )
    model = st.text_input("모델", value=MODEL_ID)


st.title("📘 워드 문제 → 엑셀 문제은행 자동 변환기")
st.caption("문제모음집(워드, 빨간색=정답) + 챕터/셀 참고문서를 업로드하면 AI가 난이도/해설/문법개념코드를 채워 타겟 엑셀 스키마로 변환합니다.")

# ================= STEP 1: 챕터/셀 참고문서 =================
st.header("1️⃣ 챕터 / CELL 참고문서 업로드")
ref_file = st.file_uploader("예: FG_H_챕터.docx (CHAPTER / CELL 목차)", type=["docx"], key="ref_upload")
if ref_file is not None:
    with open("_ref_tmp.docx", "wb") as f:
        f.write(ref_file.read())
    st.session_state.chapter_ref = parse_chapter_reference("_ref_tmp.docx")

if st.session_state.chapter_ref:
    rows = []
    for ch, info in sorted(st.session_state.chapter_ref.items()):
        for c, title in sorted(info["cells"].items()):
            rows.append({"chapter": ch, "chapter_title": info["title"], "cell": c, "cell_title": title})
    st.dataframe(pd.DataFrame(rows), use_container_width=True, height=200)
else:
    st.info("아직 업로드되지 않았습니다. 이 문서 없이도 진행할 수 있지만 cell_id/cell_title은 AI가 임의로 판단합니다.")

# ================= STEP 2: 문제모음집 =================
st.header("2️⃣ 문제모음집(워드, 빨간색=정답) 업로드")
qbank_file = st.file_uploader("예: FG_H_Sim_Test.docx", type=["docx"], key="qbank_upload")
if qbank_file is not None:
    with open("_qbank_tmp.docx", "wb") as f:
        f.write(qbank_file.read())
    st.session_state.raw_questions = parse_question_bank("_qbank_tmp.docx")
    st.session_state.enriched = None  # 새 업로드시 이전 결과 초기화

if st.session_state.raw_questions:
    st.success(f"{len(st.session_state.raw_questions)}개 문항을 파싱했습니다. (챕터별 개수는 아래에서 확인)")
    from collections import Counter
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
        st.dataframe(pd.DataFrame(preview_rows), use_container_width=True, height=350)
        st.caption("⚠️ 파싱은 휴리스틱입니다. 정답 표시가 비어있거나 이상하면 원본 문서 서식(빨간색 지정 여부)을 확인해주세요.")

# ================= STEP 3: 해설/난이도 문서 (선택) =================
st.header("3️⃣ (선택) 해설/난이도 문서 업로드")
st.caption("있으면 그 문서 값을 우선 사용하고, 없으면 AI가 자동 생성합니다. 현재는 준비되지 않았으므로 AI 자동 생성으로 진행합니다.")
expl_file = st.file_uploader("해설/난이도 문서 (준비되는 대로 업로드)", type=["docx", "xlsx"], key="expl_upload")
if expl_file is not None:
    st.warning("해설/난이도 문서 자동 매칭 파서는 문서 형식을 받는 대로 추가 구현이 필요합니다. 지금은 참고용으로만 저장됩니다.")

# ================= STEP 4: AI 자동 채우기 =================
st.header("4️⃣ AI로 난이도 / 해설 / cell_id / Answer 포맷 자동 채우기")
if st.session_state.raw_questions:
    if st.button("🤖 AI 자동 채우기 실행", type="primary"):
        if not api_key:
            st.error("사이드바에 Anthropic API Key를 입력해주세요.")
        else:
            client = Anthropic(api_key=api_key)
            progress = st.progress(0.0, text="시작...")

            def cb(done, total):
                progress.progress(done / total, text=f"{done}/{total} 문항 처리 중...")

            try:
                results = enrich_all(
                    client,
                    st.session_state.raw_questions,
                    st.session_state.chapter_ref or {},
                    model=model,
                    progress_cb=cb,
                )
                st.session_state.enriched = results
                st.success("완료되었습니다. 아래에서 검수/수정 후 내보내세요.")
            except Exception as e:
                st.error(f"AI 처리 중 오류: {e}")
else:
    st.info("먼저 2️⃣ 문제모음집을 업로드하세요.")

# ================= STEP 5: 검수 & 편집 =================
st.header("5️⃣ 검수 및 편집")
if st.session_state.enriched:
    final_rows = []
    for e in st.session_state.enriched:
        row = dict(meta)
        row["page_order_seq"] = e.get("q_num")
        for col in [
            "cell_id", "cell_title", "q_type", "tp_type", "difficulty",
            "text_passage", "question", "text_prompt", "image_promt",
            "text_example_1", "text_example_2", "text_example_3",
            "text_example_4", "text_example_5", "Answer", "explanation",
        ]:
            row[col] = e.get(col, "")
        final_rows.append(row)

    df = pd.DataFrame(final_rows)[TARGET_COLUMNS]
    edited_df = st.data_editor(df, use_container_width=True, height=500, num_rows="dynamic")

    st.header("6️⃣ 엑셀로 내보내기")
    buf = io.BytesIO()
    edited_df.to_excel(buf, index=False, engine="openpyxl")
    buf.seek(0)
    st.download_button(
        "⬇️ 엑셀 다운로드",
        data=buf,
        file_name="question_bank_export.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("4️⃣ 단계를 먼저 실행하세요.")
