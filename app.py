
import streamlit as st
import pandas as pd
from collections import defaultdict, Counter
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from pathlib import Path
import os

st.set_page_config(page_title="엑셀 행 재정렬 안전 비교 (전체열 + 색상)", layout="wide")
st.title("📘 엑셀 행 재정렬 안전 비교 (전체열 + 색상)")
st.caption("기준 파일과 비교 파일을 선택하면, 행 순서가 달라도 전체 열에서 **값 변경**과 **배경색(채우기) 변경**을 잡아냅니다.")

# ----------------------- 색상/채우기 라벨링 -----------------------
def _fill_is_nonempty(fill) -> bool:
    if fill is None:
        return False
    pt = getattr(fill, "patternType", None)
    if not pt or str(pt).lower() == "none":
        return False
    fg = getattr(fill, "fgColor", None)
    if fg is None:
        return True
    if getattr(fg, "rgb", None) or getattr(fg, "indexed", None) is not None or getattr(fg, "theme", None) is not None:
        return True
    return True

def _color_hex_from_fg(fg) -> str | None:
    if fg is None:
        return None
    rgb = getattr(fg, "rgb", None)
    if isinstance(rgb, str):
        s = rgb.replace("#", "").upper()
        if len(s) == 8:
            s = s[2:]
        if len(s) == 6:
            return "#" + s
    idx = getattr(fg, "indexed", None)
    if idx is not None:
        mapping = {1:"#000000", 2:"#FFFFFF", 6:"#FFFF00"}
        return mapping.get(idx, f"indexed-{idx}")
    return None

def fill_to_label(fill) -> str:
    if fill is None:
        return "No Fill"
    pt = getattr(fill, "patternType", None)
    if not pt or str(pt).lower() == "none":
        return "No Fill"
    fg = getattr(fill, "fgColor", None)
    hx = _color_hex_from_fg(fg)
    if hx is None:
        return "Fill"
    friendly = {
        "#FFFFFF":"White",
        "#000000":"Black",
        # Yellow shades
        "#FFFF00":"Yellow",
        "#FFF2CC":"Light Yellow",
        "#FFD966":"Gold",
        "#FFEB9C":"Light Yellow 2",
        "#FFFF99":"Light Yellow (Alt)",
        "#FFFFCC":"Pale Yellow",
        # Red shades
        "#FF0000":"Red",
        "#FFC7CE":"Light Red",
        "#FFCCCC":"Pale Red",
        "#FF6666":"Light Red 2",
        # Green shades
        "#00FF00":"Green",
        "#00B050":"Dark Green",
        "#92D050":"Light Green",
        "#C6E0B4":"Pale Green",
        "#E2EFDA":"Very Light Green",
        # Blue shades
        "#0000FF":"Blue",
        "#00B0F0":"Light Blue",
        "#BDD7EE":"Pale Blue",
        "#DDEBF7":"Very Light Blue",
        # Orange shades
        "#FFA500":"Orange",
        "#F8CBAD":"Light Orange",
        "#FFC000":"Dark Orange",
        # Purple shades
        "#7030A0":"Purple",
        "#B4A7D6":"Light Purple",
        # Gray shades
        "#D9D9D9":"Light Gray",
        "#BFBFBF":"Gray",
        "#808080":"Dark Gray",
    }.get(hx)
    return friendly or hx

# ----------------------- 범위(행/열) 계산 -----------------------
def compute_used_bounds(ws):
    max_r, max_c = 0, 0
    for r in range(1, ws.max_row + 1):
        row_has_any = False
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            if (cell.value not in (None, "")) or _fill_is_nonempty(cell.fill):
                row_has_any = True
                if c > max_c:
                    max_c = c
        if row_has_any:
            max_r = r
    if max_r == 0:
        max_r = ws.max_row
    if max_c == 0:
        max_c = ws.max_column
    return max_r, max_c

# ----------------------- 정규화 -----------------------
def normalize_value(v, trim_spaces=True, case_sensitive=True):
    if isinstance(v, str):
        s = v.strip() if trim_spaces else v
        return s if case_sensitive else s.lower()
    return v

# ----------------------- 시트 읽기 -----------------------
def read_sheet_values_and_fills(file, sheet_name=None, trim_spaces=True, case_sensitive=True):
    wb = load_workbook(file, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.active
    max_r, max_c = compute_used_bounds(ws)
    cols = [get_column_letter(c) for c in range(1, max_c + 1)]

    rows = []
    fills = {}
    for r in range(1, max_r + 1):
        orig = {}
        norm = {}
        empty_all = True
        for c in range(1, max_c + 1):
            cell = ws.cell(row=r, column=c)
            v = cell.value
            col = get_column_letter(c)
            orig[col] = v
            norm[col] = normalize_value(v, trim_spaces, case_sensitive)
            fills[(r, c)] = fill_to_label(cell.fill)
            if (v not in (None, "")) or _fill_is_nonempty(cell.fill):
                empty_all = False
        if not empty_all:
            rows.append({"_row": r, "orig": orig, "norm": norm})
    return rows, fills, cols

# ----------------------- 페어링 -----------------------
def row_tuple(norm_row, columns):
    return tuple(norm_row.get(col) for col in columns)

def best_pairing(new_rows, old_rows, columns):
    candidates = []
    for i, o in enumerate(old_rows):
        for j, n in enumerate(new_rows):
            eq = sum(1 for col in columns if o["norm"].get(col) == n["norm"].get(col))
            if eq > 0:
                candidates.append((eq, i, j))
    candidates.sort(reverse=True)
    used_old, used_new = set(), set()
    pairs = []
    for eq, i, j in candidates:
        if i in used_old or j in used_new:
            continue
        pairs.append((i, j, eq))
        used_old.add(i); used_new.add(j)
    leftover_old = [i for i in range(len(old_rows)) if i not in used_old]
    leftover_new = [j for j in range(len(new_rows)) if j not in used_new]
    return pairs, leftover_old, leftover_new

# ----------------------- 변경 레코드 -----------------------
def build_diff_record(old_row, new_row, old_fills, new_fills, columns):
    changes = []
    for idx, col in enumerate(columns, start=1):
        r_old = old_row["_row"]
        r_new = new_row["_row"]
        ov = old_row["orig"].get(col)
        nv = new_row["orig"].get(col)
        value_changed = old_row["norm"].get(col) != new_row["norm"].get(col)

        ofill = old_fills.get((r_old, idx), "No Fill")
        nfill = new_fills.get((r_new, idx), "No Fill")
        fill_changed = ofill != nfill

        if value_changed or fill_changed:
            if value_changed and fill_changed:
                changes.append(f"{col}열 값 '{ov}'→'{nv}', 색 '{ofill}'→'{nfill}'")
            elif value_changed:
                changes.append(f"{col}열 값 '{ov}'→'{nv}'")
            elif fill_changed:
                changes.append(f"{col}열 색 '{ofill}'→'{nfill}'")
    msg = "; ".join(changes) if changes else "변경 없음"
    return {
        "기준행": old_row["_row"],
        "비교행": new_row["_row"],
        "변경요약": msg
    }

# ----------------------- 로컬 폴더에서 파일 가져오기 -----------------------
def get_excel_files_in_folder(folder_path):
    """폴더 내의 모든 엑셀 파일 목록 반환"""
    try:
        if not folder_path or not os.path.exists(folder_path):
            return []
        path = Path(folder_path)
        excel_files = list(path.glob("*.xlsx")) + list(path.glob("*.xls"))
        # 임시 파일 제외
        excel_files = [f for f in excel_files if not f.name.startswith("~$")]
        return sorted([f.name for f in excel_files])
    except Exception as e:
        st.error(f"폴더 읽기 오류: {e}")
        return []

# ----------------------- UI -----------------------
with st.expander("⚙️ 설정", expanded=True):
    col_opt1, col_opt2 = st.columns(2)
    with col_opt1:
        trim_spaces = st.checkbox("앞뒤 공백 무시", value=True)
        case_sensitive = st.checkbox("대소문자 구분", value=True)
    with col_opt2:
        # 파일 입력 방식 선택
        input_mode = st.radio("파일 입력 방식", ["로컬 폴더", "파일 업로드"], horizontal=True)

st.subheader("1️⃣ 기준(이전) 파일 선택")

if input_mode == "로컬 폴더":
    # 현재 작업 디렉토리를 기본값으로 사용
    default_folder = os.getcwd()
    folder_path = st.text_input("📁 폴더 경로", value=default_folder, help="엑셀 파일이 있는 폴더 경로를 입력하세요")
    
    if folder_path and os.path.exists(folder_path):
        excel_files = get_excel_files_in_folder(folder_path)
        
        if excel_files:
            c1, c2 = st.columns(2)
            with c1:
                selected_old_file = st.selectbox("기준 파일 선택", options=excel_files, key="old_file_select")
                file_old = os.path.join(folder_path, selected_old_file) if selected_old_file else None
            with c2:
                sheet_old = None
                if file_old:
                    try:
                        wb = load_workbook(file_old, read_only=True, data_only=True)
                        sheet_old = st.selectbox("시트 선택(기준)", options=wb.sheetnames, index=0, key="old_sheet")
                        wb.close()
                    except Exception as e:
                        st.error(f"기준 파일 시트 읽기 실패: {e}")
        else:
            st.warning("⚠️ 선택한 폴더에 엑셀 파일이 없습니다.")
            file_old = None
            sheet_old = None
    else:
        st.warning("⚠️ 유효한 폴더 경로를 입력하세요.")
        file_old = None
        sheet_old = None
else:
    # 파일 업로드 방식
    c1, c2 = st.columns(2)
    with c1:
        file_old = st.file_uploader("기준 엑셀 파일", type=["xlsx"], key="old_allcols")
    with c2:
        sheet_old = None
        if file_old:
            try:
                wb = load_workbook(file_old, read_only=True, data_only=True)
                sheet_old = st.selectbox("시트 선택(기준)", options=wb.sheetnames, index=0)
                wb.close()
            except Exception as e:
                st.error(f"기준 파일 시트 읽기 실패: {e}")

if st.button("✅ 기준 데이터 저장", type="primary", disabled=not (file_old and sheet_old)):
    try:
        with st.spinner("기준 파일을 읽는 중..."):
            old_rows, old_fills, cols = read_sheet_values_and_fills(file_old, sheet_old, trim_spaces, case_sensitive)
            
            if not old_rows:
                st.error("❌ 기준 파일에 데이터가 없습니다.")
            else:
                st.session_state["old_rows"] = old_rows
                st.session_state["old_fills"] = old_fills
                st.session_state["columns"] = cols
                st.session_state["trim_spaces"] = trim_spaces
                st.session_state["case_sensitive"] = case_sensitive

                multiset = Counter([row_tuple(r["norm"], cols) for r in old_rows])
                mapping = defaultdict(list)
                for idx, r in enumerate(old_rows):
                    mapping[row_tuple(r["norm"], cols)].append(idx)

                st.session_state["old_rows_norm_multiset"] = multiset
                st.session_state["old_rows_by_tuple_indices"] = mapping
                st.success(f"✅ 기준 데이터 저장 완료: {len(old_rows)} 행, 사용 열: {len(cols)}개 ({cols[0]}~{cols[-1]})")
    except Exception as e:
        st.error(f"❌ 기준 파일 처리 중 오류 발생")
        st.exception(e)

st.subheader("2️⃣ 비교(이후) 파일 선택")

if input_mode == "로컬 폴더":
    # 같은 폴더에서 비교 파일 선택
    if folder_path and os.path.exists(folder_path):
        excel_files = get_excel_files_in_folder(folder_path)
        
        if excel_files:
            c3, c4 = st.columns(2)
            with c3:
                selected_new_file = st.selectbox("비교 파일 선택", options=excel_files, key="new_file_select")
                file_new = os.path.join(folder_path, selected_new_file) if selected_new_file else None
            with c4:
                sheet_new = None
                if file_new:
                    try:
                        wb2 = load_workbook(file_new, read_only=True, data_only=True)
                        sheet_new = st.selectbox("시트 선택(비교)", options=wb2.sheetnames, index=0, key="new_sheet")
                        wb2.close()
                    except Exception as e:
                        st.error(f"비교 파일 시트 읽기 실패: {e}")
        else:
            file_new = None
            sheet_new = None
    else:
        file_new = None
        sheet_new = None
else:
    # 파일 업로드 방식
    c3, c4 = st.columns(2)
    with c3:
        file_new = st.file_uploader("비교 엑셀 파일", type=["xlsx"], key="new_allcols")
    with c4:
        sheet_new = None
        if file_new:
            try:
                wb2 = load_workbook(file_new, read_only=True, data_only=True)
                sheet_new = st.selectbox("시트 선택(비교)", options=wb2.sheetnames, index=0)
                wb2.close()
            except Exception as e:
                st.error(f"비교 파일 시트 읽기 실패: {e}")

if st.button("🔍 변경 사항 분석 실행", type="primary",
             disabled=not (file_new and sheet_new and ("old_rows" in st.session_state))):
    try:
        # 저장된 설정값 사용
        old_rows = st.session_state["old_rows"]
        old_fills = st.session_state["old_fills"]
        columns_old = st.session_state["columns"]
        old_multiset = st.session_state["old_rows_norm_multiset"]
        old_tuple_to_indices = st.session_state["old_rows_by_tuple_indices"]
        saved_trim_spaces = st.session_state.get("trim_spaces", trim_spaces)
        saved_case_sensitive = st.session_state.get("case_sensitive", case_sensitive)

        # 진행 상황 표시
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        status_text.text("📖 비교 파일을 읽는 중...")
        progress_bar.progress(10)
        
        new_rows, new_fills, cols_new = read_sheet_values_and_fills(
            file_new, sheet_new, saved_trim_spaces, saved_case_sensitive
        )
        
        if not new_rows:
            st.error("❌ 비교 파일에 데이터가 없습니다.")
            progress_bar.empty()
            status_text.empty()
        else:
            progress_bar.progress(20)
            
            # 열 범위: 기준/비교 중 더 넓은 범위를 사용 (기존 columns_old는 유지)
            all_columns = list(set(columns_old + cols_new))
            all_columns.sort(key=lambda x: (len(x), x))  # A, B, ... Z, AA, AB ...
            columns = all_columns

            status_text.text("🔄 동일한 행 매칭 중...")
            progress_bar.progress(30)
            
            remaining_old_indices = set(range(len(old_rows)))
            remaining_new_indices = set(range(len(new_rows)))

            exact_pairs = []
            temp_multiset = old_multiset.copy()
            temp_tuple_to_indices = {k: v.copy() for k, v in old_tuple_to_indices.items()}

            for j, nr in enumerate(new_rows):
                t = row_tuple(nr["norm"], columns)
                if temp_multiset.get(t, 0) > 0:
                    i = temp_tuple_to_indices[t].pop(0)
                    temp_multiset[t] -= 1
                    exact_pairs.append((i, j))
                    remaining_old_indices.discard(i)
                    remaining_new_indices.discard(j)

            progress_bar.progress(50)
            status_text.text("🔍 변경된 행 매칭 중...")
            
            old_left = [old_rows[i] for i in sorted(remaining_old_indices)]
            new_left = [new_rows[j] for j in sorted(remaining_new_indices)]
            pairs, leftover_old_idx, leftover_new_idx = best_pairing(new_left, old_left, columns)

            progress_bar.progress(60)
            status_text.text("📊 변경 내역 생성 중...")
            
            best_pairs = []
            sorted_old_left = sorted(remaining_old_indices)
            sorted_new_left = sorted(remaining_new_indices)
            for eq, i, j in sorted([(p[2], p[0], p[1]) for p in pairs], reverse=True):
                old_idx_global = sorted_old_left[i]
                new_idx_global = sorted_new_left[j]
                best_pairs.append((old_idx_global, new_idx_global, eq))

            unchanged_records = [{
                "기준행": old_rows[i]["_row"],
                "비교행": new_rows[j]["_row"],
                "상태": "동일(재정렬만)"
            } for i, j in exact_pairs]

            progress_bar.progress(70)
            
            changes_records = []
            for i, j, eq in best_pairs:
                rec = build_diff_record(old_rows[i], new_rows[j], old_fills, new_fills, columns)
                rec["일치열수"] = eq
                rec["상태"] = "변경"
                changes_records.append(rec)

            progress_bar.progress(80)
            
            used_old = set([i for i, _, _ in best_pairs] + [i for i, _ in exact_pairs])
            used_new = set([j for _, j, _ in best_pairs] + [j for _, j in exact_pairs])

            removed_records = [{"기준행": old_rows[i]["_row"], "상태": "제거됨"} for i in range(len(old_rows)) if i not in used_old]
            added_records = [{"비교행": new_rows[j]["_row"], "상태": "추가됨"} for j in range(len(new_rows)) if j not in used_new]

            progress_bar.progress(90)
            status_text.text("✨ 결과 정리 중...")
            
            df_unchanged = pd.DataFrame(unchanged_records)
            df_changes = pd.DataFrame(changes_records, columns=["기준행","비교행","일치열수","변경요약","상태"])
            df_removed = pd.DataFrame(removed_records)
            df_added = pd.DataFrame(added_records)
            
            # 세션에 저장
            st.session_state["df_unchanged"] = df_unchanged
            st.session_state["df_changes"] = df_changes
            st.session_state["df_removed"] = df_removed
            st.session_state["df_added"] = df_added
            
            progress_bar.progress(100)
            status_text.text("✅ 분석 완료!")
            
            st.success(f"✅ 분석 완료: 동일(재정렬만) {len(df_unchanged)}건, 변경 {len(df_changes)}건, 제거 {len(df_removed)}건, 추가 {len(df_added)}건")
            
            progress_bar.empty()
            status_text.empty()
    
    except Exception as e:
        if 'progress_bar' in locals():
            progress_bar.empty()
        if 'status_text' in locals():
            status_text.empty()
        st.error("❌ 분석 중 오류가 발생했습니다.")
        st.exception(e)

# ----------------------- 결과 표시 -----------------------
if "df_unchanged" in st.session_state:
    st.divider()
    st.subheader("📊 분석 결과")
    
    df_unchanged = st.session_state["df_unchanged"]
    df_changes = st.session_state["df_changes"]
    df_removed = st.session_state["df_removed"]
    df_added = st.session_state["df_added"]
    
    # 필터링 옵션
    with st.expander("🔍 결과 필터링", expanded=False):
        show_unchanged = st.checkbox("동일(재정렬만) 표시", value=True)
        show_changes = st.checkbox("변경 사항 표시", value=True)
        show_removed = st.checkbox("제거된 행 표시", value=True)
        show_added = st.checkbox("추가된 행 표시", value=True)
        
        if show_changes and not df_changes.empty:
            search_text = st.text_input("🔎 변경 내용 검색", placeholder="검색어를 입력하세요 (변경요약에서 검색)")
    
    # 동일(재정렬만)
    if show_unchanged:
        st.write("### ✅ 동일(재정렬만)")
        if not df_unchanged.empty:
            st.dataframe(df_unchanged, use_container_width=True, hide_index=True)
        else:
            st.info("동일한 행이 없습니다.")
    
    # 변경
    if show_changes:
        st.write("### 🔄 변경 (값/색상)")
        if not df_changes.empty:
            df_to_show = df_changes.copy()
            if 'search_text' in locals() and search_text:
                df_to_show = df_to_show[df_to_show["변경요약"].str.contains(search_text, case=False, na=False)]
                st.caption(f"검색 결과: {len(df_to_show)}건")
            st.dataframe(df_to_show, use_container_width=True, hide_index=True)
        else:
            st.info("변경된 행이 없습니다.")
    
    # 제거됨
    if show_removed:
        st.write("### ❌ 제거됨 (기준에는 있었으나 비교에는 없음)")
        if not df_removed.empty:
            st.dataframe(df_removed, use_container_width=True, hide_index=True)
        else:
            st.info("제거된 행이 없습니다.")
    
    # 추가됨
    if show_added:
        st.write("### ➕ 추가됨 (비교에는 있으나 기준에는 없음)")
        if not df_added.empty:
            st.dataframe(df_added, use_container_width=True, hide_index=True)
        else:
            st.info("추가된 행이 없습니다.")

    # 다운로드 버튼
    st.divider()
    st.subheader("💾 결과 다운로드")
    
    from io import BytesIO
    def to_xlsx(dfs, names):
        bio = BytesIO()
        with pd.ExcelWriter(bio, engine="openpyxl") as wr:
            for df, name in zip(dfs, names):
                if not df.empty:
                    df.to_excel(wr, index=False, sheet_name=name)
                else:
                    pd.DataFrame().to_excel(wr, index=False, sheet_name=name)
        return bio.getvalue()
    
    col_dl1, col_dl2 = st.columns(2)
    
    with col_dl1:
        # 전체 결과 다운로드
        st.download_button(
            "📥 전체 결과 다운로드",
            data=to_xlsx([df_unchanged, df_changes, df_removed, df_added],
                         ["동일", "변경", "제거", "추가"]),
            file_name="excel_compare_all_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    with col_dl2:
        # 변경/추가된 행만 다운로드
        changes_and_additions = []
        names_modified = []
        
        if not df_changes.empty:
            changes_and_additions.append(df_changes)
            names_modified.append("변경")
        if not df_added.empty:
            changes_and_additions.append(df_added)
            names_modified.append("추가")
        if not df_removed.empty:
            changes_and_additions.append(df_removed)
            names_modified.append("제거")
        
        if changes_and_additions:
            st.download_button(
                "⭐ 변경/추가/제거만 다운로드",
                data=to_xlsx(changes_and_additions, names_modified),
                file_name="excel_compare_changes_only.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary"
            )
        else:
            st.info("변경/추가/제거된 항목이 없습니다.")

st.divider()
st.info("💡 **사용 방법**: 기준 파일을 먼저 저장한 후, 비교 파일을 선택하여 분석을 실행하세요. 행 순서가 달라도 정확히 매칭하며, 모든 사용된 열(값/채우기 존재)을 자동 인식하여 비교합니다.")
