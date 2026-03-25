import streamlit as st
import pandas as pd
from io import BytesIO
from collections import Counter
from openpyxl import load_workbook
import os

st.set_page_config(page_title="광고 동선 최적화", layout="wide")
st.title("📍 광고 동선 최적화 시스템")
st.info("""
**작동 방식**
- 아파트 목록 파일(A파일)은 자동으로 불러옵니다.
- **광고게첨리스트** 파일명 → O열(센터명) + S열(아파트명) 자동 입력
- **그 외 B파일** → O열(센터명) + X열(아파트명) 자동 입력
- O열에 이미 센터명이 있는 행은 유지하고, 아파트명(S 또는 X열)만 배정합니다.
- 여러 B파일을 한번에 업로드하면 각각 처리 후 모두 다운로드할 수 있습니다.
""")

A_FILE = "타운보드_가동리스트_패키지상품__260323.xlsx"

@st.cache_data
def load_apt_map():
    if not os.path.exists(A_FILE):
        return None, f"❌ A파일({A_FILE})을 찾을 수 없습니다."
    try:
        wb = load_workbook(A_FILE, read_only=True, data_only=True)
        ws = wb.worksheets[0]
        result = {}
        header_found = False
        col_apt, col_g, col_center = 3, 7, 13

        for row in ws.iter_rows(values_only=True):
            if not header_found:
                row_vals = [str(v) if v is not None else '' for v in row]
                if '아파트명' in row_vals:
                    col_apt    = row_vals.index('아파트명')
                    col_g      = row_vals.index('지역3(법정)') if '지역3(법정)' in row_vals else 7
                    col_center = row_vals.index('센터명') if '센터명' in row_vals else 13
                    header_found = True
                continue

            aname  = str(row[col_apt]).strip()    if row[col_apt]    is not None else ''
            g_val  = str(row[col_g]).strip()      if row[col_g]      is not None else ''
            center = str(row[col_center]).strip() if row[col_center] is not None else ''

            if not aname or aname in ('None', '아파트명') or aname.startswith('합'):
                continue
            if center in ('Y', 'None'):
                center = ''

            if aname not in result:
                result[aname] = (g_val, center)
            elif not result[aname][1] and center:
                result[aname] = (g_val, center)

        wb.close()
        return result, None
    except Exception as e:
        return None, f"❌ A파일 로드 오류: {e}"


apt_map, err = load_apt_map()
if err:
    st.error(err)
    st.stop()

센터있음 = sum(1 for v in apt_map.values() if v[1])
st.success(f"✅ A파일 자동 로드 완료 → 아파트 {len(apt_map):,}개 (센터명 매핑: {센터있음:,}개)")

SKIP = {'단지명', '매체그룹명', 'nan', '', '합  계', '합계', '아파트', '구분',
        '단지수', '소재명', '광고주', 'MGID', '유형별', '선택', '아파트명', 'None'}

def is_게첨리스트(fname):
    return '광고게첨리스트' in fname

def find_header_row(ws):
    """헤더 행(광고명 열 포함) 찾기 → (header_excel_row, ad_col_index)"""
    for i, row in enumerate(ws.iter_rows(max_row=20, values_only=True), start=1):
        row_vals = [str(v).strip() if v is not None else '' for v in row]
        for j, v in enumerate(row_vals):
            if '소' in v and '재' in v and '명' in v:
                return i, j
    return 1, 5  # 기본값

def find_apts_auto(uploaded_file, apt_map):
    try:
        xl = pd.ExcelFile(uploaded_file)
    except Exception:
        return [], None, None
    best_count, best_apts, best_sheet, best_col = 0, [], None, None
    for sheet in xl.sheet_names:
        try:
            uploaded_file.seek(0)
            df = pd.read_excel(uploaded_file, sheet_name=sheet, header=None)
            for col_idx in range(min(12, df.shape[1])):
                col_vals = df.iloc[:, col_idx].dropna().astype(str).str.strip()
                matches = [v for v in col_vals
                           if v not in SKIP and not v.startswith('합') and v in apt_map]
                unique = list(dict.fromkeys(matches))
                if len(unique) > best_count:
                    best_count = len(unique)
                    best_apts = unique
                    best_sheet = sheet
                    best_col = col_idx
        except Exception:
            continue
    try:
        uploaded_file.seek(0)
    except Exception:
        pass
    return best_apts, best_sheet, best_col

def ad_key(fname):
    return fname.replace('.xlsx', '').replace('송출요청서_', '').strip()

def fuzzy_match(ad_name, keys):
    ad_clean = ad_name.strip()
    for k in keys:
        if ad_clean == k or ad_clean in k or k in ad_clean:
            return k
    return None

def process_b_file(b_file, ad_file_apts, apt_map, apt_freq, is_게첨):
    """B파일 1개 처리 → (결과 wb, results 리스트)"""
    b_file.seek(0)
    wb = load_workbook(b_file)
    ws = wb.active

    # 헤더 행 & 광고명 열 찾기
    header_row, ad_col_idx = find_header_row(ws)
    ad_col = ad_col_idx + 1  # openpyxl은 1-indexed

    # 열 설정: 게첨리스트 vs 일반 B파일
    if is_게첨:
        COL_CENTER = 15   # O열 = 센터명
        COL_APT    = 19   # S열 = 아파트명 (새로 입력)
    else:
        COL_CENTER = 15   # O열
        COL_APT    = 24   # X열

    used_g_per_ad = {}
    results = []

    for excel_row in range(header_row + 1, ws.max_row + 1):
        cell_f = ws.cell(row=excel_row, column=ad_col)
        ad_name = str(cell_f.value).strip() if cell_f.value else ''
        skip_names = {'None', '소  재  명 ', '소재명', ''}
        if not ad_name or ad_name in skip_names:
            continue

        # O열에 이미 센터명이 있으면 센터는 유지, 아파트만 배정
        existing_center = ws.cell(row=excel_row, column=COL_CENTER).value
        has_center = existing_center and str(existing_center).strip() not in ('', 'None', 'NaN')

        matched_key = fuzzy_match(ad_name, list(ad_file_apts.keys()))
        if not matched_key:
            results.append({'행': excel_row, '광고명': ad_name,
                            '센터명': '—', '아파트명': '광고파일 없음', '비고': '광고파일 미업로드'})
            continue

        candidates = ad_file_apts.get(matched_key, [])
        if not candidates:
            results.append({'행': excel_row, '광고명': ad_name,
                            '센터명': str(existing_center) if has_center else '—',
                            '아파트명': '—', '비고': '패키지 상품 (아파트 미지정)'})
            continue

        if matched_key not in used_g_per_ad:
            used_g_per_ad[matched_key] = set()

        # 이미 센터 있으면 해당 센터에 맞는 아파트 우선 선택
        sorted_cands = sorted(candidates, key=lambda a: apt_freq.get(a, 0), reverse=True)

        chosen = None
        if has_center:
            center_val = str(existing_center).strip()
            # 해당 센터의 아파트 중 지역3 중복 없는 것 선택
            for apt in sorted_cands:
                g_val, apt_center = apt_map[apt]
                if apt_center == center_val and g_val not in used_g_per_ad[matched_key]:
                    chosen = apt
                    used_g_per_ad[matched_key].add(g_val)
                    break
            # 해당 센터에 없으면 그냥 지역3 중복 없는 것
            if not chosen:
                for apt in sorted_cands:
                    g_val, _ = apt_map[apt]
                    if g_val not in used_g_per_ad[matched_key]:
                        chosen = apt
                        used_g_per_ad[matched_key].add(g_val)
                        break
        else:
            for apt in sorted_cands:
                g_val, _ = apt_map[apt]
                if g_val not in used_g_per_ad[matched_key]:
                    chosen = apt
                    used_g_per_ad[matched_key].add(g_val)
                    break

        if not chosen and sorted_cands:
            chosen = sorted_cands[0]

        if chosen:
            g_val, center = apt_map[chosen]
            final_center = str(existing_center).strip() if has_center else (center if center else '')
            if not has_center:
                ws.cell(row=excel_row, column=COL_CENTER).value = final_center
            ws.cell(row=excel_row, column=COL_APT).value = chosen
            results.append({'행': excel_row, '광고명': ad_name,
                            '센터명': final_center, '아파트명': chosen,
                            '비고': f"공통 {apt_freq.get(chosen,1)}개 광고"})

    return wb, results


# ── UI ──
st.divider()
col1, col2 = st.columns(2)
with col1:
    b_files = st.file_uploader(
        "📋 B 파일 (기존 B파일 또는 광고게첨리스트 파일, 여러 개 가능)",
        type=["xlsx"], accept_multiple_files=True
    )
with col2:
    ad_files = st.file_uploader(
        "📁 광고 파일들 (여러 개 업로드 가능)",
        type=["xlsx"], accept_multiple_files=True
    )

if b_files and ad_files:
    if st.button("🚀 최적화 실행", type="primary"):
        with st.spinner("분석 중..."):

            # 광고 파일 아파트 목록 추출
            ad_file_apts = {}
            for f in ad_files:
                key = ad_key(f.name)
                apts, sheet, col = find_apts_auto(f, apt_map)
                ad_file_apts[key] = apts
                if apts:
                    st.write(f"✅ **{key}**: '{sheet}' 시트 / {col+1}번 열 → {len(apts)}개 매칭")
                else:
                    st.write(f"⚠️ **{key}**: 아파트 목록 없음 (패키지 상품)")

            apt_freq = Counter()
            for apts in ad_file_apts.values():
                for a in set(apts):
                    apt_freq[a] += 1

            st.divider()

            # B파일 각각 처리
            for b_file in b_files:
                게첨 = is_게첨리스트(b_file.name)
                파일타입 = "광고게첨리스트 (O열+S열)" if 게첨 else "일반 B파일 (O열+X열)"
                st.subheader(f"📄 {b_file.name}")
                st.caption(f"파일 유형: {파일타입}")

                wb, results = process_b_file(b_file, ad_file_apts, apt_map, apt_freq, 게첨)

                df_result = pd.DataFrame(results)
                if not df_result.empty:
                    st.dataframe(df_result, use_container_width=True)
                    n_ok = len(df_result[df_result['아파트명'] != '—'])
                    n_no = len(df_result) - n_ok
                    c1, c2 = st.columns(2)
                    c1.metric("배정 완료", f"{n_ok}행")
                    c2.metric("미배정", f"{n_no}행")

                out = BytesIO()
                wb.save(out)
                결과파일명 = b_file.name.replace('.xlsx', '_결과.xlsx')
                st.download_button(
                    label=f"📥 {결과파일명} 다운로드",
                    data=out.getvalue(),
                    file_name=결과파일명,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=b_file.name
                )
                st.divider()
else:
    st.info("B 파일과 광고 파일을 업로드하면 자동으로 분석이 시작됩니다.")
