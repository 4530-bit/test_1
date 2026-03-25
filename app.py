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
- 광고 파일의 **모든 시트 × 모든 열**을 자동 스캔하여 아파트 목록을 찾습니다.
- 여러 광고에 공통으로 포함된 아파트를 우선 배정합니다.
- 같은 광고명(F열)이 여러 행인 경우, 지역3이 중복되지 않도록 배정합니다.
- B파일 **O열** ← 센터명, **X열** ← 배정된 아파트명
""")

A_FILE = "타운보드_가동리스트_패키지상품__260323.xlsx"

@st.cache_data
def load_apt_map():
    if not os.path.exists(A_FILE):
        return None, f"❌ A파일({A_FILE})을 찾을 수 없습니다."
    try:
        # openpyxl로 직접 읽기 (pandas보다 안정적)
        wb = load_workbook(A_FILE, read_only=True, data_only=True)
        ws = wb.worksheets[0]  # 첫 번째 시트

        result = {}
        header_found = False
        col_apt, col_g, col_center = 3, 7, 13  # 기본 인덱스 (0-based)

        for row in ws.iter_rows(values_only=True):
            # 헤더 행 찾기
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

st.divider()
col1, col2 = st.columns(2)
with col1:
    file_b = st.file_uploader("📋 B 파일 (F열: 광고명 → O열·X열에 결과 기입)", type=["xlsx"])
with col2:
    ad_files = st.file_uploader("📁 광고 파일들 (여러 개 업로드 가능)", type=["xlsx"], accept_multiple_files=True)

if file_b and ad_files:
    if st.button("🚀 최적화 실행", type="primary"):
        with st.spinner("분석 중..."):

            ad_file_apts = {}
            for f in ad_files:
                key = ad_key(f.name)
                apts, sheet, col = find_apts_auto(f, apt_map)
                ad_file_apts[key] = apts
                if apts:
                    st.write(f"✅ **{key}**: '{sheet}' 시트 / {col+1}번 열 → {len(apts)}개 매칭")
                else:
                    st.write(f"⚠️ **{key}**: 아파트 목록 없음 (패키지 상품으로 처리)")

            apt_freq = Counter()
            for apts in ad_file_apts.values():
                for a in set(apts):
                    apt_freq[a] += 1

            file_b.seek(0)
            wb = load_workbook(file_b)
            ws = wb.active
            COL_O, COL_X = 15, 24

            used_g_per_ad = {}
            results = []

            for excel_row in range(2, ws.max_row + 1):
                cell_f = ws.cell(row=excel_row, column=6)
                ad_name = str(cell_f.value).strip() if cell_f.value else ''
                if not ad_name or ad_name in ('None', '소  재  명 ', '소재명'):
                    continue

                matched_key = fuzzy_match(ad_name, list(ad_file_apts.keys()))
                if not matched_key:
                    results.append({'행': excel_row, '광고명': ad_name,
                                    '센터명(O열)': '—', '배정 아파트(X열)': '광고파일 없음',
                                    '비고': '광고파일 미업로드'})
                    continue

                candidates = ad_file_apts.get(matched_key, [])
                if not candidates:
                    results.append({'행': excel_row, '광고명': ad_name,
                                    '센터명(O열)': '—', '배정 아파트(X열)': '—',
                                    '비고': '패키지 상품 (아파트 미지정)'})
                    continue

                if matched_key not in used_g_per_ad:
                    used_g_per_ad[matched_key] = set()

                sorted_cands = sorted(candidates, key=lambda a: apt_freq.get(a, 0), reverse=True)
                chosen = None
                for apt in sorted_cands:
                    g_val, _ = apt_map[apt]
                    if g_val not in used_g_per_ad[matched_key]:
                        chosen = apt
                        used_g_per_ad[matched_key].add(g_val)
                        break
                if not chosen:
                    chosen = sorted_cands[0]

                g_val, center = apt_map[chosen]
                ws.cell(row=excel_row, column=COL_O).value = center if center else ''
                ws.cell(row=excel_row, column=COL_X).value = chosen
                results.append({'행': excel_row, '광고명': ad_name,
                                '센터명(O열)': center if center else '(센터명 없음)',
                                '배정 아파트(X열)': chosen,
                                '비고': f"공통 {apt_freq.get(chosen,1)}개 광고"})

            st.divider()
            st.subheader("📊 배정 결과")
            df_result = pd.DataFrame(results)
            st.dataframe(df_result, use_container_width=True)

            n_ok = len(df_result[df_result['센터명(O열)'] != '—'])
            n_no = len(df_result[df_result['센터명(O열)'] == '—'])
            c1, c2 = st.columns(2)
            c1.metric("배정 완료", f"{n_ok}행")
            c2.metric("미배정", f"{n_no}행")

            out = BytesIO()
            wb.save(out)
            st.download_button(
                label="📥 결과 B파일 다운로드",
                data=out.getvalue(),
                file_name="B_광고동선최적화_결과.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.info("B 파일과 광고 파일을 업로드하면 자동으로 분석이 시작됩니다.")
