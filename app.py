import streamlit as st
import pandas as pd
from io import BytesIO
from collections import Counter
from openpyxl import load_workbook
import os

st.set_page_config(page_title="광고 동선 최적화", layout="wide")
st.title("📍 광고 동선 최적화 시스템")

# A파일 경로 설정 (사용자 환경에 맞게 파일명을 확인하세요)
A_FILE = "타운보드_가동리스트_패키지상품__260323.xlsx"

@st.cache_data
def load_apt_map():
    if not os.path.exists(A_FILE):
        return None, f"❌ A파일({A_FILE})을 찾을 수 없습니다. 파일이 app.py와 같은 폴더에 있는지 확인하세요."
    try:
        # 데이터만 읽어오기 위해 data_only=True 설정
        wb = load_workbook(A_FILE, read_only=True, data_only=True)
        ws = wb.worksheets[0]
        result = {}
        header_found = False
        col_apt, col_g, col_center = 3, 7, 13 # 기본 열 인덱스

        for row in ws.iter_rows(values_only=True):
            row_vals = [str(v).strip() if v is not None else '' for v in row]
            if not header_found:
                if '아파트명' in row_vals:
                    col_apt = row_vals.index('아파트명')
                    col_g = row_vals.index('지역3(법정)') if '지역3(법정)' in row_vals else 7
                    col_center = row_vals.index('센터명') if '센터명' in row_vals else 13
                    header_found = True
                continue

            aname = str(row[col_apt]).strip() if row[col_apt] else ''
            if not aname or aname in ('None', '아파트명') or aname.startswith('합'):
                continue
            
            g_val = str(row[col_g]).strip() if row[col_g] else ''
            center = str(row[col_center]).strip() if row[col_center] else ''
            
            # 아파트명을 키로 지역과 센터 정보 저장
            result[aname] = (g_val, center)
        return result, None
    except Exception as e:
        return None, f"❌ A파일 로드 중 오류 발생: {e}"

# 데이터 로드 실행
apt_map, err = load_apt_map()
if err:
    st.error(err)
    st.stop()
else:
    st.success(f"✅ A파일 로드 완료 (아파트 {len(apt_map):,}개 매핑)")

# 유틸리티 함수들
def is_게첨리스트(fname):
    return '광고게첨리스트' in fname

def fuzzy_match(ad_name, keys):
    ad_clean = ad_name.replace(" ", "").strip()
    for k in keys:
        if ad_clean in k.replace(" ", "") or k.replace(" ", "") in ad_clean:
            return k
    return None

def find_apts_auto(uploaded_file, apt_map):
    """광고 파일 내에서 아파트 목록 자동 추출"""
    try:
        xl = pd.ExcelFile(uploaded_file)
        best_apts = []
        best_sheet = ""
        for sheet in xl.sheet_names:
            df = pd.read_excel(uploaded_file, sheet_name=sheet)
            for col in df.columns:
                matches = [str(v).strip() for v in df[col].dropna() if str(v).strip() in apt_map]
                if len(matches) > len(best_apts):
                    best_apts = list(dict.fromkeys(matches))
                    best_sheet = sheet
        return best_apts, best_sheet
    except Exception:
        return [], None

def process_b_file(b_file, ad_file_apts, apt_map, apt_freq, is_게첨):
    b_file.seek(0)
    wb = load_workbook(b_file)
    ws = wb.active
    
    # 설정: 게첨리스트(O, S열) vs 일반(O, X열)
    COL_CENTER = 15 # O열
    COL_APT = 19 if is_게첨 else 24 # S열 또는 X열
    
    results = []
    used_g_per_ad = {}

    for row_idx in range(2, ws.max_row + 1):
        ad_name_cell = ws.cell(row=row_idx, column=6) # F열(광고명)
        if not ad_name_cell.value: continue
        ad_name = str(ad_name_cell.value).strip()
        
        # 매칭되는 광고 파일 찾기
        matched_key = fuzzy_match(ad_name, list(ad_file_apts.keys()))
        if not matched_key:
            results.append({"행": row_idx, "광고명": ad_name, "상태": "광고파일 없음"})
            continue
            
        candidates = ad_file_apts[matched_key]
        if matched_key not in used_g_per_ad: used_g_per_ad[matched_key] = set()
        
        # 동선 효율(빈도) 순 정렬하여 배정
        sorted_apts = sorted(candidates, key=lambda x: apt_freq.get(x, 0), reverse=True)
        chosen = None
        for apt in sorted_apts:
            g_val, center = apt_map[apt]
            if g_val not in used_g_per_ad[matched_key]:
                chosen = apt
                used_g_per_ad[matched_key].add(g_val)
                # O열에 센터명, S/X열에 아파트명 기입
                ws.cell(row=row_idx, column=COL_CENTER).value = center
                ws.cell(row=row_idx, column=COL_APT).value = chosen
                break
        
        results.append({"행": row_idx, "광고명": ad_name, "배정아파트": chosen if chosen else "중복 제외 실패"})
        
    return wb, results

# --- UI 구성 ---
col1, col2 = st.columns(2)
with col1:
    b_files = st.file_uploader("📋 B 파일들 업로드", type=["xlsx"], accept_multiple_files=True)
with col2:
    ad_files = st.file_uploader("📁 광고 파일들 업로드", type=["xlsx"], accept_multiple_files=True)

if b_files and ad_files:
    if st.button("🚀 최적화 실행"):
        ad_data = {}
        for f in ad_files:
            name = f.name.replace("송출요청서_", "").replace(".xlsx", "")
            apts, _ = find_apts_auto(f, apt_map)
            ad_data[name] = apts
            
        apt_freq = Counter([a for apts in ad_data.values() for a in apts])
        
        for b_f in b_files:
            is_gc = is_게첨리스트(b_f.name)
            processed_wb, res_log = process_b_file(b_f, ad_data, apt_map, apt_freq, is_gc)
            
            st.write(f"### {b_f.name} 처리 결과")
            st.table(pd.DataFrame(res_log).head(10)) # 상위 10개만 미리보기
            
            output = BytesIO()
            processed_wb.save(output)
            st.download_button(f"📥 {b_f.name} 결과 다운로드", output.getvalue(), f"결과_{b_f.name}")
