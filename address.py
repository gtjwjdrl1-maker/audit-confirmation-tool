import streamlit as st
import pandas as pd
import requests
from io import BytesIO
from difflib import SequenceMatcher

# ---------------------------------------------------------
# 1. 보안 설정 및 API 로드
# ---------------------------------------------------------
try:
    KAKAO_API_KEY = st.secrets["KAKAO_API_KEY"]
except KeyError:
    st.error("🚨 API 키 설정(Secrets)을 확인해주세요.")
    st.stop()

def get_similarity(a, b):
    # 공백 및 행정구역 명칭 차이 제거 후 비교
    a, b = str(a).replace(" ", ""), str(b).replace(" ", "")
    for word in ["경기도", "서울특별시", "인천광역시", "부산광역시"]:
        a, b = a.replace(word, ""), b.replace(word, "")
    return int(SequenceMatcher(None, a, b).ratio() * 100)

# ---------------------------------------------------------
# 2. 2중 교차 검증 핵심 로직
# ---------------------------------------------------------

def get_double_validated_address(company_name, branch_name, ledger_addr):
    headers = {"Authorization": f"KakaoAK {KAKAO_API_KEY}"}
    
    # [Step 1] 장부 주소를 API로 검색하여 '표준 주소' 획득
    standard_ledger_addr = "❌ 장부주소 불명"
    try:
        addr_res = requests.get("https://dapi.kakao.com/v2/local/search/address.json", 
                                headers=headers, params={"query": ledger_addr, "size": 1}).json()
        if addr_res.get('documents'):
            standard_ledger_addr = addr_res['documents'][0]['road_address']['address_name'] if addr_res['documents'][0]['road_address'] else addr_res['documents'][0]['address_name']
    except: pass

    # [Step 2] 기업명(+지역힌트)으로 검색하여 '검증 주소' 획득
    verified_addr = "❌ 검색불가"
    city_hint = ledger_addr.split()[0] if ledger_addr else ""
    search_query = f"{city_hint} {company_name} {branch_name or ''}".strip()
    
    try:
        name_res = requests.get("https://dapi.kakao.com/v2/local/search/keyword.json", 
                                headers=headers, params={"query": search_query, "size": 1}).json()
        if name_res.get('documents'):
            verified_addr = name_res['documents'][0]['road_address_name']
    except: pass

    # [Step 3] 두 표준 주소 간 유사도 측정
    similarity = 0
    if standard_ledger_addr != "❌ 장부주소 불명" and verified_addr != "❌ 검색불가":
        similarity = get_similarity(standard_ledger_addr, verified_addr)
    
    return standard_ledger_addr, verified_addr, similarity

# ---------------------------------------------------------
# 3. UI 및 실행부
# ---------------------------------------------------------
st.set_page_config(page_title="조회서 2중 검증 시스템 V13", layout="wide")
st.title("🛡️ 조회서 실재성 2중 교차 검증 시스템")
st.info("장부 주소와 기업 검색 결과를 API 기반으로 교차 대조하여 '지방 튐' 현상을 방지합니다.")

if 'final_results' not in st.session_state:
    st.session_state.final_results = None

uploaded_file = st.file_uploader("파일럿 테스트.xlsx 업로드", type=['xlsx'])

if uploaded_file:
    raw_df = pd.read_excel(uploaded_file)
    raw_df.columns = [c.strip() for c in raw_df.columns]
    df_main = raw_df[raw_df['기업명'].notna()].copy()
    e_list = raw_df['전자조회가능회사'].dropna().unique().tolist() if '전자조회가능회사' in raw_df.columns else []

    if st.button("🚀 2중 교차 검증 시작"):
        results_list = []
        progress_bar = st.progress(0)
        
        for i, row in df_main.iterrows():
            c_name = str(row['기업명']).strip()
            b_name = str(row['분지점']).strip() if '분지점' in row and pd.notna(row['분지점']) else ""
            ledger_addr = str(row['주소']).strip()
            
            # 전자조회 체크
            is_e = any(c_name in str(org) or str(org) in c_name for org in e_list)
            
            # 2중 검증 실행
            std_ledger, v_addr, sim = get_double_validated_address(c_name, b_name, ledger_addr)
            
            results_list.append({
                "기업명": c_name,
                "장부 주소(Original)": ledger_addr,
                "표준화 주소(장부)": std_ledger,
                "검색된 주소(API)": v_addr,
                "유사도": f"{sim}%",
                "최종판정": "✅ 일치" if sim >= 80 else "🚨 확인필요",
                "전자조회": "🔵 가능" if is_e else "⚪ 서면"
            })
            progress_bar.progress((i + 1) / len(df_main))

        st.session_state.final_results = pd.DataFrame(results_list)

if st.session_state.final_results is not None:
    st.markdown("---")
    st.subheader("📊 2중 교차 검증 리포트")
    st.table(st.session_state.final_results)
    
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        st.session_state.final_results.to_excel(writer, index=False)
    st.download_button("📥 검증 결과 다운로드", output.getvalue(), "Double_Check_Results.xlsx")
