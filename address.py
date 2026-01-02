import streamlit as st
import pandas as pd
import requests
from io import BytesIO
from difflib import SequenceMatcher

# ---------------------------------------------------------
# 1. 보안 설정 (st.secrets를 사용하여 키 로드)
# ---------------------------------------------------------
try:
    # Streamlit Secrets에서 키를 가져옵니다.
    JUSO_API_KEY = st.secrets["JUSO_API_KEY"]
    KAKAO_API_KEY = st.secrets["KAKAO_API_KEY"]
except KeyError:
    st.error("🚨 API 키 설정이 발견되지 않았습니다. .streamlit/secrets.toml 파일 혹은 Streamlit Cloud 설정을 확인해주세요.")
    st.stop()

def get_similarity(a, b):
    a, b = str(a).replace(" ", ""), str(b).replace(" ", "")
    return int(SequenceMatcher(None, a, b).ratio() * 100)

if 'final_results' not in st.session_state:
    st.session_state.final_results = None

# ---------------------------------------------------------
# 2. UI 및 로직 (V10과 동일하지만 키 로딩 방식만 변경됨)
# ---------------------------------------------------------
st.set_page_config(page_title="조회서 검증 V11 (보안)", layout="wide")
st.title("🛡️ 조회서 검증 시스템 (API 보안 모드)")

uploaded_file = st.file_uploader("파일럿 테스트.xlsx 업로드", type=['xlsx'])

if uploaded_file:
    try:
        raw_df = pd.read_excel(uploaded_file)
        raw_df.columns = [c.strip() for c in raw_df.columns]
        df_main = raw_df[raw_df['기업명'].notna()].copy()
        e_list = raw_df['전자조회가능회사'].dropna().unique().tolist() if '전자조회가능회사' in raw_df.columns else []
        st.info(f"분석 준비 완료: {len(df_main)}건")
    except Exception as e:
        st.error(f"파일 로드 에러: {e}")
        st.stop()

    if st.button("🚀 분석 실행"):
        results_list = []
        progress_bar = st.progress(0)
        
        for i, row in df_main.iterrows():
            c_name = str(row['기업명']).strip()
            b_name = str(row['분지점']).strip() if '분지점' in row and pd.notna(row['분지점']) else ""
            addr_orig = str(row['주소']).strip()
            
            # 전자조회 체크
            e_status = "🔵 가능" if any(c_name in str(org) or str(org) in c_name for org in e_list) else "⚪ 서면"
            
            # API 호출
            headers = {"Authorization": f"KakaoAK {KAKAO_API_KEY}"}
            query = f"{c_name} {b_name}".strip()
            kakao_addr, sim_score = "❌ 검색불가", 0
            
            try:
                res = requests.get("https://dapi.kakao.com/v2/local/search/keyword.json", 
                                   headers=headers, params={"query": query, "size": 1}, timeout=5).json()
                if res.get('documents'):
                    kakao_addr = res['documents'][0]['road_address_name']
                    sim_score = get_similarity(addr_orig, kakao_addr)
            except: pass
            
            results_list.append({
                "기업명": c_name, "장부주소": addr_orig, "전자조회": e_status,
                "검증주소": kakao_addr, "유사도": f"{sim_score}%",
                "판정": "✅ 일치" if sim_score >= 70 else "🚨 확인"
            })
            progress_bar.progress((i + 1) / len(df_main))

        st.session_state.final_results = pd.DataFrame(results_list)

if st.session_state.final_results is not None:
    st.markdown("---")
    st.subheader("📊 검증 결과 리포트")
    st.table(st.session_state.final_results)
    
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        st.session_state.final_results.to_excel(writer, index=False)
    st.download_button("📥 결과 엑셀 다운로드", output.getvalue(), "audit_results.xlsx")