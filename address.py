import streamlit as st
import pandas as pd
import requests
from io import BytesIO
from pathlib import Path
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

def _kakao_keyword_search(headers, query):
    """카카오 키워드 검색 1건 실행 후 도로명 주소를 반환. 실패 시 None."""
    try:
        res = requests.get("https://dapi.kakao.com/v2/local/search/keyword.json",
                            headers=headers, params={"query": query, "size": 1}).json()
        if res.get('documents'):
            doc = res['documents'][0]
            addr = doc.get('road_address_name') or doc.get('address_name')
            return addr or None
    except:
        pass
    return None


@st.cache_data(ttl=3600, show_spinner=False)
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

    # [Step 2] 1차 검색 — 분지점이 비어있으면 '본사'를 기본값으로 사용
    #   (다지점 기업은 분지점을 명시하지 않으면 검색 결과가 엉뚱한 지방으로 튈 수 있음)
    city_hint = ledger_addr.split()[0] if ledger_addr else ""
    effective_branch = branch_name.strip() if branch_name and branch_name.strip() else "본사"
    search_query_1 = f"{city_hint} {company_name} {effective_branch}".strip()
    search_method = f"기업명+{effective_branch}"

    verified_addr = _kakao_keyword_search(headers, search_query_1)

    # [Step 3] 1차 검색이 '검색불가'면 — 분지점/본사 없이 기업명만으로 재검색
    #   (일부 기업은 '본사' 키워드가 붙으면 카카오 검색이 오히려 실패하는 경우가 있음)
    if not verified_addr:
        search_query_2 = f"{city_hint} {company_name}".strip()
        retry_addr = _kakao_keyword_search(headers, search_query_2)
        if retry_addr:
            verified_addr = retry_addr
            search_method = "기업명만(재검색)"

    if not verified_addr:
        verified_addr = "❌ 검색불가"
        search_method = "검색실패(2회 시도)"

    # [Step 4] 두 표준 주소 간 유사도 측정
    similarity = 0
    if standard_ledger_addr != "❌ 장부주소 불명" and verified_addr != "❌ 검색불가":
        similarity = get_similarity(standard_ledger_addr, verified_addr)

    return standard_ledger_addr, verified_addr, similarity, search_method

# ---------------------------------------------------------
# 3. 샘플 명단 로드 및 xlsx 템플릿 생성
# ---------------------------------------------------------

SAMPLE_PATH = Path(__file__).parent / "조회처_명단_템플릿.xlsx"
REQUIRED_COLS = ["기업명", "분지점", "주소", "전자조회가능회사"]

@st.cache_data
def load_sample_df():
    """repo에 포함된 샘플 명단을 읽어온다 (템플릿·샘플 실행의 단일 기준)."""
    df = pd.read_excel(SAMPLE_PATH)
    df.columns = [str(c).strip() for c in df.columns]
    for col in REQUIRED_COLS:
        if col not in df.columns:
            df[col] = None
    return df[REQUIRED_COLS]

def make_template_bytes():
    template_df = load_sample_df().fillna("")
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        template_df.to_excel(writer, index=False, sheet_name="조회처명단")
        workbook = writer.book
        worksheet = writer.sheets["조회처명단"]
        header_fmt = workbook.add_format({
            "bold": True, "bg_color": "#1F4E79", "font_color": "white",
            "border": 1, "align": "center", "valign": "vcenter"
        })
        example_fmt = workbook.add_format({"font_color": "#808080", "italic": True})
        for col_idx, col_name in enumerate(template_df.columns):
            worksheet.write(0, col_idx, col_name, header_fmt)
            worksheet.set_column(col_idx, col_idx, 26)
        for row_idx in range(1, len(template_df) + 1):
            for col_idx in range(len(template_df.columns)):
                worksheet.write(row_idx, col_idx, template_df.iloc[row_idx - 1, col_idx], example_fmt)
        worksheet.freeze_panes(1, 0)
    return output.getvalue()

# ---------------------------------------------------------
# 4. 검증 실행 (업로드/샘플 공용)
# ---------------------------------------------------------

def prepare_inputs(raw_df):
    raw_df = raw_df.copy()
    raw_df.columns = [str(c).strip() for c in raw_df.columns]
    df_main = raw_df[raw_df['기업명'].notna()].copy()
    e_list = raw_df['전자조회가능회사'].dropna().unique().tolist() if '전자조회가능회사' in raw_df.columns else []
    return df_main, e_list

def run_validation(df_main, e_list):
    results_list = []
    progress_bar = st.progress(0)
    total = len(df_main)

    for n, (_, row) in enumerate(df_main.iterrows(), start=1):
        c_name = str(row['기업명']).strip()
        b_name = str(row['분지점']).strip() if '분지점' in row and pd.notna(row['분지점']) else ""
        ledger_addr = str(row['주소']).strip()

        # 전자조회 체크
        is_e = any(c_name in str(org) or str(org) in c_name for org in e_list)

        # 2중 검증 실행 (본사 기본 검색 → 실패 시 기업명만으로 재검색)
        std_ledger, v_addr, sim, method = get_double_validated_address(c_name, b_name, ledger_addr)

        results_list.append({
            "기업명": c_name,
            "장부 주소(Original)": ledger_addr,
            "표준화 주소(장부)": std_ledger,
            "검색된 주소(API)": v_addr,
            "검색방식": method,
            "유사도": f"{sim}%",
            "최종판정": "✅ 일치" if sim >= 80 else "🚨 확인필요",
            "전자조회": "🔵 가능" if is_e else "⚪ 서면"
        })
        progress_bar.progress(n / total)

    progress_bar.empty()
    return pd.DataFrame(results_list)

# ---------------------------------------------------------
# 5. 페이지 설정 및 헤더
# ---------------------------------------------------------
st.set_page_config(page_title="조회서 실재성 검증 시스템", page_icon="🛡️", layout="wide")

with st.sidebar:
    st.header("🛡️ 조회서 실재성 검증")
    st.caption("감사 조회처 주소를 Kakao 지도 API로 자동 교차검증하는 도구입니다.")
    st.markdown("---")
    st.markdown(
        "**진행 순서**\n"
        "1. 템플릿 다운로드\n"
        "2. 조회처 명단 작성\n"
        "3. 파일 업로드\n"
        "4. 검증 실행\n"
        "5. 결과 확인 및 다운로드"
    )
    st.markdown("---")
    st.caption("장부상 주소 ↔ 기업 검색 주소를 대조하여 '지방 튐'(주소 불일치) 현상을 찾아냅니다.")

st.title("🛡️ 조회서 실재성 2중 교차 검증 시스템")
st.caption("장부상 주소와 기업 검색 결과를 API 기반으로 교차 대조하여 '지방 튐' 현상을 방지합니다.")

# --- 샘플 즉시 실행 ---
sample_df = load_sample_df()
st.info(f"처음이시면 업로드 없이 샘플 **{len(sample_df)}건**으로 바로 실행해 보세요.")
if st.button(f"⚡ 샘플 {len(sample_df)}건으로 즉시 검증", type="primary", use_container_width=True):
    df_main, e_list = prepare_inputs(sample_df)
    with st.spinner("샘플 명단 2중 교차 검증 중..."):
        st.session_state.final_results = run_validation(df_main, e_list)

with st.expander("📖 사용법 보기", expanded=False):
    st.markdown(
        """
1. **템플릿 다운로드** — 아래 ①에서 xlsx 템플릿을 내려받아 조회처 명단을 정리합니다.
2. **필수 입력 항목**

| 컬럼명 | 설명 | 필수 여부 |
|---|---|---|
| 기업명 | 조회 대상 회사명 | 필수 |
| 분지점 | 지점·사업소명 등 | 선택 |
| 주소 | 장부상 주소 | 필수 |
| 전자조회가능회사 | 전자조회가 가능한 회사명 목록 | 선택 |

> ⚠️ **다지점 기업 주의사항**: 지점이 여러 곳인 기업은 '분지점'을 비워두면 시스템이 기본값으로 '본사'를 붙여 검색합니다. 그래도 검색이 안 되는 경우 기업명만으로 자동 재검색하지만, 그 결과가 실제 조회 대상 지점과 다를 수 있으니 **가능하면 '분지점'란에 정확한 지점명을 직접 입력**해야 정확도가 올라갑니다.

3. **파일 업로드** — ②에서 작성한 엑셀 파일을 업로드합니다.
4. **검증 실행** — '2중 교차 검증 시작' 버튼을 누르면 장부 주소와 기업명 검색 결과를 각각 표준 주소로 변환한 뒤 유사도를 비교합니다. 기업명 검색은 1차로 '분지점(기본값: 본사)'을 포함해 시도하고, 검색불가 시 분지점 없이 기업명만으로 2차 재검색합니다.
5. **결과 확인** — 유사도 80% 미만인 건은 🚨 확인필요로 표시되며, '검색방식' 열에서 1차(본사 포함)/2차(기업명만 재검색) 여부를 확인할 수 있습니다. 결과표는 엑셀로 다운로드할 수 있습니다.
        """
    )

st.divider()

col1, col2 = st.columns(2)
with col1:
    st.subheader("① 템플릿 준비")
    st.write("조회처 명단 작성용 xlsx 템플릿을 다운로드하세요.")
    st.download_button(
        "📄 xlsx 템플릿 다운로드",
        data=make_template_bytes(),
        file_name="조회처_명단_템플릿.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
with col2:
    st.subheader("② 파일 업로드")
    uploaded_file = st.file_uploader("분석할 엑셀 파일 업로드 (xlsx)", type=['xlsx'])

if 'final_results' not in st.session_state:
    st.session_state.final_results = None

if uploaded_file:
    raw_df = pd.read_excel(uploaded_file)
    df_main, e_list = prepare_inputs(raw_df)

    st.divider()
    st.subheader("③ 검증 실행")
    st.write(f"업로드된 조회처: **{len(df_main)}건**")

    if st.button("🚀 2중 교차 검증 시작", use_container_width=True):
        st.session_state.final_results = run_validation(df_main, e_list)

if st.session_state.final_results is not None:
    result_df = st.session_state.final_results
    st.divider()
    st.subheader("📊 2중 교차 검증 리포트")

    total_cnt = len(result_df)
    match_cnt = int((result_df["최종판정"] == "✅ 일치").sum())
    check_cnt = total_cnt - match_cnt

    m1, m2, m3 = st.columns(3)
    m1.metric("전체 건수", f"{total_cnt}건")
    m2.metric("✅ 일치", f"{match_cnt}건")
    m3.metric("🚨 확인필요", f"{check_cnt}건")

    st.dataframe(result_df, use_container_width=True, hide_index=True)

    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        result_df.to_excel(writer, index=False)
    st.download_button("📥 검증 결과 다운로드", output.getvalue(), "Double_Check_Results.xlsx", use_container_width=True)
