import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="주식 가치평가 청사진 스캐너", layout="wide")

@st.cache_data
def load_excel_data(filepath="◆종목관리◆.xlsx"):
    df_raw = pd.read_excel(filepath, sheet_name="품질_가격_추세", header=None)
    df_data = df_raw.iloc[2:].copy()
    df_data.columns = range(117)
    
    clean = pd.DataFrame()
    clean["종목코드"] = df_data[0].astype(str).str.zfill(6)
    clean["종목명"] = df_data[1].astype(str)
    clean["업종"] = df_data[2].fillna("-").astype(str)
    clean["퀀트등급"] = df_data[3].fillna("-").astype(str)
    clean["성장점수"] = pd.to_numeric(df_data[4], errors="coerce")
    clean["TTM_ROE(%)"] = pd.to_numeric(df_data[6], errors="coerce") * 100
    clean["PBR"] = pd.to_numeric(df_data[7], errors="coerce")
    clean["SRIM_주주가치(억)"] = pd.to_numeric(df_data[9], errors="coerce")
    clean["10년_ROE(%)"] = pd.to_numeric(df_data[11], errors="coerce")
    clean["5년_ROE(%)"] = pd.to_numeric(df_data[12], errors="coerce")
    clean["평가자본_26_2Q(억)"] = pd.to_numeric(df_data[107], errors="coerce")
    clean["합산시총(억)"] = pd.to_numeric(df_data[111], errors="coerce")
    clean["상태"] = df_data[116].astype(str)
    return clean, df_data

st.title("📐 주식 가치평가 청사진 — 엑셀 연동 웹 스캐너")
clean, df_raw_rows = load_excel_data()

with st.sidebar:
    st.header("⚙️ 가치평가 파라미터 (청사진 공식)")
    roe_option = st.selectbox(
        "적용 ROE 기준",
        ["보수적: MIN(5년 ROE, TTM ROE)", "5년 평균 ROE", "10년 평균 ROE", "최근 TTM ROE"]
    )
    r_rate = st.slider("할인율 / 요구수익률 (r, %)", 6.0, 15.0, 9.0, 0.5) / 100.0
    n_years = st.slider("지속 기간 (n, 년)", 3, 10, 5, 1)
    min_roe = st.slider("최소 ROE 커트라인 (%)", 0.0, 25.0, 9.0, 1.0)
    min_mos = st.slider("최소 안전마진율 (%)", -30.0, 70.0, 20.0, 5.0)
    min_mcap = st.number_input("최소 합산 시가총액 (억원)", value=3000, step=500)
    grades = st.multiselect("퀀트 등급 필터", ["S", "A", "B", "C", "D", "B-"], default=["S", "A"])
    exclude_holding = st.checkbox("지주사 업종 제외", value=True)

if roe_option == "보수적: MIN(5년 ROE, TTM ROE)":
    clean["적용_ROE(%)"] = np.minimum(clean["5년_ROE(%)"], clean["TTM_ROE(%)"])
elif roe_option == "5년 평균 ROE":
    clean["적용_ROE(%)"] = clean["5년_ROE(%)"]
elif roe_option == "10년 평균 ROE":
    clean["적용_ROE(%)"] = clean["10년_ROE(%)"]
else:
    clean["적용_ROE(%)"] = clean["TTM_ROE(%)"]

valid = clean[clean["상태"] == "계산 가능"].copy()
valid["청사진_PV(억)"] = valid["평가자본_26_2Q(억)"] * (((1 + valid["적용_ROE(%)"]/100.0) ** n_years) / ((1 + r_rate) ** n_years))
valid["청사진_안전마진(%)"] = (1 - valid["합산시총(억)"] / valid["청사진_PV(억)"]) * 100

cond = (
    valid["퀀트등급"].isin(grades) &
    (valid["적용_ROE(%)"] >= min_roe) &
    (valid["청사진_안전마진(%)"] >= min_mos) &
    (valid["합산시총(억)"] >= min_mcap)
)
if exclude_holding:
    cond = cond & (~valid["업종"].str.contains("지주"))

result = valid[cond].sort_values(by="청사진_안전마진(%)", ascending=False)
st.subheader(f"🔍 스캔 결과: 총 {len(result)}개 저평가 우량 기업 발굴")
st.dataframe(result[["종목코드", "종목명", "업종", "퀀트등급", "성장점수", "적용_ROE(%)", "PBR", "평가자본_26_2Q(억)", "합산시총(억)", "청사진_PV(억)", "청사진_안전마진(%)"]], use_container_width=True)
