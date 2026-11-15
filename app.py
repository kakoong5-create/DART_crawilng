import os
import json
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import numpy as np

st.set_page_config(
    page_title="주식 가치평가 청사진 — 퀀트 스캐너 & 10년 추세 대시보드",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Streamlit 기본 여백 제거 및 다크 테마 배경 일치화
st.markdown("""
<style>
  .block-container { padding-top: 0.8rem !important; padding-bottom: 0rem !important; padding-left: 1rem !important; padding-right: 1rem !important; max-width: 100% !important; }
  header[data-testid="stHeader"] { background: transparent !important; height: 1.5rem !important; }
  .stApp { background-color: #0f172a; }
</style>
""", unsafe_allow_html=True)

# 폴더 내 엑셀 파일 자동 탐색 (stock_master.xlsx, 종목관리.xlsx, ◆종목관리◆.xlsx 모두 호환)
def find_excel_file():
    candidates = ["◆종목관리◆.xlsx", "stock_master.xlsx", "종목관리.xlsx"]
    for c in candidates:
        if os.path.exists(c):
            return c
    for f in os.listdir("."):
        if f.endswith(".xlsx") and "종목코드" not in f:
            return f
    return "◆종목관리◆.xlsx"

EXCEL_PATH = find_excel_file()

@st.cache_data
def load_stocks_for_html(filepath, mtime):
    df_raw = pd.read_excel(filepath, sheet_name="품질_가격_추세", header=None)
    df_data = df_raw.iloc[2:].copy()
    df_data.columns = range(117)
    valid_rows = df_data[df_data[116] == "계산 가능"].copy()

    stocks_json = []
    for _, row in valid_rows.iterrows():
        code = str(row[0]).zfill(6)
        name = str(row[1])
        sector = str(row[2]) if pd.notna(row[2]) else "-"
        q_grade = str(row[3]) if pd.notna(row[3]) else "-"
        g_score = round(float(row[4]), 1) if pd.notna(row[4]) else 0.0
        ttm_roe = round(float(row[6]) * 100, 2) if pd.notna(row[6]) else 0.0
        pbr = round(float(row[7]), 2) if pd.notna(row[7]) else 0.0
        srim_val = round(float(row[9]), 1) if pd.notna(row[9]) else 0.0
        roe_10y = round(float(row[11]), 2) if pd.notna(row[11]) else 0.0
        roe_5y = round(float(row[12]), 2) if pd.notna(row[12]) else 0.0
        cap_26_2q = round(float(row[107]), 1) if pd.notna(row[107]) else 0.0
        mcap_total = round(float(row[111]), 1) if pd.notna(row[111]) else 0.0
        srim_mos = round(float(row[112]) * 100, 1) if pd.notna(row[112]) else 0.0

        roe_hist = [round(float(row[c]), 1) if pd.notna(row[c]) else None for c in range(13, 23)]
        rev_hist = [round(float(row[c]), 0) if pd.notna(row[c]) else None for c in range(27, 37)]
        ni_hist = [round(float(row[c]), 0) if pd.notna(row[c]) else None for c in range(44, 54)]
        eq_hist = [round(float(row[c]), 0) if pd.notna(row[c]) else None for c in range(61, 70)] + [cap_26_2q]
        debt_26_2q = round(float(row[105]) * 100, 1) if pd.notna(row[105]) else 0.0

        stocks_json.append({
            "code": code,
            "name": name,
            "sector": sector,
            "grade": q_grade,
            "growth": g_score,
            "ttm_roe": ttm_roe,
            "roe_5y": roe_5y,
            "roe_10y": roe_10y,
            "pbr": pbr,
            "capital": cap_26_2q,
            "mcap": mcap_total,
            "srim_pv": srim_val,
            "srim_mos": srim_mos,
            "debt_ratio": debt_26_2q,
            "roe_h": roe_hist,
            "rev_h": rev_hist,
            "ni_h": ni_hist,
            "eq_h": eq_hist
        })
    return stocks_json

if not os.path.exists(EXCEL_PATH):
    st.error(f"엑셀 파일({EXCEL_PATH})을 찾을 수 없습니다. GitHub 저장소에 엑셀 파일이 올라가 있는지 확인해 주세요.")
    st.stop()

mtime = os.path.getmtime(EXCEL_PATH)
stocks_data = load_stocks_for_html(EXCEL_PATH, mtime)

html_code = f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<style>
  :root {{
    --bg: #0f172a;
    --panel: #1e293b;
    --border: #334155;
    --text: #f8fafc;
    --muted: #94a3b8;
    --accent: #38bdf8;
    --green: #22c55e;
    --amber: #f59e0b;
    --red: #ef4444;
  }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; }}
  body {{ background: var(--bg); color: var(--text); padding: 10px 14px; }}
  .header {{ display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid var(--border); padding-bottom: 12px; margin-bottom: 16px; }}
  .header h1 {{ font-size: 21px; color: var(--accent); }}
  .header p {{ font-size: 12.5px; color: var(--muted); margin-top: 4px; }}
  .layout {{ display: grid; grid-template-columns: 300px 1fr; gap: 16px; }}
  .sidebar {{ background: var(--panel); border: 1px solid var(--border); border-radius: 10px; padding: 15px; height: fit-content; }}
  .sidebar h3 {{ font-size: 14.5px; margin-bottom: 12px; color: var(--accent); border-bottom: 1px solid var(--border); padding-bottom: 6px; }}
  .control-group {{ margin-bottom: 12px; }}
  .control-group label {{ display: flex; justify-content: space-between; font-size: 12.5px; color: var(--muted); margin-bottom: 4px; }}
  .control-group label span {{ color: var(--text); font-weight: 600; }}
  .control-group input[type=range], .control-group select, .control-group input[type=text] {{
    width: 100%; padding: 6px; background: #0f172a; border: 1px solid var(--border); color: var(--text); border-radius: 6px; font-size: 12.5px;
  }}
  .grade-checks {{ display: flex; gap: 6px; flex-wrap: wrap; }}
  .grade-checks label {{ background: #0f172a; padding: 4px 8px; border-radius: 6px; border: 1px solid var(--border); cursor: pointer; font-size: 11.5px; color: var(--text); }}
  .main {{ display: flex; flex-direction: column; gap: 16px; }}
  .kpi-row {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; }}
  .kpi-card {{ background: var(--panel); border: 1px solid var(--border); border-radius: 10px; padding: 12px 14px; }}
  .kpi-card .label {{ font-size: 12px; color: var(--muted); }}
  .kpi-card .val {{ font-size: 21px; font-weight: 700; margin-top: 4px; color: var(--accent); }}
  .grid-2 {{ display: grid; grid-template-columns: 1.15fr 1fr; gap: 16px; }}
  .card {{ background: var(--panel); border: 1px solid var(--border); border-radius: 10px; padding: 14px; }}
  .card h3 {{ font-size: 14.5px; margin-bottom: 10px; color: var(--text); display: flex; justify-content: space-between; align-items: center; }}
  .table-wrap {{ max-height: 560px; overflow-y: auto; border: 1px solid var(--border); border-radius: 8px; }}
  table {{ width: 100%; border-collapse: collapse; font-size: 12px; }}
  th {{ position: sticky; top: 0; background: #0f172a; color: var(--muted); padding: 8px 6px; text-align: right; border-bottom: 1px solid var(--border); }}
  th:nth-child(1), th:nth-child(2), th:nth-child(3) {{ text-align: left; }}
  td {{ padding: 7px 6px; text-align: right; border-bottom: 1px solid #263349; }}
  td:nth-child(1), td:nth-child(2), td:nth-child(3) {{ text-align: left; }}
  tr:hover {{ background: #28364f; cursor: pointer; }}
  .badge {{ padding: 2px 6px; border-radius: 4px; font-weight: 700; font-size: 11px; display: inline-block; }}
  .badge-S {{ background: rgba(34,197,94,0.2); color: #4ade80; border: 1px solid #22c55e; }}
  .badge-A {{ background: rgba(56,189,248,0.2); color: #38bdf8; border: 1px solid #0284c7; }}
  .badge-B {{ background: rgba(245,158,11,0.2); color: #fbbf24; border: 1px solid #d97706; }}
  .badge-other {{ background: rgba(148,163,184,0.15); color: #94a3b8; }}
  .pos {{ color: #4ade80; font-weight: 600; }}
  .neg {{ color: #f87171; }}
  .formula-box {{ background: #0f172a; border: 1px solid var(--border); padding: 10px; border-radius: 8px; font-size: 11.5px; color: var(--muted); line-height: 1.5; margin-top: 10px; }}
</style>
</head>
<body>

<div class="header">
  <div>
    <h1>📐 주식 가치평가 청사진 — 전 종목 퀀트 스캐너 & 10년 추세 대시보드</h1>
    <p>실시간 연동 파일: {EXCEL_PATH} (정상 평가대상 {len(stocks_data):,}개 기업) | 공식: PV = 자본 × ((1 + ROE)^n / (1 + 할인율)^n)</p>
  </div>
  <div style="text-align:right; font-size:12px; color:var(--muted);">
    기준 재무: 2016 ~ 2026.2Q<br>표에서 종목 클릭 시 우측 10년 재무 차트 즉시 연동
  </div>
</div>

<div class="layout">
  <div class="sidebar">
    <h3>⚙️ 청사진 가치평가 파라미터</h3>
    <div class="control-group">
      <label>검색 (종목명 / 코드 / 업종)</label>
      <input type="text" id="searchInput" placeholder="예: 현대차, 반도체, 기아..." oninput="updateDashboard()">
    </div>
    <div class="control-group">
      <label>적용 ROE 기준 선택</label>
      <select id="roeMode" onchange="updateDashboard()">
        <option value="min_5y_ttm" selected>보수적: MIN(5년 평균 ROE, TTM ROE)</option>
        <option value="roe_5y">5년 평균 ROE (2021~2025)</option>
        <option value="roe_10y">10년 평균 ROE (2016~2025)</option>
        <option value="ttm_roe">최근 TTM ROE (26.2Q 기준)</option>
      </select>
    </div>
    <div class="control-group">
      <label>요구수익률 / 할인율 (r) <span id="rVal">9.0%</span></label>
      <input type="range" id="discountRate" min="6.0" max="15.0" step="0.5" value="9.0" oninput="updateDashboard()">
    </div>
    <div class="control-group">
      <label>지속 기간 (n년) <span id="nVal">5년</span></label>
      <input type="range" id="yearsN" min="3" max="10" step="1" value="5" oninput="updateDashboard()">
    </div>
    <div class="control-group">
      <label>최소 적용 ROE 커트라인 <span id="minRoeVal">9.0%</span></label>
      <input type="range" id="minRoe" min="0" max="25" step="1" value="9" oninput="updateDashboard()">
    </div>
    <div class="control-group">
      <label>최소 안전마진율 (MoS) <span id="minMosVal">20%</span></label>
      <input type="range" id="minMos" min="-50" max="70" step="5" value="20" oninput="updateDashboard()">
    </div>
    <div class="control-group">
      <label>최소 합산 시가총액 (억원) <span id="minMcapVal">3,000억</span></label>
      <input type="range" id="minMcap" min="0" max="50000" step="1000" value="3000" oninput="updateDashboard()">
    </div>
    <div class="control-group">
      <label>퀀트 등급 필터 (ROE-PBR 매트릭스)</label>
      <div class="grade-checks">
        <label><input type="checkbox" class="grade-cb" value="S" checked onchange="updateDashboard()"> S등급</label>
        <label><input type="checkbox" class="grade-cb" value="A" checked onchange="updateDashboard()"> A등급</label>
        <label><input type="checkbox" class="grade-cb" value="B" checked onchange="updateDashboard()"> B등급</label>
        <label><input type="checkbox" class="grade-cb" value="C" onchange="updateDashboard()"> C/D/기타</label>
      </div>
    </div>
    <div class="control-group">
      <label><input type="checkbox" id="excludeHolding" checked onchange="updateDashboard()"> 지주사 업종 제외</label>
    </div>
    <div class="formula-box">
      <b>💡 주식_가치평가_청사진 점검 기준</b><br>
      • <b>Step 1 (품질):</b> ROE &gt; 할인율(r) 확인 (5% 이하 지하실 제외)<br>
      • <b>Step 2 (가격):</b> ROE 10%=PBR 1.0, ROE 20%=PBR 2.0 매트릭스 대비 저평가 확인<br>
      • <b>Step 3 (추세):</b> 우측 차트에서 10년간 자본과 ROE의 지속성(Y) 검증
    </div>
  </div>

  <div class="main">
    <div class="kpi-row">
      <div class="kpi-card">
        <div class="label">스캔 조건 부합 종목 수</div>
        <div class="val" id="kpiCount">0개</div>
      </div>
      <div class="kpi-card">
        <div class="label">스캔 종목 평균 ROE</div>
        <div class="val" id="kpiAvgRoe">0.0%</div>
      </div>
      <div class="kpi-card">
        <div class="label">스캔 종목 평균 PBR</div>
        <div class="val" id="kpiAvgPbr">0.00배</div>
      </div>
      <div class="kpi-card">
        <div class="label">평균 청사진 안전마진</div>
        <div class="val" id="kpiAvgMos">0.0%</div>
      </div>
    </div>

    <div class="grid-2">
      <div class="card">
        <h3>
          <span>📋 저평가 스캔 리스트 (안전마진 순)</span>
          <span style="font-size:11.5px; color:var(--muted);">종목 클릭 시 우측 10년 추세 분석</span>
        </h3>
        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>코드</th>
                <th>종목명</th>
                <th>업종</th>
                <th>등급</th>
                <th>성장점수</th>
                <th>적용ROE</th>
                <th>현재PBR</th>
                <th>합산시총(억)</th>
                <th>청사진PV(억)</th>
                <th>안전마진</th>
              </tr>
            </thead>
            <tbody id="stockTableBody"></tbody>
          </table>
        </div>
      </div>

      <div style="display:flex; flex-direction:column; gap:14px;">
        <div class="card">
          <h3 id="detailTitle">📈 개별 기업 10년 품질(ROE) & 자본 팽창 추세 (2016~2025)</h3>
          <div id="detailMeta" style="font-size:12px; color:var(--muted); margin-bottom:8px;"></div>
          <canvas id="trendChart" height="165"></canvas>
        </div>
        <div class="card">
          <h3>🎯 실전 투자 기준값: ROE-PBR 퀀트 매트릭스 분포 (청사진 Slide 12)</h3>
          <canvas id="scatterChart" height="155"></canvas>
        </div>
      </div>
    </div>
  </div>
</div>

<script>
const RAW_STOCKS = {json.dumps(stocks_data, ensure_ascii=False)};
let trendChartInstance = null;
let scatterChartInstance = null;
let selectedStockCode = "000270";

function getAppliedRoe(s, mode) {{
  if (mode === "min_5y_ttm") return Math.min(s.roe_5y, s.ttm_roe);
  if (mode === "roe_5y") return s.roe_5y;
  if (mode === "roe_10y") return s.roe_10y;
  return s.ttm_roe;
}}

function updateDashboard() {{
  const query = document.getElementById("searchInput").value.trim().toLowerCase();
  const roeMode = document.getElementById("roeMode").value;
  const r = parseFloat(document.getElementById("discountRate").value);
  const n = parseInt(document.getElementById("yearsN").value);
  const minRoe = parseFloat(document.getElementById("minRoe").value);
  const minMos = parseFloat(document.getElementById("minMos").value);
  const minMcap = parseFloat(document.getElementById("minMcap").value);
  const excludeHolding = document.getElementById("excludeHolding").checked;

  document.getElementById("rVal").innerText = r.toFixed(1) + "%";
  document.getElementById("nVal").innerText = n + "년";
  document.getElementById("minRoeVal").innerText = minRoe.toFixed(0) + "%";
  document.getElementById("minMosVal").innerText = minMos.toFixed(0) + "%";
  document.getElementById("minMcapVal").innerText = minMcap.toLocaleString() + "억";

  const checkedGrades = Array.from(document.querySelectorAll(".grade-cb:checked")).map(cb => cb.value);

  const filtered = [];
  for (const s of RAW_STOCKS) {{
    const appRoe = getAppliedRoe(s, roeMode);
    if (appRoe <= -90 || s.capital <= 0 || s.mcap <= 0) continue;

    const mult = Math.pow((1 + appRoe / 100.0) / (1 + r / 100.0), n);
    const pv = s.capital * mult;
    const mos = (1 - s.mcap / pv) * 100;

    const gradeGroup = ["S", "A", "B"].includes(s.grade) ? s.grade : "C";
    if (!checkedGrades.includes(gradeGroup)) continue;
    if (excludeHolding && s.sector.includes("지주")) continue;
    if (appRoe < minRoe) continue;
    if (mos < minMos) continue;
    if (s.mcap < minMcap) continue;
    if (query && !(s.name.toLowerCase().includes(query) || s.code.includes(query) || s.sector.toLowerCase().includes(query))) continue;

    filtered.push({{ ...s, appRoe, pv, mos }});
  }}

  filtered.sort((a, b) => b.mos - a.mos);

  document.getElementById("kpiCount").innerText = filtered.length.toLocaleString() + "개";
  if (filtered.length > 0) {{
    const avgRoe = filtered.reduce((acc, x) => acc + x.appRoe, 0) / filtered.length;
    const avgPbr = filtered.reduce((acc, x) => acc + x.pbr, 0) / filtered.length;
    const avgMos = filtered.reduce((acc, x) => acc + x.mos, 0) / filtered.length;
    document.getElementById("kpiAvgRoe").innerText = avgRoe.toFixed(1) + "%";
    document.getElementById("kpiAvgPbr").innerText = avgPbr.toFixed(2) + "배";
    document.getElementById("kpiAvgMos").innerText = avgMos.toFixed(1) + "%";
  }} else {{
    document.getElementById("kpiAvgRoe").innerText = "-";
    document.getElementById("kpiAvgPbr").innerText = "-";
    document.getElementById("kpiAvgMos").innerText = "-";
  }}

  const tbody = document.getElementById("stockTableBody");
  tbody.innerHTML = "";
  const displayList = filtered.slice(0, 200);
  for (const s of displayList) {{
    const tr = document.createElement("tr");
    tr.onclick = () => selectStock(s.code, r, n, roeMode);
    const badgeClass = ["S", "A", "B"].includes(s.grade) ? "badge-" + s.grade : "badge-other";
    tr.innerHTML = `
      <td>${{s.code}}</td>
      <td style="font-weight:600; color:#f8fafc;">${{s.name}}</td>
      <td style="color:#94a3b8; font-size:11px;">${{s.sector}}</td>
      <td><span class="badge ${{badgeClass}}">${{s.grade}}</span></td>
      <td>${{s.growth.toFixed(0)}}</td>
      <td>${{s.appRoe.toFixed(1)}}%</td>
      <td>${{s.pbr.toFixed(2)}}</td>
      <td>${{Math.round(s.mcap).toLocaleString()}}</td>
      <td style="color:#38bdf8; font-weight:600;">${{Math.round(s.pv).toLocaleString()}}</td>
      <td class="${{s.mos >= 0 ? 'pos' : 'neg'}}">${{s.mos.toFixed(1)}}%</td>
    `;
    tbody.appendChild(tr);
  }}

  renderScatter(filtered.slice(0, 120));
  const target = filtered.find(x => x.code === selectedStockCode) || filtered[0] || RAW_STOCKS[0];
  if (target) selectStock(target.code, r, n, roeMode);
}}

function selectStock(code, r, n, roeMode) {{
  selectedStockCode = code;
  const s = RAW_STOCKS.find(x => x.code === code);
  if (!s) return;
  const appRoe = getAppliedRoe(s, roeMode);
  const pv = s.capital * Math.pow((1 + appRoe / 100.0) / (1 + r / 100.0), n);
  const mos = (1 - s.mcap / pv) * 100;

  document.getElementById("detailTitle").innerText = `📈 [${{s.name}} (${{s.code}})] 10년 ROE 지속성(Y) 및 자본 팽창 추세`;
  document.getElementById("detailMeta").innerHTML = `
    업종: <b>${{s.sector}}</b> | 26.2Q 자본: <b>${{Math.round(s.capital).toLocaleString()}}억</b> | 합산시총: <b>${{Math.round(s.mcap).toLocaleString()}}억</b> |
    청사진 PV: <b style="color:#38bdf8">${{Math.round(pv).toLocaleString()}}억 (안전마진 ${{mos.toFixed(1)}}%)</b> | 기존 S-RIM: <b>${{Math.round(s.srim_pv).toLocaleString()}}억</b> | 26.2Q 부채비율: <b>${{s.debt_ratio}}%</b>
  `;

  const years = ["2016", "2017", "2018", "2019", "2020", "2021", "2022", "2023", "2024", "2025/26.2Q"];
  const ctx = document.getElementById("trendChart").getContext("2d");
  if (trendChartInstance) trendChartInstance.destroy();
  trendChartInstance = new Chart(ctx, {{
    type: "bar",
    data: {{
      labels: years,
      datasets: [
        {{
          type: "line",
          label: "연도별 ROE (%)",
          data: s.roe_h,
          borderColor: "#38bdf8",
          backgroundColor: "#38bdf8",
          yAxisID: "y1",
          tension: 0.25,
          borderWidth: 2.5
        }},
        {{
          type: "bar",
          label: "자본총계 (억원)",
          data: s.eq_h,
          backgroundColor: "rgba(34, 197, 94, 0.35)",
          borderColor: "#22c55e",
          borderWidth: 1,
          yAxisID: "y"
        }}
      ]
    }},
    options: {{
      responsive: true,
      interaction: {{ mode: "index", intersect: false }},
      plugins: {{ legend: {{ labels: {{ color: "#cbd5e1", font: {{ size: 11 }} }} }} }},
      scales: {{
        x: {{ ticks: {{ color: "#94a3b8" }}, grid: {{ color: "#1e293b" }} }},
        y: {{ position: "left", title: {{ display: true, text: "자본 (억원)", color: "#94a3b8" }}, ticks: {{ color: "#94a3b8" }}, grid: {{ color: "#334155" }} }},
        y1: {{ position: "right", title: {{ display: true, text: "ROE (%)", color: "#38bdf8" }}, ticks: {{ color: "#38bdf8" }}, grid: {{ drawOnChartArea: false }} }}
      }}
    }}
  }});
}}

function renderScatter(list) {{
  const ctx = document.getElementById("scatterChart").getContext("2d");
  if (scatterChartInstance) scatterChartInstance.destroy();
  const points = list.map(s => ({{ x: s.appRoe, y: s.pbr, name: s.name, code: s.code }}));
  scatterChartInstance = new Chart(ctx, {{
    type: "scatter",
    data: {{
      datasets: [
        {{
          label: "스캔 종목",
          data: points,
          backgroundColor: "rgba(56, 189, 248, 0.75)",
          pointRadius: 5
        }},
        {{
          type: "line",
          label: "청사진 매트릭스 기준선 (ROE 10%=PBR 1.0)",
          data: [{{x: 5, y: 0.5}}, {{x: 10, y: 1.0}}, {{x: 15, y: 1.5}}, {{x: 20, y: 2.0}}, {{x: 25, y: 2.5}}],
          borderColor: "#f59e0b",
          borderDash: [5, 5],
          borderWidth: 2,
          pointRadius: 0
        }}
      ]
    }},
    options: {{
      responsive: true,
      plugins: {{
        legend: {{ labels: {{ color: "#cbd5e1", font: {{ size: 11 }} }} }},
        tooltip: {{
          callbacks: {{
            label: function(ctx) {{
              const p = ctx.raw;
              return p.name ? `${{p.name}}: ROE ${{p.x.toFixed(1)}}%, PBR ${{p.y.toFixed(2)}}배` : "";
            }}
          }}
        }}
      }},
      scales: {{
        x: {{ title: {{ display: true, text: "적용 ROE (%)", color: "#94a3b8" }}, ticks: {{ color: "#94a3b8" }}, grid: {{ color: "#334155" }} }},
        y: {{ title: {{ display: true, text: "현재 PBR (배)", color: "#94a3b8" }}, ticks: {{ color: "#94a3b8" }}, grid: {{ color: "#334155" }} }}
      }}
    }}
  }});
}}

window.onload = updateDashboard;
</script>
</body>
</html>
"""

components.html(html_code, height=820, scrolling=True)
