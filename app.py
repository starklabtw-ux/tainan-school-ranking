import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import hashlib
import io
import os
import re
import json
import urllib.request
import urllib.error
from data_manager import DataManager
from chart_generator import ChartGenerator
import base64

def _get_base64_image(image_path):
    """Convert image to base64 string"""
    try:
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    except:
        return ""

def _get_gas_url():
    """Get GAS Web App URL from Streamlit secrets or environment variable"""
    try:
        url = st.secrets.get("gas", {}).get("web_app_url", None)
        if url:
            return url
    except Exception:
        pass
    return os.environ.get("GAS_WEB_APP_URL", None)

def _count_visit_via_gas(gas_url: str):
    """Call GAS countVisit endpoint and return the new count, or None on failure"""
    try:
        api_url = gas_url + "?action=countVisit"
        req = urllib.request.Request(api_url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=8) as response:
            result = json.loads(response.read().decode("utf-8"))
        return result.get("count", None)
    except Exception as e:
        print(f"GAS countVisit error: {e}")
        return None

def _count_visit_local():
    """Increment local visitor count file and return the new count"""
    try:
        count = 3932
        if os.path.exists('data/visitor_count.txt'):
            with open('data/visitor_count.txt', 'r') as f:
                count = int(f.read().strip())
        count += 1
        os.makedirs('data', exist_ok=True)
        with open('data/visitor_count.txt', 'w') as f:
            f.write(str(count))
        return count
    except Exception:
        return 3932

# Initialize session state
if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False
if 'data_manager' not in st.session_state:
    st.session_state.data_manager = DataManager()
if 'visitor_count' not in st.session_state:
    try:
        with open('data/visitor_count.txt', 'r') as f:
            st.session_state.visitor_count = int(f.read().strip())
    except:
        st.session_state.visitor_count = 3932
if 'page_view_counted' not in st.session_state:
    st.session_state.page_view_counted = True
    gas_url = _get_gas_url()
    if gas_url:
        new_count = _count_visit_via_gas(gas_url)
        if new_count is not None:
            st.session_state.visitor_count = new_count
        else:
            st.session_state.visitor_count = _count_visit_local()
    else:
        st.session_state.visitor_count = _count_visit_local()

# Load initial data
data_manager = st.session_state.data_manager
chart_generator = ChartGenerator()

st.set_page_config(
    page_title="台南市公私立國中第一志願錄取數據查詢系統",
    page_icon="📊",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<meta name="description" content="台南市公私立國中第一志願錄取數據查詢系統 - 提供2020-2025學年度台南一中、台南女中錄取率統計，支援學校搜尋、多校比較功能，協助家長選校參考">
<meta name="keywords" content="台南市國中,第一志願,錄取率,台南一中,台南女中,教育統計,學校查詢">
<meta name="author" content="小史塔克實驗室">
<meta property="og:title" content="台南市公私立國中第一志願錄取數據查詢系統">
<meta property="og:description" content="提供台南市國中升學統計數據查詢，包含歷年錄取率、學校比較功能">
<meta property="og:type" content="website">
""", unsafe_allow_html=True)

st.markdown("""
<style>
    .main .block-container {
        max-width: 1100px;
        margin: 0 auto;
        padding: 1.2rem 2rem;
    }

    /* ── Hero Banner ── */
    .hero-banner {
        background: linear-gradient(135deg, #1B3A5C 0%, #2E6B8A 50%, #1ABC9C 100%);
        border-radius: 12px;
        padding: 2.2rem 1.5rem 1.6rem;
        margin-bottom: 1.8rem;
        text-align: center;
        box-shadow: 0 4px 20px rgba(27, 58, 92, 0.25);
        position: relative;
        overflow: hidden;
    }
    .hero-banner::before {
        content: '';
        position: absolute;
        top: -50%;
        right: -20%;
        width: 300px;
        height: 300px;
        background: radial-gradient(circle, rgba(255,255,255,0.08) 0%, transparent 70%);
        border-radius: 50%;
    }
    .hero-banner h1 {
        color: #ffffff;
        font-size: 2.4rem;
        font-weight: 800;
        margin: 0 0 0.6rem 0;
        letter-spacing: 0.04em;
        text-shadow: 0 2px 8px rgba(0,0,0,0.15);
        position: relative;
    }
    .hero-meta {
        color: rgba(255,255,255,0.85);
        font-size: 1.05rem;
        position: relative;
    }
    .hero-meta a {
        color: #A8EDEA;
        text-decoration: none;
        font-weight: 600;
    }
    .hero-meta a:hover { text-decoration: underline; }
    .hero-badge {
        display: inline-block;
        background: rgba(255,255,255,0.18);
        padding: 4px 16px;
        border-radius: 20px;
        font-size: 0.95rem;
        margin-right: 12px;
        backdrop-filter: blur(4px);
    }

    /* ── Section Headers ── */
    .section-header {
        font-size: 1.7rem;
        font-weight: 700;
        color: #1B3A5C;
        margin-top: 2rem;
        margin-bottom: 1rem;
        padding: 0.6rem 1rem;
        border-left: 4px solid transparent;
        border-image: linear-gradient(180deg, #1B3A5C, #1ABC9C) 1;
        background: linear-gradient(90deg, rgba(26,188,156,0.06), transparent);
        border-radius: 0 6px 6px 0;
    }

    .subsection-header {
        font-size: 1.35rem;
        font-weight: 600;
        color: #2C3E50;
        margin-top: 1.2rem;
        margin-bottom: 0.6rem;
        padding: 0.3rem 0 0.3rem 0.6rem;
        border-left: 3px solid #1ABC9C;
    }

    /* ── Card Wrapper ── */
    .ui-card {
        background: #ffffff;
        border-radius: 10px;
        padding: 1.2rem 1.4rem;
        margin-bottom: 1rem;
        box-shadow: 0 1px 8px rgba(27, 58, 92, 0.08);
        border: 1px solid rgba(27, 58, 92, 0.06);
        transition: box-shadow 0.2s ease;
    }
    .ui-card:hover {
        box-shadow: 0 3px 16px rgba(27, 58, 92, 0.13);
    }

    /* ── Metric Cards ── */
    .metrics-row {
        display: flex;
        gap: 10px;
        margin: 1rem 0;
        flex-wrap: wrap;
    }
    .metric-card {
        flex: 1;
        min-width: 120px;
        background: #ffffff;
        border-radius: 10px;
        padding: 0.9rem 0.7rem;
        text-align: center;
        box-shadow: 0 1px 6px rgba(27, 58, 92, 0.08);
        border: 1px solid rgba(27, 58, 92, 0.06);
        transition: transform 0.15s ease, box-shadow 0.15s ease;
    }
    .metric-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 14px rgba(27, 58, 92, 0.14);
    }
    .metric-card .school-name {
        font-size: 1.05rem;
        font-weight: 700;
        color: #1B3A5C;
        margin-bottom: 4px;
    }
    .metric-card .metric-value {
        font-size: 1.75rem;
        font-weight: 800;
        background: linear-gradient(135deg, #1B3A5C, #1ABC9C);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
    }
    .metric-card .metric-trend {
        font-size: 0.9rem;
        margin-top: 2px;
    }
    .trend-up { color: #27AE60; }
    .trend-down { color: #E74C3C; }
    .trend-flat { color: #95A5A6; }

    /* ── Info Box ── */
    .info-box {
        background: linear-gradient(135deg, #F8FFFE, #F0F7FF);
        border-radius: 10px;
        padding: 1.2rem 1.4rem;
        margin: 1.2rem 0;
        border-left: 4px solid;
        border-image: linear-gradient(180deg, #1B3A5C, #1ABC9C) 1;
        box-shadow: 0 1px 6px rgba(27, 58, 92, 0.06);
    }
    .info-box h3 {
        text-align: center;
        margin-top: 0;
        margin-bottom: 12px;
        color: #1B3A5C;
        font-size: 1.35rem;
        font-weight: 700;
    }
    .info-box ol {
        margin: 0;
        padding-left: 20px;
        color: #374151;
        font-size: 1.1rem;
        line-height: 1.7;
    }
    .info-box li { margin-bottom: 6px; }

    /* ── Tables ── */
    .stDataFrame {
        border: 1px solid rgba(27, 58, 92, 0.1);
        border-radius: 8px;
        overflow: hidden;
    }
    .stDataFrame table { text-align: center !important; }
    .stDataFrame th, .stDataFrame td {
        text-align: center !important;
        vertical-align: middle !important;
    }
    div[data-testid="stDataFrame"] table th,
    div[data-testid="stDataFrame"] table td {
        text-align: center !important;
    }

    /* ── Footer ── */
    .site-footer {
        background: linear-gradient(135deg, #F8FFFE, #F0F4F8);
        border-radius: 12px;
        padding: 1.5rem;
        margin-top: 2rem;
        text-align: center;
        border-top: 3px solid;
        border-image: linear-gradient(90deg, #1B3A5C, #1ABC9C) 1;
    }
    .site-footer p {
        margin: 8px 0;
        font-size: 1.0rem;
        color: #6B7280;
    }
    .site-footer a {
        color: #1B3A5C;
        text-decoration: none;
        font-weight: 600;
    }
    .site-footer a:hover { color: #1ABC9C; }

    /* ── Divider ── */
    .gradient-divider {
        height: 2px;
        background: linear-gradient(90deg, transparent, #1ABC9C, #1B3A5C, #1ABC9C, transparent);
        margin: 2rem 0;
        border: none;
        border-radius: 1px;
    }

    /* ── Streamlit overrides ── */
    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #1B3A5C, #1ABC9C);
        border: none;
        border-radius: 8px;
        font-weight: 600;
        font-size: 1.05rem;
    }
    .stButton > button[kind="primary"]:hover {
        background: linear-gradient(135deg, #15304D, #17A68A);
    }
    hr { display: none; }

    /* ── Responsive ── */
    @media (max-width: 768px) {
        .main .block-container {
            max-width: 100%;
            padding: 0.8rem 1rem;
        }
        .hero-banner { padding: 1.6rem 1rem 1.2rem; }
        .hero-banner h1 { font-size: 1.6rem; }
        .hero-meta { font-size: 0.9rem; }
        .hero-badge { font-size: 0.82rem; }
        .section-header { font-size: 1.3rem; }
        .subsection-header { font-size: 1.1rem; }
        .info-box h3 { font-size: 1.15rem; }
        .info-box ol { font-size: 0.95rem; }
        .metrics-row { gap: 6px; }
        .metric-card {
            min-width: 90px;
            padding: 0.7rem 0.4rem;
        }
        .metric-card .school-name { font-size: 0.85rem; }
        .metric-card .metric-value { font-size: 1.3rem; }
        .metric-card .metric-trend { font-size: 0.75rem; }
        .academic-table {{ font-size: 0.9rem; }}
        .academic-table th {{ font-size: 0.85rem; padding: 8px 6px; }}
        .academic-table td {{ font-size: 0.88rem; padding: 7px 6px; }}
        .site-footer p { font-size: 0.88rem; }
    }

    @media (max-width: 480px) {
        .hero-banner h1 { font-size: 1.3rem; }
        .metrics-row {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 6px;
        }
        .metric-card { min-width: unset; }
    }
</style>
""", unsafe_allow_html=True)

st.markdown(f'''
<div class="hero-banner">
    <h1>台南市公私立國中<br>第一志願錄取數據查詢系統</h1>
    <div class="hero-meta">
        <span class="hero-badge">瀏覽人次 {st.session_state.visitor_count:,}</span>
        系統製作：<a href="https://www.facebook.com/starklabtw" target="_blank">小史塔克實驗室</a>
    </div>
</div>
''', unsafe_allow_html=True)

TABLE_STYLE = """
<style>
.academic-table {{
    margin: 0 auto;
    border-collapse: collapse;
    width: 100%;
    border-radius: 8px;
    overflow: hidden;
    font-size: 1.1rem;
    box-shadow: 0 1px 6px rgba(27, 58, 92, 0.08);
}}
.academic-table th {{
    background: linear-gradient(135deg, #1B3A5C 0%, #2E6B8A 100%);
    color: #ffffff;
    font-weight: 600;
    text-align: center !important;
    padding: 12px 12px;
    border: none;
    border-right: 1px solid rgba(255,255,255,0.15);
    font-size: 1.05rem;
    letter-spacing: 0.02em;
}}
.academic-table th:last-child {{
    border-right: none;
}}
.academic-table td {{
    text-align: center !important;
    padding: 11px 12px;
    border: none;
    border-right: 1px solid #EEF2F7;
    color: #374151;
    font-size: 1.05rem;
}}
.academic-table td:last-child {{
    border-right: none;
}}
.academic-table tr:nth-child(even) td {{
    background-color: #F8FFFE;
}}
.academic-table tr:nth-child(odd) td {{
    background-color: #ffffff;
}}
.academic-table tr:hover td {{
    background-color: rgba(26, 188, 156, 0.08);
    transition: background-color 0.15s ease;
}}
.academic-table td:first-child {{
    font-weight: 600;
    color: #1B3A5C;
}}
</style>
"""

def render_academic_table(df: pd.DataFrame, table_id: str) -> str:
    """Render a DataFrame as an academic-styled HTML table"""
    html = df.to_html(index=False, escape=False, table_id=table_id, classes="academic-table")
    return TABLE_STYLE.format() + html

# Main content area - cache data in session_state to avoid re-fetching GAS on every rerun
if 'df_cache' not in st.session_state:
    st.session_state.df_cache = data_manager.get_data()
df = st.session_state.df_cache

if df is not None and not df.empty:
    usage_info = """
    <div class="info-box">
        <h3>系統使用說明</h3>
        <ol>
            <li>本系統為免費查詢，所有數據僅提供家長作為挑選學校的參考資料，相關結論請使用者自行判斷</li>
            <li>本頁數據僅採計免試錄取人數、南一中科學班錄取人數，錄取率的計算方式為 ( 該年免試錄取人數 + 南一中科學班錄取人數 ) / 該校當年三年級學生人數</li>
            <li>若該生為達到第一志願分數卻沒有去南一中、南女，或者是錄取外縣市科學班皆不會採計進資料</li>
            <li>私校的相關數據需要考慮直升人數，並且有些並沒有向台南市學籍系統回報學生人數，本系統即無法呈現資料</li>
        </ol>
    </div>
    """
    
    st.markdown(usage_info, unsafe_allow_html=True)

    # Fixed display section for 5 key schools
    st.markdown('<div class="section-header">重點五校統計</div>', unsafe_allow_html=True)
    
    key_schools = ['建興國中', '復興國中', '後甲國中', '崇明國中', '民德國中']
    key_schools_data = df[df['學校'].isin(key_schools)].copy()
    
    if not key_schools_data.empty:
        year_columns = data_manager.get_year_columns(df)

        # Metric cards for latest year — find the latest year with data
        display_year = None
        prev_year_label = None
        for check_year in ['114', '113', '112', '111', '110', '109']:
            col_name = f"{check_year}學年第一志願錄取率"
            if col_name in key_schools_data.columns:
                valid = key_schools_data[col_name].apply(
                    lambda v: pd.notna(v) and str(v) not in ['-', '#VALUE!']
                )
                if valid.any():
                    if display_year is None:
                        display_year = check_year
                    elif prev_year_label is None:
                        prev_year_label = check_year
                        break

        metric_cards_html = f'<div style="text-align:center; color:#6B7280; font-size:1.0rem; margin-bottom:4px;">{display_year} 學年度第一志願錄取率（較 {prev_year_label} 學年度）</div>'
        metric_cards_html += '<div class="metrics-row">'
        for _, school_row in key_schools_data.iterrows():
            school_name = school_row['學校']
            latest_rate = None
            prev_rate = None
            col_latest = f"{display_year}學年第一志願錄取率"
            col_prev = f"{prev_year_label}學年第一志願錄取率" if prev_year_label else None
            if col_latest in school_row and pd.notna(school_row[col_latest]) and str(school_row[col_latest]) not in ['-', '#VALUE!']:
                try:
                    latest_rate = float(school_row[col_latest]) * 100
                except:
                    pass
            if col_prev and col_prev in school_row and pd.notna(school_row[col_prev]) and str(school_row[col_prev]) not in ['-', '#VALUE!']:
                try:
                    prev_rate = float(school_row[col_prev]) * 100
                except:
                    pass
            if latest_rate is not None:
                rate_str = f"{latest_rate:.1f}%"
                if prev_rate is not None:
                    diff = latest_rate - prev_rate
                    if diff > 0.5:
                        trend_html = f'<div class="metric-trend trend-up">&#9650; +{diff:.1f}%</div>'
                    elif diff < -0.5:
                        trend_html = f'<div class="metric-trend trend-down">&#9660; {diff:.1f}%</div>'
                    else:
                        trend_html = '<div class="metric-trend trend-flat">&#9644; 持平</div>'
                else:
                    trend_html = '<div class="metric-trend trend-flat">—</div>'
                metric_cards_html += f'''
                <div class="metric-card">
                    <div class="school-name">{school_name}</div>
                    <div class="metric-value">{rate_str}</div>
                    {trend_html}
                </div>'''
        metric_cards_html += '</div>'
        st.markdown(metric_cards_html, unsafe_allow_html=True)

        st.markdown('<div class="subsection-header">五校<span style="background:linear-gradient(135deg,#1B3A5C,#1ABC9C);color:#fff;padding:2px 8px;border-radius:4px;font-size:0.85em;">錄取率</span>歷年變化（含科學班）</div>', unsafe_allow_html=True)
        admission_rate_chart = chart_generator.create_admission_rate_comparison(key_schools_data)
        if admission_rate_chart:
            st.plotly_chart(admission_rate_chart, use_container_width=True)
        
        admission_rate_table_data = []
        for _, school_row in key_schools_data.iterrows():
            row_data = {'學校': school_row['學校']}
            for year in ['109', '110', '111', '112', '113', '114']:
                col_name = f"{year}學年第一志願錄取率"
                if col_name in school_row and pd.notna(school_row[col_name]) and str(school_row[col_name]) != '-':
                    try:
                        rate_value = float(school_row[col_name]) * 100
                        row_data[f"{year}年"] = f"{rate_value:.2f}%"
                    except:
                        row_data[f"{year}年"] = '-'
                else:
                    row_data[f"{year}年"] = '-'
            admission_rate_table_data.append(row_data)
        
        if admission_rate_table_data:
            admission_rate_df = pd.DataFrame(admission_rate_table_data)
            st.html(render_academic_table(admission_rate_df, "rate-table"))
        
        st.markdown('<div class="subsection-header">五校<span style="background:linear-gradient(135deg,#E74C3C,#F39C12);color:#fff;padding:2px 8px;border-radius:4px;font-size:0.85em;">錄取人數</span>歷年變化（含科學班）</div>', unsafe_allow_html=True)
        student_count_chart = chart_generator.create_student_count_comparison(key_schools_data)
        if student_count_chart:
            st.plotly_chart(student_count_chart, use_container_width=True)
        
        admission_count_table_data = []
        for _, school_row in key_schools_data.iterrows():
            row_data = {'學校': school_row['學校']}
            for year in ['109', '110', '111', '112', '113', '114']:
                col_name = f"{year}學年免試人數"
                if col_name in school_row and pd.notna(school_row[col_name]) and str(school_row[col_name]) != '-':
                    row_data[f"{year}年"] = str(int(float(school_row[col_name])))
                else:
                    row_data[f"{year}年"] = '-'
            admission_count_table_data.append(row_data)
        
        if admission_count_table_data:
            st.markdown('<div class="subsection-header">免試錄取人數</div>', unsafe_allow_html=True)
            admission_count_df = pd.DataFrame(admission_count_table_data)
            st.html(render_academic_table(admission_count_df, "count-table"))
        
        science_class_table_data = []
        for _, school_row in key_schools_data.iterrows():
            row_data = {'學校': school_row['學校']}
            for year in ['109', '110', '111', '112', '113', '114']:
                col_name = f"{year}學年考取科學班人數"
                if col_name in school_row and pd.notna(school_row[col_name]) and str(school_row[col_name]) != '-':
                    row_data[f"{year}年"] = str(int(float(school_row[col_name])))
                else:
                    row_data[f"{year}年"] = '-'
            science_class_table_data.append(row_data)
        
        if science_class_table_data:
            st.markdown('<div class="subsection-header">南一中科學班錄取人數</div>', unsafe_allow_html=True)
            science_class_df = pd.DataFrame(science_class_table_data)
            st.html(render_academic_table(science_class_df, "science-table"))
    
    st.markdown('<div class="gradient-divider"></div>', unsafe_allow_html=True)

    st.markdown('<div class="section-header">您想查詢哪一所學校的資料呢？</div>', unsafe_allow_html=True)
    
    if 'comparison_schools' not in st.session_state:
        st.session_state.comparison_schools = []
    
    col_mode1, col_mode2 = st.columns(2)
    with col_mode1:
        query_mode = st.radio("查詢模式", ["單校查詢", "多校查詢比較"], horizontal=True)
    
    if query_mode == "單校查詢":
        search_col1, search_col2 = st.columns([3, 1])
        with search_col1:
            search_term = st.text_input("快速搜尋學校（支援模糊搜尋）", placeholder="例如：建興、復興、後甲...", key="single_search")
        with search_col2:
            search_clicked = st.button("搜尋", key="search_single")
        
        if 'single_search_results' not in st.session_state:
            st.session_state.single_search_results = None
        if 'single_search_term' not in st.session_state:
            st.session_state.single_search_term = ""
        
        if search_clicked or (search_term and search_term != st.session_state.single_search_term):
            if search_term:
                search_results = data_manager.search_schools(df, search_term)
                st.session_state.single_search_results = search_results
                st.session_state.single_search_term = search_term
            else:
                st.session_state.single_search_results = None
                st.session_state.single_search_term = ""
        
        base_df = df.copy()
        if st.session_state.single_search_results is not None and not st.session_state.single_search_results.empty:
            base_df = st.session_state.single_search_results
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            school_types = ['全部', '不限'] + sorted(base_df['公立/私立'].dropna().unique().tolist())
            selected_type = st.selectbox("公立/私立", school_types)
        
        filtered_df = base_df.copy()
        if selected_type != '全部' and selected_type != '不限':
            filtered_df = filtered_df[filtered_df['公立/私立'] == selected_type]
        
        with col2:
            unique_districts = filtered_df['台南市行政區'].dropna().unique().tolist()
            sorted_districts = sorted(unique_districts, key=lambda x: (len(x), x))
            districts = ['全部'] + sorted_districts
            selected_district = st.selectbox("台南市行政區", districts)
        
        if selected_district != '全部':
            filtered_df = filtered_df[filtered_df['台南市行政區'] == selected_district]
        
        with col3:
            schools = ['全部'] + sorted(filtered_df['學校'].dropna().unique().tolist())
            selected_school = st.selectbox("學校名稱", schools)
            
        if selected_school != '全部':
            filtered_df = filtered_df[filtered_df['學校'] == selected_school]
        
        if st.session_state.single_search_results is not None:
            if not st.session_state.single_search_results.empty:
                st.success(f"找到 {len(st.session_state.single_search_results)} 所學校包含「{st.session_state.single_search_term}」")
                
                search_display_df = st.session_state.single_search_results[['學校', '公立/私立', '台南市行政區']].copy()
                st.dataframe(search_display_df, use_container_width=True, hide_index=True)
                
                if len(st.session_state.single_search_results) == 1:
                    school_data = st.session_state.single_search_results.iloc[0]
                    school_name = school_data['學校']
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        st.markdown(f'<div class="subsection-header">{school_name} 歷年錄取率</div>', unsafe_allow_html=True)
                        single_school_rate_chart = chart_generator.create_single_school_admission_rate(school_data)
                        if single_school_rate_chart:
                            st.plotly_chart(single_school_rate_chart, use_container_width=True)
                    
                    with col2:
                        st.markdown(f'<div class="subsection-header">{school_name} 歷年第一志願錄取人數</div>', unsafe_allow_html=True)
                        single_school_count_chart = chart_generator.create_single_school_student_count(school_data)
                        if single_school_count_chart:
                            st.plotly_chart(single_school_count_chart, use_container_width=True)
                    
                    st.markdown(f'<div class="subsection-header">{school_name} 詳細資料</div>', unsafe_allow_html=True)
                    
                    table_data = []
                    
                    row_admission = {'學校': '免試錄取人數'}
                    for year in ['109', '110', '111', '112', '113', '114']:
                        col_name = f"{year}學年免試人數"
                        if col_name in school_data and pd.notna(school_data[col_name]) and str(school_data[col_name]) != '-':
                            row_admission[f"{year}年"] = str(int(float(school_data[col_name])))
                        else:
                            row_admission[f"{year}年"] = '-'
                    table_data.append(row_admission)
                    
                    row_science = {'學校': '科學班錄取人數'}
                    for year in ['109', '110', '111', '112', '113', '114']:
                        col_name = f"{year}學年考取科學班人數"
                        if col_name in school_data and pd.notna(school_data[col_name]) and str(school_data[col_name]) != '-':
                            row_science[f"{year}年"] = str(int(float(school_data[col_name])))
                        else:
                            row_science[f"{year}年"] = '-'
                    table_data.append(row_science)
                    
                    row_rate = {'學校': '第一志願錄取率'}
                    for year in ['109', '110', '111', '112', '113', '114']:
                        col_name = f"{year}學年第一志願錄取率"
                        if col_name in school_data and pd.notna(school_data[col_name]) and str(school_data[col_name]) != '-' and str(school_data[col_name]) != '#VALUE!':
                            try:
                                rate_value = float(school_data[col_name]) * 100
                                row_rate[f"{year}年"] = f"{rate_value:.2f}%"
                            except:
                                row_rate[f"{year}年"] = '-'
                        else:
                            row_rate[f"{year}年"] = '-'
                    table_data.append(row_rate)
                    
                    result_df = pd.DataFrame(table_data)
                    st.html(render_academic_table(result_df, "search-result-table"))
            else:
                st.warning(f"找不到包含「{st.session_state.single_search_term}」的學校")
    
    else:
        st.markdown("### 多校查詢比較 (最多可選擇5所學校)")
        
        search_col1, search_col2 = st.columns([3, 1])
        with search_col1:
            multi_search_term = st.text_input("快速搜尋學校（支援模糊搜尋）", placeholder="例如：建興、復興、後甲...", key="multi_search")
        with search_col2:
            multi_search_clicked = st.button("搜尋", key="search_multi")
        
        if 'multi_search_results' not in st.session_state:
            st.session_state.multi_search_results = None
        if 'multi_search_term' not in st.session_state:
            st.session_state.multi_search_term = ""
        
        if multi_search_clicked or (multi_search_term and multi_search_term != st.session_state.multi_search_term):
            if multi_search_term:
                multi_search_results = data_manager.search_schools(df, multi_search_term)
                st.session_state.multi_search_results = multi_search_results
                st.session_state.multi_search_term = multi_search_term
            else:
                st.session_state.multi_search_results = None
                st.session_state.multi_search_term = ""
        
        filter_col1, filter_col2, filter_col3 = st.columns(3)
        
        with filter_col1:
            multi_school_types = ['全部', '不限'] + sorted(df['公立/私立'].dropna().unique().tolist())
            multi_selected_type = st.selectbox("公立/私立", multi_school_types, key="multi_type")
        
        multi_filtered_df = df.copy()
        if multi_selected_type != '全部' and multi_selected_type != '不限':
            multi_filtered_df = multi_filtered_df[multi_filtered_df['公立/私立'] == multi_selected_type]
        
        with filter_col2:
            multi_unique_districts = multi_filtered_df['台南市行政區'].dropna().unique().tolist()
            multi_sorted_districts = sorted(multi_unique_districts, key=lambda x: (len(x), x))
            multi_districts = ['全部'] + multi_sorted_districts
            multi_selected_district = st.selectbox("台南市行政區", multi_districts, key="multi_district")
        
        if multi_selected_district != '全部':
            multi_filtered_df = multi_filtered_df[multi_filtered_df['台南市行政區'] == multi_selected_district]
        
        with filter_col3:
            multi_schools = ['全部'] + sorted(multi_filtered_df['學校'].dropna().unique().tolist())
            multi_selected_school = st.selectbox("學校名稱", multi_schools, key="multi_school")
        
        if multi_selected_school != '全部':
            col_filter_info, col_filter_add = st.columns([4, 1])
            with col_filter_info:
                selected_school_info = multi_filtered_df[multi_filtered_df['學校'] == multi_selected_school].iloc[0]
                st.write(f"已選擇：**{multi_selected_school}** ({selected_school_info['公立/私立']}, {selected_school_info['台南市行政區']})")
            with col_filter_add:
                if multi_selected_school not in st.session_state.comparison_schools:
                    if len(st.session_state.comparison_schools) < 5:
                        if st.button("加入比較", key="add_filter_school"):
                            st.session_state.comparison_schools.append(multi_selected_school)
                            st.rerun()
                    else:
                        st.write("清單已滿")
                else:
                    st.write("已加入")
        
        if st.session_state.multi_search_results is not None:
            if not st.session_state.multi_search_results.empty:
                st.success(f"找到 {len(st.session_state.multi_search_results)} 所學校包含「{st.session_state.multi_search_term}」")
                
                for _, school_row in st.session_state.multi_search_results.iterrows():
                    school_name = school_row['學校']
                    school_type = school_row['公立/私立']
                    school_district = school_row['台南市行政區']
                    
                    col_info, col_add = st.columns([4, 1])
                    with col_info:
                        st.write(f"**{school_name}** ({school_type}, {school_district})")
                    with col_add:
                        if school_name not in st.session_state.comparison_schools:
                            if len(st.session_state.comparison_schools) < 5:
                                if st.button("加入比較", key=f"add_search_{school_name}"):
                                    st.session_state.comparison_schools.append(school_name)
                                    st.session_state.multi_search_results = None
                                    st.session_state.multi_search_term = ""
                                    st.rerun()
                            else:
                                st.write("已滿")
                        else:
                            st.write("已加入")
            else:
                st.warning(f"找不到包含「{st.session_state.multi_search_term}」的學校")
        
        st.markdown("---")
        
        st.markdown("### 比較清單")
        
        if st.session_state.comparison_schools:
            for i, school in enumerate(st.session_state.comparison_schools):
                col_school, col_remove = st.columns([4, 1])
                with col_school:
                    st.write(f"**{i+1}.** {school}")
                with col_remove:
                    if st.button("移除", key=f"remove_{i}", help="移除此學校"):
                        st.session_state.comparison_schools.pop(i)
                        st.session_state.show_comparison = False
                        st.rerun()
            
            col1, col2, col3 = st.columns([2, 2, 1])
            
            with col1:
                if st.button("開始比較", type="primary", use_container_width=True):
                    st.session_state.show_comparison = True
                    st.rerun()
            
            with col2:
                if st.button("清除全部", use_container_width=True):
                    st.session_state.comparison_schools = []
                    st.session_state.show_comparison = False
                    st.rerun()
            
            remaining_slots = 5 - len(st.session_state.comparison_schools)
            if remaining_slots > 0:
                st.info(f"您還可以再加入 {remaining_slots} 所學校。可使用搜尋或篩選功能加入學校。")
            else:
                st.warning("已達到最大比較數量（5所學校）")
        else:
            st.info("使用說明：請使用搜尋功能或三層篩選選單來加入學校進行比較。最多可比較5所學校。")
        
        st.markdown("---")
        
        if 'show_comparison' not in st.session_state:
            st.session_state.show_comparison = False
        
        if st.session_state.comparison_schools and st.session_state.show_comparison:
            filtered_df = df[df['學校'].isin(st.session_state.comparison_schools)].copy()
            selected_school = None
        else:
            filtered_df = pd.DataFrame()
            selected_school = None
    
    if query_mode == "單校查詢" and selected_school != '全部':
        filtered_df = filtered_df[filtered_df['學校'] == selected_school]
    
    if query_mode == "單校查詢" and selected_school != '全部' and not filtered_df.empty:
        st.markdown('<div class="subsection-header">查詢結果</div>', unsafe_allow_html=True)
        
        school_data = filtered_df.iloc[0]
        school_name = school_data['學校']
        
        table_data = []
        
        row_admission = {'學校': '免試錄取人數'}
        for year in ['109', '110', '111', '112', '113', '114']:
            col_name = f"{year}學年免試人數"
            if col_name in school_data and pd.notna(school_data[col_name]) and str(school_data[col_name]) != '-':
                row_admission[f"{year}年"] = str(int(float(school_data[col_name])))
            else:
                row_admission[f"{year}年"] = '-'
        table_data.append(row_admission)
        
        row_science = {'學校': '科學班錄取人數'}
        for year in ['109', '110', '111', '112', '113', '114']:
            col_name = f"{year}學年考取科學班人數"
            if col_name in school_data and pd.notna(school_data[col_name]) and str(school_data[col_name]) != '-':
                row_science[f"{year}年"] = str(int(float(school_data[col_name])))
            else:
                row_science[f"{year}年"] = '-'
        table_data.append(row_science)
        
        row_rate = {'學校': '第一志願錄取率'}
        for year in ['109', '110', '111', '112', '113', '114']:
            col_name = f"{year}學年第一志願錄取率"
            if col_name in school_data and pd.notna(school_data[col_name]) and str(school_data[col_name]) != '-' and str(school_data[col_name]) != '#VALUE!':
                try:
                    rate_value = float(school_data[col_name]) * 100
                    row_rate[f"{year}年"] = f"{rate_value:.2f}%"
                except:
                    row_rate[f"{year}年"] = '-'
            else:
                row_rate[f"{year}年"] = '-'
        table_data.append(row_rate)
        
        result_df = pd.DataFrame(table_data)
        st.html(render_academic_table(result_df, "center-table"))
        
        if len(filtered_df) == 1:
            school_data = filtered_df.iloc[0]
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown(f'<div class="subsection-header">{selected_school} 歷年錄取率</div>', unsafe_allow_html=True)
                single_school_rate_chart = chart_generator.create_single_school_admission_rate(school_data)
                if single_school_rate_chart:
                    st.plotly_chart(single_school_rate_chart, use_container_width=True)
            
            with col2:
                st.markdown(f'<div class="subsection-header">{selected_school} 歷年第一志願錄取人數</div>', unsafe_allow_html=True)
                single_school_count_chart = chart_generator.create_single_school_student_count(school_data)
                if single_school_count_chart:
                    st.plotly_chart(single_school_count_chart, use_container_width=True)
    
    elif query_mode == "多校查詢比較" and st.session_state.comparison_schools and not filtered_df.empty:
        st.markdown('<div class="subsection-header">多校查詢比較結果</div>', unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown('<div class="subsection-header">錄取率比較</div>', unsafe_allow_html=True)
            comparison_rate_chart = chart_generator.create_admission_rate_comparison(filtered_df)
            if comparison_rate_chart:
                st.plotly_chart(comparison_rate_chart, use_container_width=True)
        
        with col2:
            st.markdown('<div class="subsection-header">錄取人數比較</div>', unsafe_allow_html=True)
            comparison_count_chart = chart_generator.create_student_count_comparison(filtered_df)
            if comparison_count_chart:
                st.plotly_chart(comparison_count_chart, use_container_width=True)
        
        st.markdown('<div class="subsection-header">詳細比較數據</div>', unsafe_allow_html=True)
        
        st.markdown("**歷年錄取率比較**")
        admission_rate_table_data = []
        for _, school_row in filtered_df.iterrows():
            row_data = {'學校': school_row['學校']}
            for year in ['109', '110', '111', '112', '113', '114']:
                col_name = f"{year}學年第一志願錄取率"
                if col_name in school_row and pd.notna(school_row[col_name]) and str(school_row[col_name]) != '-':
                    try:
                        rate_value = float(school_row[col_name]) * 100
                        row_data[f"{year}年"] = f"{rate_value:.2f}%"
                    except:
                        row_data[f"{year}年"] = '-'
                else:
                    row_data[f"{year}年"] = '-'
            admission_rate_table_data.append(row_data)
        
        if admission_rate_table_data:
            rate_df = pd.DataFrame(admission_rate_table_data)
            st.html(render_academic_table(rate_df, "comparison-rate-table"))
        
        st.markdown("**歷年免試錄取人數比較**")
        admission_count_table_data = []
        for _, school_row in filtered_df.iterrows():
            row_data = {'學校': school_row['學校']}
            for year in ['109', '110', '111', '112', '113', '114']:
                col_name = f"{year}學年免試人數"
                if col_name in school_row and pd.notna(school_row[col_name]) and str(school_row[col_name]) != '-':
                    row_data[f"{year}年"] = str(int(float(school_row[col_name])))
                else:
                    row_data[f"{year}年"] = '-'
            admission_count_table_data.append(row_data)
        
        if admission_count_table_data:
            count_df = pd.DataFrame(admission_count_table_data)
            st.html(render_academic_table(count_df, "comparison-count-table"))
    
    elif query_mode == "多校查詢" and not st.session_state.comparison_schools:
        st.info("請選擇要比較的學校")
    
    elif selected_school != '全部' and filtered_df.empty:
        st.info("沒有符合條件的學校資料")

else:
    st.error("無法載入學校資料，請檢查資料檔案是否存在")

# Footer
st.markdown('<div class="gradient-divider"></div>', unsafe_allow_html=True)

with st.expander("管理者專區"):
    if not st.session_state.authenticated:
        password = st.text_input("請輸入管理者密碼", type="password")
        if st.button("登入"):
            if password == "95960506":
                st.session_state.authenticated = True
                st.success("登入成功！")
                st.rerun()
            else:
                st.error("密碼錯誤！")
    else:
        st.success("已登入管理者模式")
        
        st.markdown('<div class="subsection-header">資料上傳</div>', unsafe_allow_html=True)
        uploaded_file = st.file_uploader("選擇CSV檔案", type=['csv'])
        
        if uploaded_file is not None:
            try:
                df_new = pd.read_csv(uploaded_file)
                
                validation_result = data_manager.validate_csv_structure(df_new)
                
                if validation_result['valid']:
                    st.success("檔案格式驗證通過")
                    
                    st.markdown('<div class="subsection-header">資料預覽</div>', unsafe_allow_html=True)
                    st.dataframe(df_new.head(10))
                    
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("學校總數", len(df_new))
                    with col2:
                        st.metric("公立學校", len(df_new[df_new['公立/私立'] == '公立']))
                    with col3:
                        st.metric("私立學校", len(df_new[df_new['公立/私立'] == '私立']))
                    
                    if st.button("確認更新資料", type="primary"):
                        success = data_manager.update_data(df_new)
                        if success:
                            st.success("資料更新成功！")
                            st.rerun()
                        else:
                            st.error("資料更新失敗！")
                else:
                    st.error(f"檔案格式驗證失敗: {validation_result['error']}")
                    
            except Exception as e:
                st.error(f"檔案讀取錯誤: {str(e)}")
        
        if st.button("登出"):
            st.session_state.authenticated = False
            st.rerun()

last_update = data_manager.get_last_update_time()
st.markdown(f'''
<div class="site-footer">
    <img src="data:image/png;base64,{_get_base64_image('attached_assets/Logo_1752303088536.png')}" style="width: 80px; margin-bottom: 8px;" />
    <p>資料來源：台南一中、台南女中、雪莉的數位生活、台南市國中學籍系統</p>
    <p>意見與勘誤回報：<a href="https://line.me/R/ti/p/%40starklab" target="_blank">小史塔克實驗室</a></p>
    <p style="font-size: 0.88rem; color: #9CA3AF;">最近系統資料更新時間：{last_update}</p>
</div>
''', unsafe_allow_html=True)
