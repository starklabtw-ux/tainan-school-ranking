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

st.set_page_config(
    page_title="台南市公私立國中第一志願錄取數據查詢系統",
    page_icon="📊",
    initial_sidebar_state="collapsed"
)

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
        with urllib.request.urlopen(req, timeout=3) as response:
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
        max-width: 800px;
        margin: 0 auto;
        padding: 1rem;
    }

    .main-title {
        text-align: center;
        font-size: 2rem;
        font-weight: bold;
        color: #1B3A5C;
        margin-bottom: 0.4rem;
        padding-bottom: 0.8rem;
        border-bottom: 2px solid #1B3A5C;
        letter-spacing: 0.02em;
    }

    .section-header {
        font-size: 1.35rem;
        font-weight: 700;
        color: #1B3A5C;
        margin-top: 2rem;
        margin-bottom: 1rem;
        padding: 0.4rem 0 0.4rem 0.75rem;
        border-left: 4px solid #1B3A5C;
        background-color: transparent;
        border-radius: 0;
        text-align: left;
    }

    .subsection-header {
        font-size: 1.1rem;
        font-weight: 600;
        color: #1B3A5C;
        margin-top: 1.2rem;
        margin-bottom: 0.6rem;
        padding-left: 0.5rem;
        border-left: 3px solid #8B9DB0;
    }

    .stDataFrame {
        border: 1px solid #D5DCE5;
        border-radius: 3px;
    }

    .stDataFrame table {
        text-align: center !important;
    }

    .stDataFrame th, .stDataFrame td {
        text-align: center !important;
        vertical-align: middle !important;
    }

    .stDataFrame div[data-testid="column"] {
        text-align: center !important;
    }

    div[data-testid="stDataFrame"] table th,
    div[data-testid="stDataFrame"] table td {
        text-align: center !important;
    }

    .stDataFrame .data {
        text-align: center !important;
    }

    .footer-text {
        text-align: center;
        font-style: italic;
        margin-top: 3rem;
        padding-top: 2rem;
        border-top: 1px solid #D5DCE5;
        color: #6B7280;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">台南市公私立國中第一志願錄取數據查詢系統</div>', unsafe_allow_html=True)
st.markdown(f'<div style="text-align: center; font-size: 10pt; margin-bottom: 2rem; color: #6B7280;">瀏覽人次：{st.session_state.visitor_count:,} &nbsp;&nbsp;&nbsp; 系統製作：<a href="https://www.facebook.com/starklabtw" target="_blank" style="color: #1B3A5C; text-decoration: none;">小史塔克實驗室</a></div>', unsafe_allow_html=True)

TABLE_STYLE = """
<style>
.academic-table {{
    margin: 0 auto;
    border-collapse: collapse;
    width: 100%;
    border-top: 3px double #1B3A5C;
    border-bottom: 3px double #1B3A5C;
    font-size: 10pt;
}}
.academic-table th {{
    background-color: #1B3A5C;
    color: #ffffff;
    font-weight: 700;
    text-align: center !important;
    padding: 8px 10px;
    border: none;
    border-right: 1px solid #2E5A8C;
}}
.academic-table th:last-child {{
    border-right: none;
}}
.academic-table td {{
    text-align: center !important;
    padding: 7px 10px;
    border: none;
    border-right: 1px solid #E8ECF0;
}}
.academic-table td:last-child {{
    border-right: none;
}}
.academic-table tr:nth-child(even) td {{
    background-color: #F7F9FC;
}}
.academic-table tr:nth-child(odd) td {{
    background-color: #ffffff;
}}
</style>
"""

def render_academic_table(df: pd.DataFrame, table_id: str) -> str:
    """Render a DataFrame as an academic-styled HTML table"""
    html = df.to_html(index=False, escape=False, table_id=table_id, classes="academic-table")
    return TABLE_STYLE.format() + html

# Main content area
if 'cached_df' not in st.session_state or st.session_state.cached_df is None:
    with st.spinner("正在載入學校資料..."):
        st.session_state.cached_df = data_manager.get_data()
df = st.session_state.cached_df

if df is not None and not df.empty:
    usage_info = """
    <div style="
        background-color: #F7F9FC;
        border: 1px solid #1B3A5C;
        padding: 20px;
        margin: 20px 0;
        border-radius: 2px;
        font-size: 12pt;
        line-height: 1.6;
    ">
        <h3 style="text-align: center; margin-top: 0; margin-bottom: 16px; color: #1B3A5C; font-size: 14pt; font-weight: 700;">系統使用說明</h3>
        <ol style="margin: 0; padding-left: 20px; color: #2C3E50;">
            <li style="margin-bottom: 10px;">本系統為免費查詢，所有數據僅提供家長作為挑選學校的參考資料，相關結論請使用者自行判斷</li>
            <li style="margin-bottom: 10px;">本頁數據僅採計免試錄取人數、南一中科學班錄取人數，錄取率的計算方式為 ( 該年免試錄取人數 + 南一中科學班錄取人數 ) / 該校當年三年級學生人數</li>
            <li style="margin-bottom: 10px;">若該生為達到第一志願分數卻沒有去南一中、南女，或者是錄取外縣市科學班皆不會採計進資料</li>
            <li style="margin-bottom: 0;">私校的相關數據需要考慮直升人數，並且有些並沒有向台南市學籍系統回報學生人數，本系統即無法呈現資料</li>
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

        st.markdown('<div class="subsection-header">五校錄取率歷年變化</div>', unsafe_allow_html=True)
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
        
        st.markdown('<div class="subsection-header">五校免試人數+科學班人數歷年變化</div>', unsafe_allow_html=True)
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
    
    st.markdown("---")
    
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
st.markdown("---")

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

st.markdown("---")

st.markdown(f'''
<div style="display: flex; justify-content: center; align-items: center; margin: 20px 0;">
    <img src="data:image/png;base64,{_get_base64_image('attached_assets/Logo_1752303088536.png')}" style="width: 100px;" />
</div>
''', unsafe_allow_html=True)

st.markdown('<div style="text-align: center; margin-top: 10px; font-size: 11pt; color: #6B7280;">資料來源：台南一中、台南女中、雪莉的數位生活、台南市國中學籍系統</div>', unsafe_allow_html=True)

st.markdown('<div style="text-align: center; margin-top: 10px; font-size: 11pt; color: #6B7280;">意見與勘誤回報：<a href="https://line.me/R/ti/p/%40starklab" target="_blank" style="color: #1B3A5C; text-decoration: none;">小史塔克實驗室</a></div>', unsafe_allow_html=True)

last_update = data_manager.get_last_update_time()
st.markdown(f'<div style="text-align: center; margin-top: 10px; font-size: 10pt; color: #6B7280;">最近系統資料更新時間：{last_update}</div>', unsafe_allow_html=True)
