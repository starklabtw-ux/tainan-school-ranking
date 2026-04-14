import pandas as pd
import os
import re
import json
import urllib.request
import urllib.error
from typing import Dict, List, Optional, Any
from datetime import datetime

class DataManager:
    def __init__(self):
        self.data_file = "data/schools_data.csv"
        self.backup_file = "data/schools_data_backup.csv"
        self._ensure_data_directory()
        self._initialize_data()
    
    def _get_gas_url(self) -> Optional[str]:
        """Get GAS Web App URL from Streamlit secrets or environment variable"""
        try:
            import streamlit as st
            url = st.secrets.get("gas", {}).get("web_app_url", None)
            if url:
                return url
        except Exception:
            pass
        return os.environ.get("GAS_WEB_APP_URL", None)
    
    def load_from_gas(self) -> Optional[pd.DataFrame]:
        """Load school data from GAS Web App API"""
        gas_url = self._get_gas_url()
        if not gas_url:
            return None
        
        try:
            api_url = gas_url + "?action=getData"
            req = urllib.request.Request(api_url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=10) as response:
                raw = response.read().decode("utf-8")
            
            result = json.loads(raw)
            
            if "error" in result:
                print(f"GAS API error: {result['error']}")
                return None
            
            columns = result.get("columns", [])
            data = result.get("data", [])
            
            if not columns or not data:
                return None
            
            df = pd.DataFrame(data, columns=columns)
            return df
            
        except urllib.error.URLError as e:
            print(f"GAS API network error: {e}")
            return None
        except Exception as e:
            print(f"GAS API unexpected error: {e}")
            return None
    
    def _ensure_data_directory(self):
        """Ensure data directory exists"""
        os.makedirs("data", exist_ok=True)
    
    def _initialize_data(self):
        """Initialize data from CSV file or create default"""
        if not os.path.exists(self.data_file):
            attached_csv_path = "attached_assets/20250712_1752290889615.csv"
            if os.path.exists(attached_csv_path):
                try:
                    initial_data = pd.read_csv(attached_csv_path, encoding='utf-8')
                    initial_data.to_csv(self.data_file, index=False, encoding='utf-8-sig')
                    return
                except Exception as e:
                    print(f"Error loading attached CSV: {e}")
            
            initial_data = self._create_initial_data()
            initial_data.to_csv(self.data_file, index=False, encoding='utf-8-sig')
    
    def _create_initial_data(self) -> pd.DataFrame:
        """Create initial data structure based on the provided CSV"""
        data = {
            '學校': ['崑山高中', '北門國中', '佳興國中', '大內國中', '東山國中', '光華國中', '菁寮國中', '黎明高中', '鹽行國中', '九份子國中(小)', '學甲國中', '太子國中', '柳營國中', '建興國中', '復興國中', '後甲國中', '民德國中', '崇明國中', '德光中學', '忠孝國中', '中山國中', '金城國中', '大橋國中', '永康國中', '歸仁國中', '南科實中', '安南國中', '港明高中', '安平國中', '大成國中', '新東國中', '瀛海高中', '海佃國中', '新興國中', '大灣高中', '慈濟高中', '文賢國中', '昭明國中', '長榮中學', '安順國中', '新化國中', '聖功女中', '佳里國中', '興國高中', '安定國中', '和順國中', '仁德國中', '麻豆國中', '沙崙國中', '南新國中', '關廟國中', '南寧高中', '南化國中', '土城高中', '善化國中', '白河國中', '東原國中', '新市國中', '延平國中', '仁德文賢國中', '成功國中', '官田國中', '玉井國中', '永仁高中', '山上國中', '將軍國中', '西港國中', '六甲國中', '龍崎國中', '左鎮國中', '光華高中', '下營國中', '後壁國中', '明達高中', '鹽水國中', '南光高中', '楠西國中'],
            '完整正式校名': ['臺南市私立崑山高級中等學校附設國中部', '臺南市立北門國民中學', '臺南市立佳興國民中學', '臺南市立大內國民中學', '臺南市立東山國民中學', '臺南市私立光華高級中學國中部', '臺南市立菁寮國民中學', '臺南市私立黎明高級中學附設國中部', '臺南市立鹽行國民中學', '臺南市立九份子國民中學', '臺南市立學甲國民中學', '臺南市立太子國民中學', '臺南市立柳營國民中學', '臺南市立建興國民中學', '臺南市立復興國民中學', '臺南市立後甲國民中學', '臺南市立民德國民中學', '臺南市立崇明國民中學', '臺南市私立德光高級中學附設國中部', '臺南市立忠孝國民中學', '臺南市立中山國民中學', '臺南市立金城國民中學', '臺南市立大橋國民中學', '臺南市立永康國民中學', '臺南市立歸仁國民中學', '國立南科國際實驗高級中學附設國中部', '臺南市立安南國民中學', '臺南市私立港明高級中學附設國中部', '臺南市立安平國民中學', '臺南市立大成國民中學', '臺南市立新東國民中學', '臺南市私立瀛海高級中學附設國中部', '臺南市立海佃國民中學', '臺南市立新興國民中學', '臺南市立大灣高級中學附設國中部', '臺南市私立慈濟高級中學附設國中部', '臺南市立文賢國民中學', '臺南市私立昭明國民中學', '臺南市私立長榮高級中學附設國中部', '臺南市立安順國民中學', '臺南市立新化國民中學', '臺南市私立聖功女子高級中學附設國中部', '臺南市立佳里國民中學', '臺南市私立興國高級中學附設國中部', '臺南市立安定國民中學', '臺南市立和順國民中學', '臺南市立仁德國民中學', '臺南市立麻豆國民中學', '臺南市立沙崙國民中學', '臺南市立南新國民中學', '臺南市立關廟國民中學', '臺南市立南寧高級中學附設國中部', '臺南市立南化國民中學', '臺南市立土城高級中學附設國中部', '臺南市立善化國民中學', '臺南市立白河國民中學', '臺南市立東原國民中學', '臺南市立新市國民中學', '臺南市立延平國民中學', '臺南市立仁德文賢國民中學', '臺南市立成功國民中學', '臺南市立官田國民中學', '臺南市立玉井國民中學', '臺南市立永仁高級中學附設國中部', '臺南市立山上國民中學', '臺南市立將軍國民中學', '臺南市立西港國民中學', '臺南市立六甲國民中學', '臺南市立龍崎國民中學', '臺南市立左鎮國民中學', '臺南市私立光華高級中學國中部', '臺南市立下營國民中學', '臺南市立後壁國民中學', '臺南市私立明達高級中學附設國中部', '臺南市立鹽水國民中學', '臺南市私立南光高級中學附設國中部', '臺南市立楠西國民中學'],
            '公立/私立': ['私立', '公立', '公立', '公立', '公立', '私立', '公立', '私立', '公立', '公立', '公立', '公立', '公立', '公立', '公立', '公立', '公立', '公立', '私立', '公立', '公立', '公立', '公立', '公立', '公立', '公立', '公立', '私立', '公立', '公立', '公立', '私立', '公立', '公立', '公立', '私立', '公立', '私立', '私立', '公立', '公立', '私立', '公立', '私立', '公立', '公立', '公立', '公立', '公立', '公立', '公立', '公立', '公立', '公立', '公立', '公立', '公立', '公立', '公立', '公立', '公立', '公立', '公立', '公立', '公立', '公立', '公立', '公立', '公立', '公立', '私立', '公立', '公立', '私立', '公立', '私立', '公立'],
            '台南市行政區': ['北區', '北門區', '佳里區', '大內區', '東山區', '東區', '後壁區', '麻豆區', '永康區', '安南區', '學甲區', '新營區', '柳營區', '中西區', '東區', '東區', '北區', '東區', '東區', '東區', '中西區', '安平區', '永康區', '永康區', '歸仁區', '新市區', '安南區', '西港區', '安平區', '南區', '新營區', '安南區', '安南區', '南區', '永康區', '安平區', '北區', '七股區', '東區', '安南區', '新化區', '北區', '佳里區', '新營區', '安定區', '安南區', '仁德區', '麻豆區', '歸仁區', '新營區', '關廟區', '南區', '南化區', '安南區', '善化區', '白河區', '東山區', '新市區', '北區', '仁德區', '北區', '官田區', '玉井區', '永康區', '山上區', '將軍區', '西港區', '六甲區', '龍崎區', '左鎮區', '東區', '下營區', '後壁區', '鹽水區', '鹽水區', '新營區', '楠西區'],
            '109學年免試人數': ['-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', 200, 177, 129, 79, 59, 45, 45, 37, 33, 28, 26, 24, 22, 22, 18, 15, 15, 15, 15, 15, 13, 11, 11, 11, 10, 10, 9, 8, 7, 6, 6, 6, 5, 5, 5, 5, 4, 4, 4, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 1, 1, 1, 1, 1, 1, 1],
            '109學年考取科學班人數': ['-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-'],
            '109學年第一志願錄取率': ['-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', 0.369685767, 0.209467456, 0.202511774, 0.141323792, 0.118236473, 0.100896861, 0.142405063, 0.101369863, 0.094827586, 0.061269147, 0.048964218, 0.047808765, 0.099099099, 0.058510638, 0.041002278, 0.055555556, 0.059760956, 0.040106952, 0.046153846, 0.047770701, 0.066666667, 0.022869023, 0.063218391, 0.192982456, 0.032467532, 0.083333333, 0.041474654, 0.030534351, 0.029411765, 0.016528926, 0.025, 0.039473684, 0.022123894, 0.024875622, 0.020491803, 0.046296296, 0.013289037, 0.020942408, 0.025157233, 0.083333333, 0.023529412, 0.00921659, 0.014814815, 0.033333333, 0.009049774, 0.026666667, 0.035087719, 0.041666667, 0.04, 0.021052632, 0.010204082, 0.068965517, 0.047619048, 0.022988506, 0.010752688, 0.142857143, 0.074074074, 0.05, 0.008928571, 0.014285714, 0.014925373, 0.01010101, 0.003412969, 0.016666667],
            '110學年免試人數': ['-', 1, '-', 2, '-', '-', 1, 7, '-', '-', '-', '-', 1, 188, 179, 126, 70, 68, 41, 31, 31, 26, 28, 32, 22, 27, 14, 19, 38, 15, 26, 12, 17, 10, 16, 13, 13, 8, 5, 8, 4, 15, 6, 2, 1, 11, 8, 4, 5, 3, 4, 4, 1, 1, 4, 2, 1, 3, 2, 2, 2, 2, 3, 4, '-', 1, 4, 2, 1, 1, 1, 2, '-', '-', 1, 2, 2],
            '110學年考取科學班人數': ['-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-'],
            '110學年第一志願錄取率': ['-', 0.026315789, '-', 0.042553191, '-', '-', 0.055555556, 0.021341463, '-', '-', '-', '-', 0.041666667, 0.346863469, 0.211084906, 0.21, 0.116666667, 0.136820926, 0.092970522, 0.099358974, 0.088825215, 0.07761194, 0.061674009, 0.067510549, 0.052757794, 0.122727273, 0.036939314, 0.042696629, 0.121794872, 0.059288538, 0.065822785, 0.035928144, 0.050295858, 0.062893082, 0.036281179, 0.077844311, 0.066666667, 0.025974026, 0.042372881, 0.036036036, 0.015503876, 0.064377682, 0.016666667, 0.008695652, 0.007042254, 0.042801556, 0.036697248, 0.016260163, 0.056179775, 0.011627907, 0.024691358, 0.03125, 0.037037037, 0.009803922, 0.017316017, 0.014925373, 0.021276596, 0.013636364, 0.029850746, 0.032786885, 0.030769231, 0.033898305, 0.03125, 0.021621622, '-', 0.027777778, 0.043010753, 0.01459854, 0.125, 0.043478261, 0.058823529, 0.018348624, '-', '-', 0.0125, 0.00660066, 0.05]
        }
        
        for year in ['111', '112', '113', '114']:
            data[f'{year}學年免試人數'] = ['-'] * len(data['學校'])
            data[f'{year}學年考取科學班人數'] = ['-'] * len(data['學校'])
            data[f'{year}學年第一志願錄取率'] = ['-'] * len(data['學校'])
        
        return pd.DataFrame(data)
    
    def get_data(self) -> Optional[pd.DataFrame]:
        """Get current data — try GAS API first, fall back to local CSV"""
        gas_df = self.load_from_gas()
        if gas_df is not None and not gas_df.empty:
            return gas_df
        
        try:
            if os.path.exists(self.data_file):
                return pd.read_csv(self.data_file)
            return None
        except Exception as e:
            print(f"Error loading data: {e}")
            return None
    
    def get_year_columns(self, df: pd.DataFrame) -> List[str]:
        """Get all year-related columns from dataframe"""
        year_columns = []
        for col in df.columns:
            if re.match(r'\d{3}學年', col):
                year_columns.append(col)
        return sorted(year_columns)
    
    def validate_csv_structure(self, df: pd.DataFrame) -> Dict[str, Any]:
        """Validate uploaded CSV structure"""
        try:
            required_base_columns = ['學校', '完整正式校名', '公立/私立', '台南市行政區']
            
            missing_base_cols = [col for col in required_base_columns if col not in df.columns]
            if missing_base_cols:
                return {
                    'valid': False,
                    'error': f"缺少必要欄位: {', '.join(missing_base_cols)}"
                }
            
            year_columns = []
            for col in df.columns:
                if re.match(r'\d{3}學年', col):
                    year_columns.append(col)
            
            if not year_columns:
                return {
                    'valid': False,
                    'error': "未找到學年度相關欄位（格式：XXX學年...）"
                }
            
            if df['學校'].isnull().any():
                return {
                    'valid': False,
                    'error': "學校名稱欄位包含空值"
                }
            
            valid_school_types = ['公立', '私立']
            invalid_types = df[~df['公立/私立'].isin(valid_school_types)]['公立/私立'].dropna().unique()
            if len(invalid_types) > 0:
                return {
                    'valid': False,
                    'error': f"公立/私立欄位包含無效值: {', '.join(invalid_types)}"
                }
            
            return {
                'valid': True,
                'year_columns': year_columns,
                'total_schools': len(df)
            }
            
        except Exception as e:
            return {
                'valid': False,
                'error': f"驗證過程發生錯誤: {str(e)}"
            }
    
    def update_data(self, new_df: pd.DataFrame) -> bool:
        """Update data with new CSV"""
        try:
            if os.path.exists(self.data_file):
                current_df = pd.read_csv(self.data_file)
                current_df.to_csv(self.backup_file, index=False, encoding='utf-8-sig')
            
            new_df.to_csv(self.data_file, index=False, encoding='utf-8-sig')
            
            timestamp_file = os.path.join("data", 'last_update.txt')
            with open(timestamp_file, 'w', encoding='utf-8') as f:
                f.write(datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            
            return True
            
        except Exception as e:
            print(f"Error updating data: {e}")
            if os.path.exists(self.backup_file):
                try:
                    backup_df = pd.read_csv(self.backup_file)
                    backup_df.to_csv(self.data_file, index=False, encoding='utf-8-sig')
                except:
                    pass
            return False
    
    def get_available_years(self, df: pd.DataFrame) -> List[str]:
        """Get available academic years from data"""
        years = set()
        for col in df.columns:
            match = re.match(r'(\d{3})學年', col)
            if match:
                years.add(match.group(1))
        return sorted(list(years))
    
    def get_last_update_time(self) -> str:
        """Get the last data update time"""
        timestamp_file = os.path.join("data", 'last_update.txt')
        try:
            if os.path.exists(timestamp_file):
                with open(timestamp_file, 'r', encoding='utf-8') as f:
                    return f.read().strip()
            else:
                if os.path.exists(self.data_file):
                    mtime = os.path.getmtime(self.data_file)
                    return datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M:%S')
                return "未知"
        except Exception:
            return "未知"
    
    def analyze_admission_trend(self, school_data: pd.Series) -> Dict[str, str]:
        """Analyze admission rate trend for a school"""
        trend_analysis = {
            'trend': '無足夠數據',
            'direction': '',
            'description': ''
        }
        
        try:
            rates = []
            years = ['109', '110', '111', '112', '113', '114']
            
            for year in years:
                col_name = f"{year}學年第一志願錄取率"
                if col_name in school_data and pd.notna(school_data[col_name]):
                    try:
                        rate = float(school_data[col_name])
                        if rate > 0:
                            rates.append(rate)
                    except:
                        continue
            
            if len(rates) < 3:
                return trend_analysis
            
            x = list(range(len(rates)))
            y = rates
            
            n = len(rates)
            sum_x = sum(x)
            sum_y = sum(y)
            sum_xy = sum(x[i] * y[i] for i in range(n))
            sum_x2 = sum(x[i] ** 2 for i in range(n))
            
            slope = (n * sum_xy - sum_x * sum_y) / (n * sum_x2 - sum_x ** 2)
            
            if abs(slope) < 0.01:
                trend_analysis['trend'] = '穩定'
                trend_analysis['direction'] = '→'
                trend_analysis['description'] = '錄取率保持相對穩定'
            elif slope > 0:
                if slope > 0.05:
                    trend_analysis['trend'] = '顯著上升'
                    trend_analysis['direction'] = '↗'
                    trend_analysis['description'] = '錄取率呈現明顯上升趨勢'
                else:
                    trend_analysis['trend'] = '微幅上升'
                    trend_analysis['direction'] = '↗'
                    trend_analysis['description'] = '錄取率緩慢上升'
            else:
                if slope < -0.05:
                    trend_analysis['trend'] = '顯著下降'
                    trend_analysis['direction'] = '↘'
                    trend_analysis['description'] = '錄取率呈現明顯下降趨勢'
                else:
                    trend_analysis['trend'] = '微幅下降'
                    trend_analysis['direction'] = '↘'
                    trend_analysis['description'] = '錄取率緩慢下降'
            
        except Exception as e:
            trend_analysis['description'] = f'趨勢分析錯誤: {str(e)}'
        
        return trend_analysis
    
    def search_schools(self, df: pd.DataFrame, search_term: str) -> pd.DataFrame:
        """Search schools by name with fuzzy matching"""
        if not search_term or search_term.strip() == '':
            return df
        
        search_term = search_term.strip().lower()
        
        exact_match = df[df['學校'].str.contains(search_term, case=False, na=False)]
        if not exact_match.empty:
            return exact_match
        
        fuzzy_matches = []
        for idx, row in df.iterrows():
            school_name = str(row['學校']).lower()
            if any(char in school_name for char in search_term):
                fuzzy_matches.append(idx)
        
        if fuzzy_matches:
            return df.loc[fuzzy_matches]
        
        return pd.DataFrame()
