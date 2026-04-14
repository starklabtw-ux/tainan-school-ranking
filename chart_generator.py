import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import pandas as pd
import re
from typing import Optional, List, Dict, Any

ACADEMIC_COLORS = ['#1B3A5C', '#1E6B4A', '#8B1A2D', '#B05A00', '#4A2070']
GRID_COLOR = '#E8ECF0'
AXIS_COLOR = '#6B7280'
BG_COLOR = 'white'
TEXT_COLOR = '#2C3E50'

class ChartGenerator:
    def __init__(self):
        self.colors = ACADEMIC_COLORS
        self.background_color = BG_COLOR
        self.text_color = TEXT_COLOR
    
    def _academic_layout(self, fig: go.Figure, **extra_kwargs) -> go.Figure:
        """Apply academic journal style layout to a figure"""
        base = dict(
            plot_bgcolor=BG_COLOR,
            paper_bgcolor=BG_COLOR,
            font=dict(color=TEXT_COLOR, size=11),
            legend=dict(
                orientation="h",
                yanchor="top",
                y=-0.18,
                xanchor="center",
                x=0.5,
                font=dict(size=10, color=TEXT_COLOR),
                bgcolor="rgba(0,0,0,0)",
                borderwidth=0
            ),
            xaxis=dict(
                showgrid=False,
                zeroline=False,
                linecolor=GRID_COLOR,
                tickfont=dict(color=AXIS_COLOR, size=10),
                title_font=dict(color=TEXT_COLOR, size=11)
            ),
            yaxis=dict(
                showgrid=True,
                gridcolor=GRID_COLOR,
                gridwidth=0.5,
                zeroline=False,
                linecolor=GRID_COLOR,
                tickfont=dict(color=AXIS_COLOR, size=10),
                title_font=dict(color=TEXT_COLOR, size=11)
            ),
            margin=dict(l=50, r=20, t=20, b=80),
            hovermode='x unified'
        )
        base.update(extra_kwargs)
        fig.update_layout(**base)
        return fig

    def _extract_year_data(self, df: pd.DataFrame, data_type: str) -> Dict[str, Dict[str, float]]:
        """Extract year data for specific data type"""
        year_data = {}
        
        for col in df.columns:
            match = re.match(r'(\d{3})學年' + data_type, col)
            if match:
                year = match.group(1)
                year_data[year] = {}
                
                for idx, row in df.iterrows():
                    school = row['學校']
                    value = row[col]
                    
                    if pd.notna(value) and str(value) not in ['-', '#VALUE!']:
                        try:
                            year_data[year][school] = float(value)
                        except (ValueError, TypeError):
                            pass
        
        return year_data
    
    def create_admission_rate_comparison(self, schools_df: pd.DataFrame) -> Optional[go.Figure]:
        """Create admission rate comparison chart for multiple schools"""
        try:
            year_data = self._extract_year_data(schools_df, '第一志願錄取率')
            
            if not year_data:
                return None
            
            fig = go.Figure()
            
            years = sorted(year_data.keys())
            schools = schools_df['學校'].tolist()
            
            for i, school in enumerate(schools):
                rates = []
                display_years = []
                
                for year in years:
                    if school in year_data[year]:
                        rates.append(year_data[year][school] * 100)
                        display_years.append(f"{year}學年")
                
                if rates:
                    fig.add_trace(go.Scatter(
                        x=display_years,
                        y=rates,
                        mode='lines+markers',
                        name=school,
                        line=dict(color=self.colors[i % len(self.colors)], width=1.5),
                        marker=dict(size=5, symbol='circle')
                    ))
            
            fig = self._academic_layout(fig, yaxis_title="錄取率（％）", xaxis_title="學年度")
            return fig
            
        except Exception as e:
            print(f"Error creating admission rate chart: {e}")
            return None
    
    def create_student_count_comparison(self, schools_df: pd.DataFrame) -> Optional[go.Figure]:
        """Create student count comparison chart (免試人數 + 科學班人數)"""
        try:
            admission_data = self._extract_year_data(schools_df, '免試人數')
            science_data = self._extract_year_data(schools_df, '考取科學班人數')
            
            if not admission_data:
                return None
            
            fig = go.Figure()
            
            years = sorted(admission_data.keys())
            schools = schools_df['學校'].tolist()
            
            for i, school in enumerate(schools):
                total_counts = []
                display_years = []
                
                for year in years:
                    admission_count = admission_data[year].get(school, 0)
                    science_count = science_data.get(year, {}).get(school, 0)
                    total_count = admission_count + science_count
                    
                    if total_count > 0:
                        total_counts.append(total_count)
                        display_years.append(f"{year}學年")
                
                if total_counts:
                    fig.add_trace(go.Scatter(
                        x=display_years,
                        y=total_counts,
                        mode='lines+markers',
                        name=school,
                        line=dict(color=self.colors[i % len(self.colors)], width=1.5),
                        marker=dict(size=5, symbol='circle')
                    ))
            
            fig = self._academic_layout(fig, yaxis_title="學生人數（人）", xaxis_title="學年度")
            return fig
            
        except Exception as e:
            print(f"Error creating student count chart: {e}")
            return None
    
    def create_single_school_admission_rate(self, school_data: pd.Series) -> Optional[go.Figure]:
        """Create admission rate trend chart for single school"""
        try:
            years = []
            rates = []
            
            for col in school_data.index:
                match = re.match(r'(\d{3})學年第一志願錄取率', col)
                if match:
                    year = match.group(1)
                    value = school_data[col]
                    
                    if pd.notna(value) and str(value) not in ['-', '#VALUE!']:
                        try:
                            rate = float(value) * 100
                            years.append(f"{year}學年")
                            rates.append(rate)
                        except (ValueError, TypeError):
                            pass
            
            if not years:
                return None
            
            fig = go.Figure()
            
            fig.add_trace(go.Scatter(
                x=years,
                y=rates,
                mode='lines+markers+text',
                name='錄取率',
                line=dict(color=self.colors[0], width=1.5),
                marker=dict(size=7, symbol='circle'),
                text=[f"{rate:.2f}%" for rate in rates],
                textposition="top center",
                textfont=dict(size=10, color=TEXT_COLOR)
            ))
            
            fig = self._academic_layout(fig, showlegend=False, yaxis_title="錄取率（％）", xaxis_title="學年度")
            return fig
            
        except Exception as e:
            print(f"Error creating single school admission rate chart: {e}")
            return None
    
    def create_single_school_student_count(self, school_data: pd.Series) -> Optional[go.Figure]:
        """Create student count trend chart for single school"""
        try:
            years = []
            admission_counts = []
            science_counts = []
            
            for col in school_data.index:
                admission_match = re.match(r'(\d{3})學年免試人數', col)
                science_match = re.match(r'(\d{3})學年考取科學班人數', col)
                
                if admission_match:
                    year = admission_match.group(1)
                    value = school_data[col]
                    
                    if pd.notna(value) and str(value) not in ['-', '#VALUE!']:
                        try:
                            count = int(float(value))
                            if f"{year}學年" not in years:
                                years.append(f"{year}學年")
                                admission_counts.append(count)
                                science_counts.append(0)
                            else:
                                idx = years.index(f"{year}學年")
                                admission_counts[idx] = count
                        except (ValueError, TypeError):
                            pass
                
                elif science_match:
                    year = science_match.group(1)
                    value = school_data[col]
                    
                    if pd.notna(value) and str(value) not in ['-', '#VALUE!']:
                        try:
                            count = int(float(value))
                            if f"{year}學年" not in years:
                                years.append(f"{year}學年")
                                admission_counts.append(0)
                                science_counts.append(count)
                            else:
                                idx = years.index(f"{year}學年")
                                science_counts[idx] = count
                        except (ValueError, TypeError):
                            pass
            
            if not years:
                return None
            
            year_data = list(zip(years, admission_counts, science_counts))
            year_data.sort(key=lambda x: x[0])
            years, admission_counts, science_counts = zip(*year_data)
            
            fig = go.Figure()
            
            fig.add_trace(go.Bar(
                x=years,
                y=admission_counts,
                name='免試人數',
                marker_color=self.colors[0],
                text=[str(int(count)) if count > 0 else '' for count in admission_counts],
                textposition='inside',
                textfont=dict(size=10, color='white')
            ))
            
            fig.add_trace(go.Bar(
                x=years,
                y=science_counts,
                name='科學班人數',
                marker_color=self.colors[1],
                text=[str(int(count)) if count > 0 else '' for count in science_counts],
                textposition='inside',
                textfont=dict(size=10, color='white')
            ))
            
            fig = self._academic_layout(fig, barmode='stack', yaxis_title="學生人數（人）", xaxis_title="學年度")
            return fig
            
        except Exception as e:
            print(f"Error creating single school student count chart: {e}")
            return None
