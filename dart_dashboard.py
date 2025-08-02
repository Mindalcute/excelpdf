# -*- coding: utf-8 -*-
import os
import sys
import locale
import io
import base64
import re
from datetime import datetime, timedelta
import random
import numpy as np
import json

# 환경 설정
os.environ['PYTHONIOENCODING'] = 'utf-8'
os.environ['LANG'] = 'ko_KR.UTF-8'

try:
    locale.setlocale(locale.LC_ALL, 'ko_KR.UTF-8')
except:
    try:
        locale.setlocale(locale.LC_ALL, 'Korean_Korea.utf8')
    except:
        pass

import streamlit as st
import pandas as pd
import requests
import zipfile
import xml.etree.ElementTree as ET
import feedparser

# 환경 변수 로드 (선택사항)
try:
    from dotenv import load_dotenv
    load_dotenv()
except:
    pass

# plotly 안전하게 import
try:
    import plotly.express as px
    import plotly.graph_objects as go
    PLOTLY_AVAILABLE = True
except ImportError:
    PLOTLY_AVAILABLE = False

from bs4 import BeautifulSoup

# PDF 생성용 라이브러리
try:
    from reportlab.lib.pagesizes import letter, A4
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Image, PageTemplate, Frame
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.lib.units import inch
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

# 워드클라우드 라이브러리
try:
    from wordcloud import WordCloud
    import matplotlib.pyplot as plt
    import matplotlib.font_manager as fm
    WORDCLOUD_AVAILABLE = True
except ImportError:
    WORDCLOUD_AVAILABLE = False

# 구글시트 연동 라이브러리 추가
try:
    import gspread
    from google.oauth2.service_account import Credentials
    from google.oauth2 import service_account
    GOOGLE_SHEETS_AVAILABLE = True
except ImportError:
    GOOGLE_SHEETS_AVAILABLE = False

st.set_page_config(page_title="SK에너지 경쟁사 분석 대시보드", page_icon="⚡", layout="wide")

# API 키 설정
DART_API_KEY = "9a153f4344ad2db546d651090f78c8770bd773cb"

# SK 브랜드 컬러 테마
SK_COLORS = {
    'primary': '#E31E24',      # SK 레드
    'secondary': '#FF6B35',    # SK 오렌지 
    'accent': '#004EA2',       # SK 블루
    'success': '#00A651',      # 성공 색상
    'warning': '#FF9500',      # 경고 색상
    'competitor': '#6C757D',   # 기본 경쟁사 색상 (회색)
    
    # 개별 경쟁사 파스텔 색상
    'competitor_1': '#AEC6CF',  # 파스텔 블루
    'competitor_2': '#FFB6C1',  # 파스텔 핑크  
    'competitor_3': '#98FB98',  # 파스텔 그린
    'competitor_4': '#F0E68C',  # 파스텔 옐로우
    'competitor_5': '#DDA0DD',  # 파스텔 퍼플
}

# SK 브랜드 CSS 설정
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;700&display=swap');

/* SK 브랜드 컬러 적용 */
html, body, [class*="css"] {{
    font-family: 'Noto Sans KR', 'Malgun Gothic', '맑은 고딕', sans-serif !important;
}}

/* SK 로고 컨테이너 */
.sk-header {{
    background: linear-gradient(135deg, {SK_COLORS['primary']} 0%, {SK_COLORS['secondary']} 100%);
    padding: 20px;
    border-radius: 15px;
    margin-bottom: 20px;
    text-align: center;
    color: white;
    box-shadow: 0 4px 12px rgba(227, 30, 36, 0.3);
}}

/* SK에너지 특별 강조 스타일 */
.sk-energy-highlight {{
    background: linear-gradient(135deg, {SK_COLORS['primary']} 0%, {SK_COLORS['secondary']} 100%);
    color: white !important;
    font-weight: 700 !important;
    padding: 8px 15px;
    border-radius: 8px;
    box-shadow: 0 2px 8px rgba(227, 30, 36, 0.4);
}}

/* 경쟁사 스타일 */
.competitor-style {{
    background: {SK_COLORS['competitor']};
    color: white;
    padding: 6px 12px;
    border-radius: 6px;
}}

/* 다운로드 버튼 SK 테마 */
.stDownloadButton > button {{
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    padding: 0.5rem 1rem !important;
    font-weight: 600 !important;
    transition: all 0.3s ease !important;
}}

.stDownloadButton:nth-child(1) > button {{
    background: linear-gradient(135deg, #4CAF50 0%, #45a049 100%) !important;
}}

.stDownloadButton:nth-child(2) > button {{
    background: linear-gradient(135deg, {SK_COLORS['primary']} 0%, {SK_COLORS['secondary']} 100%) !important;
}}

.stDownloadButton:nth-child(3) > button {{
    background: linear-gradient(135deg, #2196F3 0%, #1976D2 100%) !important;
}}

/* 분석 결과 카드 스타일 */
.analysis-card {{
    background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
    padding: 20px;
    border-radius: 15px;
    margin: 15px 0;
    border-left: 5px solid {SK_COLORS['primary']};
    box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
}}

/* SK 인사이트 카드 */
.sk-insight-card {{
    background: linear-gradient(135deg, {SK_COLORS['primary']} 0%, {SK_COLORS['secondary']} 100%);
    color: white;
    padding: 20px;
    border-radius: 12px;
    margin: 15px 0;
    box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
}}

/* 테이블 스타일 개선 */
.analysis-table {{
    width: 100%;
    border-collapse: collapse;
    margin: 20px 0;
    font-size: 14px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.1);
}}

.analysis-table th {{
    background: linear-gradient(135deg, {SK_COLORS['primary']} 0%, {SK_COLORS['secondary']} 100%);
    color: white;
    padding: 15px;
    text-align: center;
    font-weight: 600;
    border: 1px solid #ddd;
}}

.analysis-table td {{
    padding: 12px;
    border: 1px solid #ddd;
    text-align: center;
}}

.analysis-table tr:nth-child(even) {{
    background-color: #f8f9fa;
}}

.analysis-table tr:hover {{
    background-color: #e3f2fd;
}}

/* SK 데이터 특별 강조 */
.sk-data-cell {{
    background: linear-gradient(135deg, {SK_COLORS['primary']}20 0%, {SK_COLORS['secondary']}20 100%) !important;
    font-weight: 600 !important;
    color: {SK_COLORS['primary']} !important;
}}

/* DART 출처 정보 스타일 */
.dart-source-box {{
    margin-top: 20px;
    padding: 15px;
    background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
    border-radius: 8px;
    border-left: 4px solid {SK_COLORS['primary']};
    font-size: 12px;
    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
}}

.dart-source-box h4 {{
    margin: 0 0 10px 0;
    color: {SK_COLORS['primary']};
    font-weight: 600;
}}

.dart-source-box a {{
    color: {SK_COLORS['primary']};
    text-decoration: none;
    font-weight: 500;
}}

.dart-source-box a:hover {{
    text-decoration: underline;
}}

/* 손실 표시 스타일 */
.loss-indicator {{
    color: #dc3545 !important;
    font-weight: 700 !important;
}}

.profit-indicator {{
    color: #28a745 !important;
    font-weight: 600 !important;
}}
</style>
""", unsafe_allow_html=True)

# 세션 상태 초기화
if 'analysis_results' not in st.session_state:
    st.session_state.analysis_results = None

if 'comparison_metric' not in st.session_state:
    st.session_state.comparison_metric = "매출 대비 비율"

if 'quarterly_data' not in st.session_state:
    st.session_state.quarterly_data = None

# ==========================
# 회사별 색상 할당 함수
# ==========================
def get_company_color(company_name, all_companies):
    """회사별 고유 색상 반환 (SK는 빨간색, 경쟁사는 파스텔 구분)"""
    if 'SK' in company_name:
        return SK_COLORS['primary']
    else:
        # 경쟁사들에게 서로 다른 파스텔 색상 할당
        competitor_colors = [
            SK_COLORS['competitor_1'],  # 파스텔 블루
            SK_COLORS['competitor_2'],  # 파스텔 핑크
            SK_COLORS['competitor_3'],  # 파스텔 그린
            SK_COLORS['competitor_4'],  # 파스텔 옐로우
            SK_COLORS['competitor_5']   # 파스텔 퍼플
        ]
        
        # SK가 아닌 회사들의 인덱스 계산
        non_sk_companies = [comp for comp in all_companies if 'SK' not in comp]
        try:
            index = non_sk_companies.index(company_name)
            return competitor_colors[index % len(competitor_colors)]
        except ValueError:
            return SK_COLORS['competitor']

# ==========================
# 분기별 데이터 수집 클래스
# ==========================
class QuarterlyDataCollector:
    def __init__(self, dart_collector):
        self.dart_collector = dart_collector
        self.report_codes = {
            "Q1": "11013",  # 1분기
            "Q2": "11012",  # 반기 (1,2분기 누적)
            "Q3": "11014",  # 3분기
            "Q4": "11011"   # 사업보고서 (연간)
        }
    
    def collect_quarterly_data(self, company_name, year=2024):
        """분기별 재무 데이터 수집"""
        quarterly_results = []
        
        for quarter, report_code in self.report_codes.items():
            st.write(f"📊 {company_name} {quarter} 데이터 수집중...")
            
            corp_code = self.dart_collector.get_corp_code_enhanced(company_name)
            if not corp_code:
                continue
                
            df = self.dart_collector.get_financial_statement(corp_code, str(year), report_code)
            
            if not df.empty:
                # 주요 지표 추출
                quarterly_metrics = self._extract_key_metrics(df, quarter)
                if quarterly_metrics:
                    quarterly_metrics['회사'] = company_name
                    quarterly_metrics['연도'] = year
                    quarterly_results.append(quarterly_metrics)
        
        return pd.DataFrame(quarterly_results) if quarterly_results else pd.DataFrame()
    
    def _extract_key_metrics(self, df, quarter):
        """주요 재무 지표 추출"""
        metrics = {'분기': quarter}
        
        # 매출액 추출
        revenue_keywords = ['매출액', 'revenue', 'sales']
        for keyword in revenue_keywords:
            revenue_rows = df[df['account_nm'].str.contains(keyword, case=False, na=False)]
            if not revenue_rows.empty:
                try:
                    amount = float(str(revenue_rows.iloc[0]['thstrm_amount']).replace(',', '').replace('-', '0'))
                    metrics['매출액'] = amount / 1_000_000_000_000  # 조원 단위
                    break
                except:
                    continue
        
        # 영업이익 추출
        operating_keywords = ['영업이익', 'operating']
        for keyword in operating_keywords:
            op_rows = df[df['account_nm'].str.contains(keyword, case=False, na=False)]
            if not op_rows.empty:
                try:
                    amount = float(str(op_rows.iloc[0]['thstrm_amount']).replace(',', '').replace('-', '0'))
                    metrics['영업이익'] = amount / 100_000_000  # 억원 단위
                    break
                except:
                    continue
        
        # 영업이익률 계산
        if '매출액' in metrics and '영업이익' in metrics and metrics['매출액'] > 0:
            metrics['영업이익률'] = (metrics['영업이익'] * 100) / (metrics['매출액'] * 10)  # % 단위
        
        return metrics if len(metrics) > 1 else None

# ==========================
# DART API 연동 클래스 (출처 추적 기능 포함)
# ==========================
class DartAPICollector:
    def __init__(self, api_key):
        self.api_key = api_key
        self.report_codes = {
            "1분기": "11013",
            "반기": "11012", 
            "3분기": "11014",
            "사업": "11011"
        }
        
        # 출처 추적용 딕셔너리
        self.source_tracking = {}
        
        # 회사명 매핑
        self.company_name_mapping = {
            "SK에너지": [
                "SK이노베이션", "SK에너지", "에스케이이노베이션", 
                "에스케이에너지", "SK이노베이션주식회사", "SK에너지주식회사"
            ],
            "SK이노베이션": [
                "SK이노베이션", "에스케이이노베이션", "SK이노베이션주식회사"
            ],
            "GS칼텍스": [
                "GS칼텍스", "지에스칼텍스", "GS칼텍스주식회사", "지에스칼텍스주식회사"
            ],
            "HD현대오일뱅크": [
                "현대오일뱅크", "현대오일뱅크주식회사", "HD현대오일뱅크", 
                "HD현대오일뱅크주식회사"
            ],
            "현대오일뱅크": [
                "현대오일뱅크", "현대오일뱅크주식회사"
            ],
            # S-Oil 검색 강화 (종목코드 포함)
            "S-Oil": [
                "S-Oil", "S-Oil Corporation", "S-Oil Corp", "에쓰오일", "에스오일",
                "주식회사S-Oil", "S-OIL", "s-oil", "010950"
            ]
        }
    
    def get_corp_code(self, company_name):
        """기본 회사 고유번호 조회"""
        url = f"https://opendart.fss.or.kr/api/corpCode.xml?crtfc_key={self.api_key}"
        try:
            res = requests.get(url)
            with zipfile.ZipFile(io.BytesIO(res.content)) as z:
                xml_file = z.open(z.namelist()[0])
                tree = ET.parse(xml_file)
                root = tree.getroot()
                for corp in root.findall("list"):
                    corp_name = corp.find("corp_name")
                    if corp_name is not None and company_name in corp_name.text:
                        return corp.find("corp_code").text
        except Exception as e:
            st.error(f"회사 코드 조회 오류: {e}")
        return None
    
    def get_corp_code_enhanced(self, company_name):
        """강화된 회사 고유번호 조회"""
        url = f"https://opendart.fss.or.kr/api/corpCode.xml?crtfc_key={self.api_key}"
        
        search_names = self.company_name_mapping.get(company_name, [company_name])
        
        st.write(f"🔍 {company_name} 검색 시도 중... (후보: {len(search_names)}개)")
        
        try:
            res = requests.get(url)
            with zipfile.ZipFile(io.BytesIO(res.content)) as z:
                xml_file = z.open(z.namelist()[0])
                tree = ET.parse(xml_file)
                root = tree.getroot()
                
                # 모든 회사 목록에서 매칭 시도
                all_companies = []
                for corp in root.findall("list"):
                    corp_name_elem = corp.find("corp_name")
                    corp_code_elem = corp.find("corp_code")
                    stock_code_elem = corp.find("stock_code")
                    if corp_name_elem is not None and corp_code_elem is not None:
                        all_companies.append({
                            'name': corp_name_elem.text,
                            'code': corp_code_elem.text,
                            'stock_code': stock_code_elem.text if stock_code_elem is not None else None
                        })
                
                # 여러 단계로 검색
                for search_name in search_names:
                    st.write(f"   📋 '{search_name}' 검색 중...")
                    
                    # 1단계: 종목코드로 검색 (S-Oil 전용)
                    if search_name.isdigit():
                        for company in all_companies:
                            if company['stock_code'] == search_name:
                                st.success(f"✅ 종목코드 매칭: {company['name']} → {company['code']}")
                                return company['code']
                    
                    # 2단계: 정확히 일치
                    for company in all_companies:
                        if company['name'] == search_name:
                            st.success(f"✅ 정확 매칭: {company['name']} → {company['code']}")
                            return company['code']
                    
                    # 3단계: 포함 검색
                    for company in all_companies:
                        if search_name in company['name'] or company['name'] in search_name:
                            st.success(f"✅ 부분 매칭: {company['name']} → {company['code']}")
                            return company['code']
                    
                    # 4단계: 대소문자 무시 검색
                    for company in all_companies:
                        if search_name.lower() in company['name'].lower() or company['name'].lower() in search_name.lower():
                            st.success(f"✅ 대소문자 무시 매칭: {company['name']} → {company['code']}")
                            return company['code']
                
                # 특별 케이스: S-Oil 관련 직접 검색
                if 'oil' in company_name.lower() or 'Oil' in company_name:
                    st.write("   🛢️ 오일 관련 회사 전체 검색...")
                    oil_companies = []
                    for company in all_companies:
                        if any(keyword in company['name'].lower() for keyword in ['oil', '오일', '에쓰', '에스']):
                            oil_companies.append(company)
                    
                    if oil_companies:
                        st.write(f"   🔍 발견된 오일 관련 회사: {len(oil_companies)}개")
                        for oil_comp in oil_companies:
                            st.write(f"      - {oil_comp['name']} ({oil_comp['code']}) 종목:{oil_comp['stock_code']}")
                        return oil_companies[0]['code']
                
                st.error(f"❌ {company_name}의 회사 코드를 찾을 수 없습니다.")
                return None
                
        except Exception as e:
            st.error(f"회사 코드 조회 오류: {e}")
            return None
    
    def get_financial_statement(self, corp_code, bsns_year, reprt_code, fs_div="CFS"):
        """재무제표 조회"""
        url = "https://opendart.fss.or.kr/api/fnlttSinglAcntAll.json"
        params = {
            "crtfc_key": self.api_key,
            "corp_code": corp_code,
            "bsns_year": bsns_year,
            "reprt_code": reprt_code,
            "fs_div": fs_div
        }
        
        try:
            res = requests.get(url, params=params).json()
            if res.get("status") == "000" and "list" in res:
                df = pd.DataFrame(res["list"])
                df["보고서구분"] = reprt_code
                return df
            else:
                st.warning(f"❌ {reprt_code} 재무제표 데이터를 가져오지 못했습니다: {res.get('message')}")
                return pd.DataFrame()
        except Exception as e:
            st.error(f"재무제표 조회 오류: {e}")
            return pd.DataFrame()
    
    def convert_stock_to_corp_code(self, stock_code):
        """종목코드를 DART 회사코드로 변환"""
        try:
            url = f"https://opendart.fss.or.kr/api/corpCode.xml?crtfc_key={self.api_key}"
            res = requests.get(url)
            
            with zipfile.ZipFile(io.BytesIO(res.content)) as z:
                xml_file = z.open(z.namelist()[0])
                tree = ET.parse(xml_file)
                root = tree.getroot()
                
                # 종목코드로 매칭
                for corp in root.findall("list"):
                    stock_elem = corp.find("stock_code")
                    corp_code_elem = corp.find("corp_code")
                    
                    if (stock_elem is not None and 
                        corp_code_elem is not None and 
                        stock_elem.text == stock_code):
                        return corp_code_elem.text
            
            return None
        except Exception as e:
            st.error(f"종목코드 변환 오류: {e}")
            return None
    
    def get_company_financials_auto(self, company_name, bsns_year):
        """회사 재무제표 자동 수집 (출처 추적 포함)"""
        
        # 종목코드 직접 매핑
        STOCK_CODE_MAPPING = {
            "S-Oil": "010950",
            "GS칼텍스": "089590", 
            "현대오일뱅크": "267250",
            "SK에너지": "096770",
        }
        
        st.info(f"📌 {company_name} 검색 시작...")
        
        # 1. 종목코드로 직접 시도 (S-Oil 전용)
        if company_name in STOCK_CODE_MAPPING:
            stock_code = STOCK_CODE_MAPPING[company_name]
            st.info(f"📌 {company_name} 종목코드 {stock_code}로 직접 접근 시도...")
            
            corp_code = self.convert_stock_to_corp_code(stock_code)
            if corp_code:
                st.success(f"✅ 종목코드 변환 성공: {stock_code} → {corp_code}")
                
                # 재무제표 직접 조회
                report_codes = ["11011", "11014", "11012"]
                for report_code in report_codes:
                    df = self.get_financial_statement(corp_code, bsns_year, report_code)
                    if not df.empty:
                        # 출처 정보 저장
                        self._save_source_info(company_name, corp_code, report_code, bsns_year)
                        # 🔗 실제 보고서 정보 추가 저장
                        self._save_actual_report_info(company_name, corp_code, report_code, bsns_year)
                        st.success(f"✅ {company_name} 재무제표 수집 완료 (종목코드 접근)")
                        return df
        
        # 2. 기존 검색 방식으로 폴백
        corp_code = self.get_corp_code_enhanced(company_name)
        
        if not corp_code:
            st.error(f"❌ {company_name}의 회사 코드를 찾을 수 없습니다.")
            return None
        
        st.info(f"📌 {company_name} (코드: {corp_code}) 재무제표 수집 중...")
        
        # 여러 보고서 타입 시도
        report_codes = ["11011", "11014", "11012"]
        
        for report_code in report_codes:
            df = self.get_financial_statement(corp_code, bsns_year, report_code)
            if not df.empty:
                # 출처 정보 저장
                self._save_source_info(company_name, corp_code, report_code, bsns_year)
                # 🔗 실제 보고서 정보 추가 저장
                self._save_actual_report_info(company_name, corp_code, report_code, bsns_year)
                st.success(f"✅ {company_name} 재무제표 수집 완료 (보고서: {report_code})")
                return df
        
        st.error(f"❌ {company_name} 모든 보고서 타입에서 데이터 없음")
        return None
    
    def _save_source_info(self, company_name, corp_code, report_code, bsns_year):
        """출처 정보 저장"""
        report_type_map = {
            "11011": "사업보고서",
            "11014": "3분기보고서", 
            "11012": "반기보고서",
            "11013": "1분기보고서"
        }
        
        self.source_tracking[company_name] = {
            'company_code': corp_code,
            'report_code': report_code,
            'report_type': report_type_map.get(report_code, "재무제표"),
            'year': bsns_year,
            'dart_url': f"https://dart.fss.or.kr/dsaf001/main.do?rcpNo={corp_code}",
            'direct_link': f"https://dart.fss.or.kr/dsaf001/main.do?rcpNo={corp_code}&reprtCode={report_code}"
        }
    
    def _save_actual_report_info(self, company_name, corp_code, report_code, bsns_year):
        """🔗 실제 보고서 접수번호 및 직접 링크 정보 저장"""
        try:
            # DART API로 실제 보고서 목록 조회
            url = "https://opendart.fss.or.kr/api/list.json"
            params = {
                "crtfc_key": self.api_key,
                "corp_code": corp_code,
                "bgn_de": f"{bsns_year}0101",
                "end_de": f"{bsns_year}1231",
                "pblntf_ty": "A",  # 정기공시
                "page_no": 1,
                "page_count": 10
            }
            
            res = requests.get(url, params=params).json()
            
            if res.get("status") == "000" and "list" in res:
                reports = res["list"]
                
                # 해당 보고서 타입과 연도에 맞는 보고서 찾기
                report_type_names = {
                    "11011": ["사업보고서", "연간보고서"],
                    "11014": ["3분기보고서", "분기보고서"],
                    "11012": ["반기보고서", "중간보고서"],
                    "11013": ["1분기보고서", "분기보고서"]
                }
                
                target_names = report_type_names.get(report_code, [])
                
                for report in reports:
                    report_nm = report.get('report_nm', '')
                    rcept_no = report.get('rcept_no', '')
                    
                    # 보고서명에 해당 타입이 포함되어 있는지 확인
                    if any(name in report_nm for name in target_names) and rcept_no:
                        # 실제 보고서 정보 업데이트
                        if company_name in self.source_tracking:
                            self.source_tracking[company_name].update({
                                'rcept_no': rcept_no,
                                'report_name': report_nm,
                                'actual_dart_url': f"https://dart.fss.or.kr/dsaf001/main.do?rcpNo={rcept_no}",
                                'financial_statement_url': f"https://dart.fss.or.kr/report/viewer.do?rcpNo={rcept_no}&dcmNo=7957217&eleId=0&offset=0&length=0&dtd=dart3.xsd"
                            })
                        break
                        
        except Exception as e:
            st.warning(f"실제 보고서 정보 조회 중 오류: {e}")

# ==========================
# 영업손실 표시 함수 (문제 3 해결)
# ==========================
def _format_amount_with_loss_indicator(amount):
    """📉 영업손실 명확 표시 함수"""
    if amount < 0:
        abs_amount = abs(amount)
        if abs_amount >= 1_000_000_000_000:
            return f"<span class='loss-indicator'>▼ {abs_amount/1_000_000_000_000:.1f}조원 영업손실</span>"
        elif abs_amount >= 100_000_000:
            return f"<span class='loss-indicator'>▼ {abs_amount/100_000_000:.0f}억원 영업손실</span>"
        elif abs_amount >= 10_000:
            return f"<span class='loss-indicator'>▼ {abs_amount/10_000:.0f}만원 영업손실</span>"
        else:
            return f"<span class='loss-indicator'>▼ {abs_amount:,.0f}원 영업손실</span>"
    else:
        return _format_amount_profit(amount)

def _format_amount_profit(amount):
    """일반 이익 포맷팅"""
    if abs(amount) >= 1_000_000_000_000:
        return f"<span class='profit-indicator'>{amount/1_000_000_000_000:.1f}조원</span>"
    elif abs(amount) >= 100_000_000:
        return f"<span class='profit-indicator'>{amount/100_000_000:.0f}억원</span>"
    elif abs(amount) >= 10_000:
        return f"<span class='profit-indicator'>{amount/10_000:.0f}만원</span>"
    else:
        return f"<span class='profit-indicator'>{amount:,.0f}원</span>"

# ==========================
# SK 중심 재무데이터 프로세서 (손실 표시 개선)
# ==========================
class SKFinancialDataProcessor:
    INCOME_STATEMENT_MAP = {
        'sales': '매출액',
        'revenue': '매출액',
        '매출액': '매출액',
        '수익(매출액)': '매출액',
        'costofgoodssold': '매출원가',
        'cogs': '매출원가', 
        'costofrevenue': '매출원가',
        '매출원가': '매출원가',
        'operatingexpenses': '판관비',
        'sellingexpenses': '판매비',
        'administrativeexpenses': '관리비',
        '판매비와관리비': '판관비',
        '판관비': '판관비',
        'grossprofit': '매출총이익',
        '매출총이익': '매출총이익',
        'operatingincome': '영업이익',
        'operatingprofit': '영업이익',
        '영업이익': '영업이익',
        'netincome': '당기순이익',
        '당기순이익': '당기순이익',
    }
    
    def __init__(self):
        self.company_data = {}
        # SK 중심 경쟁사 정의
        self.sk_company = "SK에너지"
        self.competitors = ["GS칼텍스", "현대오일뱅크", "S-Oil"]
    
    def process_dart_data(self, dart_df, company_name):
        """DART API에서 받은 DataFrame을 표준 손익계산서로 변환 (XBRL 마이너스 값 정확 처리)"""
        try:
            if dart_df.empty:
                return None
            
            financial_data = {}
            
            for _, row in dart_df.iterrows():
                account_nm = row.get('account_nm', '')
                thstrm_amount = row.get('thstrm_amount', '0')
                
                try:
                    # 📉 마이너스 값 정확 처리 (XBRL 파서 개선)
                    amount_str = str(thstrm_amount).replace(',', '')
                    if '(' in amount_str and ')' in amount_str:
                        # 괄호로 표시된 마이너스
                        amount_str = '-' + amount_str.replace('(', '').replace(')', '')
                    
                    value = float(amount_str) if amount_str != '-' else 0
                except:
                    continue
                
                for key, mapped_name in self.INCOME_STATEMENT_MAP.items():
                    if key in account_nm or account_nm in key:
                        if mapped_name not in financial_data or abs(value) > abs(financial_data[mapped_name]):
                            financial_data[mapped_name] = value
                        break
            
            return self._create_income_statement(financial_data, company_name)
            
        except Exception as e:
            st.error(f"DART 데이터 처리 오류: {e}")
            return None
    
    def _create_income_statement(self, data, company_name):
        """표준 손익계산서 구조 생성 (손실 표시 적용)"""
        standard_items = [
            '매출액', '매출원가', '매출총이익', '판매비', '관리비', '판관비',
            '인건비', '감가상각비', '영업이익', '영업외수익', '영업외비용',
            '금융비용', '이자비용', '당기순이익'
        ]
        
        calculated_items = self._calculate_derived_items(data)
        data.update(calculated_items)
        
        income_statement = []
        
        for item in standard_items:
            value = data.get(item, 0)
            if value != 0:
                # 📉 영업이익/당기순이익은 손실 표시 적용
                if item in ['영업이익', '당기순이익']:
                    formatted_value = _format_amount_with_loss_indicator(value)
                else:
                    formatted_value = _format_amount_profit(value)
                
                income_statement.append({
                    '구분': item,
                    company_name: formatted_value
                })
        
        # 정규화 지표 추가
        ratios = self._calculate_enhanced_ratios(data)
        for ratio_name, ratio_value in ratios.items():
            if ratio_name == '매출 1조원당 영업이익(억원)':
                display_value = f"{ratio_value:.2f}억원"
            elif ratio_name.endswith('(%)'):
                display_value = f"{ratio_value:.2f}%"
            else:
                display_value = f"{ratio_value:.2f}점"
            
            income_statement.append({
                '구분': ratio_name,
                company_name: display_value
            })
            
        return pd.DataFrame(income_statement)
    
    def _calculate_derived_items(self, data):
        """파생 항목 계산"""
        calculated = {}
        
        if '매출액' in data and '매출원가' in data:
            calculated['매출총이익'] = data['매출액'] - data['매출원가']
        elif '매출액' in data and '매출총이익' not in data:
            calculated['매출총이익'] = data['매출액'] * 0.3
            calculated['매출원가'] = data['매출액'] - calculated['매출총이익']
            
        if '판매비' in data and '관리비' in data:
            calculated['판관비'] = data['판매비'] + data['관리비']
            
        return calculated
    
    def _calculate_enhanced_ratios(self, data):
        """정규화 지표 계산"""
        ratios = {}
        
        매출액 = data.get('매출액', 0)
        if 매출액 > 0:
            if '영업이익' in data:
                ratios['영업이익률(%)'] = round((data['영업이익'] / 매출액) * 100, 2)
            if '당기순이익' in data:
                ratios['순이익률(%)'] = round((data['당기순이익'] / 매출액) * 100, 2)
            if '매출원가' in data:
                ratios['매출원가율(%)'] = round((data['매출원가'] / 매출액) * 100, 2)
            if '판관비' in data:
                ratios['판관비율(%)'] = round((data['판관비'] / 매출액) * 100, 2)
                
            if '영업이익' in data:
                ratios['매출 1조원당 영업이익(억원)'] = round((data['영업이익'] / 100_000_000) / (매출액 / 1_000_000_000_000), 2)
            
            ratios['원가효율성지수(점)'] = round(100 - ratios.get('매출원가율(%)', 0), 2)
            
            operating_margin = ratios.get('영업이익률(%)', 0)
            net_margin = ratios.get('순이익률(%)', 0)
            ratios['종합수익성점수(점)'] = round((operating_margin * 2 + net_margin) / 3, 2)
            
            industry_avg_margin = 3.5
            if operating_margin > 0:
                ratios['업계대비성과(%)'] = round((operating_margin / industry_avg_margin) * 100, 2)
                
        return ratios
    
    def merge_company_data(self, dataframes):
        """여러 회사 데이터 병합 (SK에너지 우선 정렬)"""
        if not dataframes:
            return pd.DataFrame()
            
        # SK에너지가 포함된 데이터프레임을 찾아서 첫 번째로 배치
        sk_df = None
        other_dfs = []
        
        for df in dataframes:
            if any(self.sk_company in col for col in df.columns):
                sk_df = df.copy()
            else:
                other_dfs.append(df)
        
        # SK에너지부터 시작하여 병합
        if sk_df is not None:
            merged = sk_df.copy()
            dataframes_to_merge = other_dfs
        else:
            merged = dataframes[0].copy()
            dataframes_to_merge = dataframes[1:]
        
        # 나머지 데이터프레임들 병합
        for df in dataframes_to_merge:
            company_cols = [col for col in df.columns if col != '구분']
            
            for company_col in company_cols:
                company_data = df.set_index('구분')[company_col]
                merged = merged.set_index('구분').join(company_data, how='outer').reset_index()
                
        merged = merged.fillna("-")
        
        # 컬럼 순서 재정렬 (SK에너지 첫 번째)
        cols = ['구분']
        sk_cols = [col for col in merged.columns if self.sk_company in col]
        competitor_cols = [col for col in merged.columns if col not in cols + sk_cols]
        
        merged = merged[cols + sk_cols + competitor_cols]
        
        return merged
    
    def apply_comparison_metric(self, merged_df, comparison_metric):
        """동적 비교 기준 적용"""
        if comparison_metric == "절대값":
            return merged_df
        elif comparison_metric == "매출 대비 비율":
            ratio_rows = merged_df[merged_df['구분'].str.contains('%|점|억원', na=False)]
            return ratio_rows
        elif comparison_metric == "업계 평균 대비":
            industry_rows = merged_df[merged_df['구분'].str.contains('업계대비|종합수익성|원가효율성', na=False)]
            return industry_rows
        elif comparison_metric == "규모 정규화":
            normalized_rows = merged_df[merged_df['구분'].str.contains('률|점|1조원당', na=False)]
            return normalized_rows
        
        return merged_df

# ==========================
# 구글시트 연동 클래스 (Make 자동화 연동)
# ==========================
class GoogleSheetsConnector:
    """Make에서 처리된 데이터를 구글시트에서 로드"""
    
    def __init__(self, credentials_path=None):
        self.credentials_path = credentials_path
        self.client = None
        
    def connect(self):
        """구글시트 연결"""
        if not GOOGLE_SHEETS_AVAILABLE:
            return False
            
        try:
            if self.credentials_path:
                creds = service_account.Credentials.from_service_account_file(
                    self.credentials_path,
                    scopes=['https://www.googleapis.com/auth/spreadsheets.readonly']
                )
            else:
                # 환경변수에서 자격증명 로드
                creds = Credentials.from_service_account_info(
                    json.loads(os.getenv('GOOGLE_CREDENTIALS', '{}')),
                    scopes=['https://www.googleapis.com/auth/spreadsheets.readonly']
                )
            
            self.client = gspread.authorize(creds)
            return True
        except Exception as e:
            st.warning(f"구글시트 연결 실패: {e}")
            return False
    
    def load_news_analysis_data(self, spreadsheet_id, worksheet_name="NewsAnalysis"):
        """Make에서 처리된 뉴스 분석 데이터 로드"""
        if not self.client:
            return None
            
        try:
            sheet = self.client.open_by_key(spreadsheet_id).worksheet(worksheet_name)
            data = sheet.get_all_records()
            
            if data:
                df = pd.DataFrame(data)
                return df
            return None
        except Exception as e:
            st.error(f"뉴스 분석 데이터 로드 실패: {e}")
            return None
    
    def load_improvement_ideas(self, spreadsheet_id, worksheet_name="ImprovementIdeas"):
        """손익 개선 아이디어 데이터 로드"""
        if not self.client:
            return None
            
        try:
            sheet = self.client.open_by_key(spreadsheet_id).worksheet(worksheet_name)
            data = sheet.get_all_records()
            
            if data:
                df = pd.DataFrame(data)
                return df
            return None
        except Exception as e:
            st.error(f"개선 아이디어 데이터 로드 실패: {e}")
            return None

# ==========================
# 외부 자료 기반 손익 개선 아이디어 분석
# ==========================
def analyze_external_improvement_ideas(news_df, financial_data):
    """외부 자료 기반 손익 개선 아이디어 도출"""
    if news_df is None or news_df.empty:
        return []
    
    improvement_ideas = []
    
    # 뉴스 데이터에서 키워드 분석
    keywords = ['원가절감', '효율성', '수익성', '마진', '비용', '투자', '혁신']
    
    for _, row in news_df.iterrows():
        content = str(row.get('content', ''))
        title = str(row.get('title', ''))
        
        for keyword in keywords:
            if keyword in content or keyword in title:
                idea = {
                    'source': '뉴스분석',
                    'keyword': keyword,
                    'title': title,
                    'impact': '높음' if keyword in ['원가절감', '효율성'] else '중간',
                    'implementation': '단기' if keyword in ['비용', '마진'] else '중장기'
                }
                improvement_ideas.append(idea)
    
    return improvement_ideas

# ==========================
# DART 출처 테이블 생성 함수 (개선된 버전)
# ==========================
def create_dart_source_table(dart_collector, collected_companies, analysis_year):
    """🔗 DART 보고서 출처 표 생성 (HTML 렌더링 수정)"""
    if not dart_collector or not dart_collector.source_tracking:
        return ""
        
    # HTML 직접 생성 대신 st.dataframe 사용
    source_data = []
    
    for company in collected_companies:
        source_info = dart_collector.source_tracking.get(company, {})
        
        report_type = source_info.get('report_type', '재무제표')
        rcept_no = source_info.get('rcept_no', 'N/A')
        
        # 링크 생성
        if rcept_no and rcept_no != 'N/A':
            dart_url = f"https://dart.fss.or.kr/dsaf001/main.do?rcpNo={rcept_no}"
        else:
            dart_url = "https://dart.fss.or.kr"
        
        source_data.append({
            '회사명': company,
            '보고서종류': report_type,
            '연도': analysis_year,
            'DART링크': dart_url
        })
    
    if source_data:
        df = pd.DataFrame(source_data)
        
        # Streamlit 표로 표시
        st.subheader("📊 DART 전자공시시스템 출처")
        st.dataframe(
            df,
            use_container_width=True,
            column_config={
                "DART링크": st.column_config.LinkColumn(
                    "🔗 DART 바로가기",
                    help="클릭하면 해당 보고서로 이동합니다",
                    validate="^https://dart\.fss\.or\.kr.*",
                    max_chars=50,
                    display_text="🔗 보기"
                )
            }
        )
        
        st.caption("💡 금융감독원 전자공시시스템(https://dart.fss.or.kr)에서 제공하는 공식 재무제표 데이터입니다.")
        
    return ""  # HTML 반환하지 않음

# ==========================
# SK 테마 RSS 뉴스 수집 클래스
# ==========================
class SKNewsRSSCollector:
    def __init__(self):
        self.rss_feeds = {
            '연합뉴스_경제': 'https://www.yna.co.kr/rss/economy.xml',
            '조선일보_경제': 'https://www.chosun.com/arc/outboundfeeds/rss/category/economy/',
            '한국경제': 'https://www.hankyung.com/feed/economy',
            '서울경제': 'https://www.sedaily.com/RSSFeed.xml',
            '매일경제': 'https://www.mk.co.kr/rss/30000001/',
            '이데일리': 'https://www.edaily.co.kr/rss/rss_economy.xml',
            '아시아경제': 'https://rss.asiae.co.kr/economy.xml',
            '파이낸셜뉴스': 'https://www.fnnews.com/rss/fn_realestate_all.xml'
        }
        
        # 정유업계 키워드 (SK 강조)
        self.oil_keywords = [
            # SK 관련 (높은 가중치)
            'SK', 'SK에너지', 'SK이노베이션', 'SK그룹', 'SK온', 
            # 경쟁사
            'GS칼텍스', 'HD현대오일', '현대오일뱅크', 'S-Oil', '에쓰오일',
            # 업종 관련
            '정유', '유가', '원유', '석유', '화학', '에너지', '나프타', '휘발유', '경유', '등유', '중유',
            '석유화학', '정제', '주유소', '정제마진', '크래킹스프레드', '두바이유', 'WTI', '브렌트유',
            # 재무 관련
            '영업이익', '매출', '수익', '실적', '손실', '적자', '흑자', '이익', '수익성', '마진',
            '매출액', '원가', '비용', '투자', '설비', '공장', '생산', '가동', '정비', '보수', '중단',
            # 시장 관련
            '국제유가', '업황', '수요', '공급', '재고', '가격', '시장', '경쟁', '점유율',
            # 정책/환경 관련
            'ESG', '탄소중립', '친환경', '바이오연료', '수소', '신재생에너지', '환경규제'
        ]
    
    def collect_real_korean_news(self, business_type='정유'):
        """실제 RSS 뉴스 수집 (SK 중심)"""
        keywords = self.oil_keywords
        all_news = []
        
        st.info(f"📡 SK에너지 중심 {business_type} 관련 뉴스 수집 중...")
        
        progress_bar = st.progress(0)
        total_feeds = len(self.rss_feeds)
        news_count = 0
        
        for idx, (source_name, rss_url) in enumerate(self.rss_feeds.items()):
            try:
                progress_bar.progress((idx + 1) / total_feeds)
                st.write(f"🔍 {source_name}에서 수집 중...")
                
                feed = feedparser.parse(rss_url)
                
                if hasattr(feed, 'entries') and len(feed.entries) > 0:
                    found_articles = 0
                    
                    for entry in feed.entries[:30]:
                        title = entry.get('title', '')
                        link = entry.get('link', '')
                        published = entry.get('published', '')
                        summary = entry.get('summary', entry.get('description', ''))
                        
                        content = f"{title} {summary}".lower()
                        matched_keywords = []
                        sk_relevance = 0
                        
                        for kw in keywords:
                            if kw.lower() in content:
                                matched_keywords.append(kw)
                                # SK 관련 키워드는 높은 점수
                                if 'sk' in kw.lower():
                                    sk_relevance += 3
                                else:
                                    sk_relevance += 1
                        
                        if matched_keywords:
                            category = self._classify_category(content)
                            company = self._extract_company(content)
                            
                            all_news.append({
                                '날짜': self._format_date(published),
                                '회사': company,
                                '제목': title,
                                '카테고리': category,
                                '키워드': ', '.join(matched_keywords[:5]),
                                '영향도': min(sk_relevance, 10),
                                'SK관련도': sk_relevance,
                                'URL': link,
                                '요약': summary[:200] + "..." if len(summary) > 200 else summary
                            })
                            found_articles += 1
                    
                    st.write(f"✅ {source_name}: {found_articles}개 관련 기사 발견")
                    news_count += found_articles
                else:
                    st.write(f"❌ {source_name}: RSS 데이터 없음")
                    
            except Exception as e:
                st.write(f"❌ {source_name} 수집 오류: {e}")
        
        progress_bar.progress(1.0)
        st.success(f"🎉 총 {news_count}개의 관련 뉴스 수집 완료!")
        
        if all_news:
            df = pd.DataFrame(all_news)
            df = df.drop_duplicates(subset=['제목'], keep='first')
            # SK 관련도 순으로 정렬
            df = df.sort_values(['SK관련도', '영향도'], ascending=[False, False])
            st.info(f"📋 중복 제거 후 {len(df)}개 뉴스 최종 선별 (SK 관련도 우선 정렬)")
            return df
        else:
            return pd.DataFrame()
    
    def create_sk_wordcloud(self, news_df):
        """SK 테마 워드클라우드 생성 (한글 폰트 개선)"""
        if news_df.empty or not WORDCLOUD_AVAILABLE:
            return None
        
        # 모든 키워드 수집
        all_keywords = []
        for keywords in news_df['키워드']:
            if pd.notna(keywords):
                all_keywords.extend([kw.strip() for kw in keywords.split(',')])
        
        # 키워드 빈도 계산
        keyword_freq = {}
        for keyword in all_keywords:
            if keyword:
                # SK 관련 키워드는 가중치 부여
                weight = 3 if 'sk' in keyword.lower() else 1
                keyword_freq[keyword] = keyword_freq.get(keyword, 0) + weight
        
        if not keyword_freq:
            return None
        
        # 한글 폰트 경로 찾기
        font_paths = [
            'C:/Windows/Fonts/malgun.ttf',      # 맑은 고딕
            'C:/Windows/Fonts/gulim.ttc',       # 굴림
            'C:/Windows/Fonts/batang.ttc',      # 바탕
            '/System/Library/Fonts/AppleGothic.ttf',  # Mac
            '/usr/share/fonts/truetype/nanum/NanumGothic.ttf'  # Linux
        ]
        
        korean_font = None
        for font_path in font_paths:
            if os.path.exists(font_path):
                korean_font = font_path
                break
        
        # 워드클라우드 생성
        plt.figure(figsize=(12, 8))
        
        wordcloud_params = {
            'width': 800, 
            'height': 400,
            'background_color': 'white',
            'max_words': 100,
            'relative_scaling': 0.5,
            'colormap': 'Reds'  # SK 레드 계열
        }
        
        # 한글 폰트가 있는 경우만 적용
        if korean_font:
            wordcloud_params['font_path'] = korean_font
        
        try:
            wordcloud = WordCloud(**wordcloud_params).generate_from_frequencies(keyword_freq)
            
            plt.imshow(wordcloud, interpolation='bilinear')
            plt.axis('off')
            
            # 한글 폰트로 제목 설정
            if korean_font:
                font_prop = fm.FontProperties(fname=korean_font)
                plt.title('SK에너지 관련 뉴스 키워드 분석', fontsize=16, pad=20, fontproperties=font_prop)
            else:
                plt.title('SK Energy Related News Keywords', fontsize=16, pad=20)
            
            # 이미지를 바이트로 변환
            img_buffer = io.BytesIO()
            plt.savefig(img_buffer, format='png', bbox_inches='tight', dpi=300)
            img_buffer.seek(0)
            plt.close()
            
            return img_buffer
            
        except Exception as e:
            st.warning(f"워드클라우드 생성 중 오류: {e}")
            return None
    
    def _classify_category(self, content):
        """카테고리 분류"""
        sk_keywords = ['sk', 'sk에너지', 'sk이노베이션']
        cost_keywords = ['보수', '중단', '유가상승', '비용', '원가', '손실', '적자', '폐기', '수율저하', '정비']
        revenue_keywords = ['영업이익', '매출', '수익', '흑자', '출하량', '납품', '계약', '증가', '개선', '성장']
        strategy_keywords = ['투자', '설비', '공장', '자동화', '디지털', 'ESG', '개발', '전환', '확장', '진출']
        
        if any(kw in content for kw in sk_keywords):
            return 'SK관련'
        elif any(kw in content for kw in cost_keywords):
            return '비용절감'
        elif any(kw in content for kw in revenue_keywords):
            return '수익개선'
        elif any(kw in content for kw in strategy_keywords):
            return '전략변화'
        else:
            return '외부환경'
    
    def _extract_company(self, content):
        """회사명 추출 (SK 우선)"""
        companies = ['SK에너지', 'SK이노베이션', 'SK', 'GS칼텍스', 'HD현대오일뱅크', 'S-Oil', '에쓰오일']
        
        for company in companies:
            if company.lower() in content:
                return company
        return '기타'
    
    def _format_date(self, date_str):
        """날짜 형식 통일"""
        try:
            from dateutil import parser
            dt = parser.parse(date_str)
            return dt.strftime('%Y-%m-%d %H:%M')
        except:
            return datetime.now().strftime('%Y-%m-%d %H:%M')
    
    def create_keyword_analysis(self, news_df):
        """키워드 분석 차트 생성 (SK 테마)"""
        if news_df.empty or not PLOTLY_AVAILABLE:
            return None
        
        category_counts = news_df['카테고리'].value_counts()
        
        # SK 테마 색상 적용
        colors = [SK_COLORS['primary'] if 'SK' in cat else SK_COLORS['competitor'] 
                 for cat in category_counts.index]
        
        fig = go.Figure(data=[go.Pie(
            labels=category_counts.index,
            values=category_counts.values,
            hole=0.4,
            marker_colors=colors
        )])
        
        fig.update_layout(
            title="📊 SK에너지 중심 뉴스 카테고리 분포",
            height=400,
            title_font_size=18,
            font_size=14
        )
        
        return fig

# ==========================
# XBRL 파서 클래스 (오류 해결)
# ==========================
class EnhancedXBRLParser:
    def __init__(self):
        self.namespaces = {
            'xbrli': 'http://www.xbrl.org/2003/instance',
            'link': 'http://www.xbrl.org/2003/linkbase',
            'xlink': 'http://www.w3.org/1999/xlink',
            'ifrs': 'http://xbrl.ifrs.org/taxonomy/2018-03-16/ifrs',
            'dart': 'http://dart.fss.or.kr'
        }
    
    def parse_xbrl_file(self, file_content, filename):
        """XBRL 파일 파싱"""
        try:
            # 인코딩 처리
            content_str = self._safe_decode(file_content, filename)
            if not content_str:
                return None
            
            # XML 파싱
            root = self._safe_xml_parse(content_str, filename)
            if root is None:
                return None
            
            # 회사 정보 및 재무 데이터 추출
            company_info = self._extract_company_info_enhanced(root, content_str)
            financial_data = self._extract_financial_data_enhanced(root, content_str)
            
            if financial_data:
                st.success(f"✅ {filename}: 재무 데이터 추출 성공")
                return {
                    'company_name': company_info.get('name', filename.split('.')[0]),
                    'period': company_info.get('period', '2024'),
                    'data': financial_data
                }
            else:
                # 더 관대한 처리 - 최소한의 데이터라도 추출
                st.warning(f"⚠️ {filename}: 표준 추출 실패, 대안 방법 시도")
                return self._try_alternative_parsing(root, content_str, filename)
                
        except Exception as e:
            st.error(f"❌ XBRL 파싱 오류 ({filename}): {e}")
            return None
    
    def _safe_decode(self, file_content, filename):
        """안전한 인코딩 처리"""
        encodings = ['utf-8', 'euc-kr', 'cp949', 'latin-1', 'utf-16']
        
        for encoding in encodings:
            try:
                content_str = file_content.decode(encoding)
                st.write(f"✅ {filename}: {encoding} 인코딩 성공")
                return content_str
            except:
                continue
        
        st.error(f"❌ {filename}: 인코딩 처리 실패")
        return None
    
    def _safe_xml_parse(self, content_str, filename):
        """안전한 XML 파싱"""
        try:
            return ET.fromstring(content_str)
        except ET.ParseError as e:
            st.warning(f"⚠️ {filename}: 표준 XML 파싱 실패, BeautifulSoup 시도")
            try:
                from bs4 import BeautifulSoup
                soup = BeautifulSoup(content_str, 'xml')
                return ET.fromstring(str(soup))
            except Exception as e2:
                st.error(f"❌ {filename}: XML 파싱 실패")
                return None
    
    def _try_alternative_parsing(self, root, content_str, filename):
        """대안적 파싱 방법"""
        financial_items = {}
        
        # 텍스트에서 직접 숫자 패턴 찾기 (정규표현식 사용)
        patterns = {
            '매출액': [r'매출액[:\s]*([0-9,\-()]+)', r'revenue[:\s]*([0-9,\-()]+)', r'sales[:\s]*([0-9,\-()]+)'],
            '영업이익': [r'영업이익[:\s]*([0-9,\-()]+)', r'operating[:\s]*([0-9,\-()]+)'],
            '당기순이익': [r'당기순이익[:\s]*([0-9,\-()]+)', r'net.*income[:\s]*([0-9,\-()]+)']
        }
        
        for category, pattern_list in patterns.items():
            for pattern in pattern_list:
                try:
                    matches = re.findall(pattern, content_str, re.IGNORECASE)
                    if matches:
                        # 가장 큰 값 선택
                        values = []
                        for match in matches:
                            try:
                                # 📉 마이너스 값 처리
                                clean_match = match.replace(',', '')
                                if '(' in clean_match and ')' in clean_match:
                                    clean_match = '-' + clean_match.replace('(', '').replace(')', '')
                                
                                if clean_match.replace('-', '').isdigit():
                                    values.append(float(clean_match))
                            except:
                                continue
                        
                        if values:
                            financial_items[category] = max(values, key=abs)  # 절대값이 가장 큰 값
                            break
                except Exception as e:
                    continue
        
        if financial_items:
            st.success(f"✅ {filename}: 대안 방법으로 데이터 추출 성공")
            return {
                'company_name': filename.split('.')[0],
                'period': '2024',
                'data': financial_items
            }
        
        st.warning(f"⚠️ {filename}: 데이터 추출 실패")
        return None
    
    def _extract_financial_data_enhanced(self, root, content_str):
        """재무 데이터 추출"""
        financial_items = {}
        
        # 재무 항목 패턴
        financial_patterns = {
            '매출액': [
                'revenue', 'sales', 'income', '매출액', '수익', 'Revenue', 'Sales',
                'Operating revenues', 'TotalRevenues', 'Net sales', '총매출액', '매출'
            ],
            '영업이익': [
                'operating', 'operational', '영업이익', 'OperatingIncome', 
                'Operating profit', 'OperatingProfit', '영업손익', '영업수익'
            ],
            '당기순이익': [
                'netincome', 'net income', 'profit', '당기순이익', '순이익', 
                'NetIncome', 'Net profit', 'ProfitLoss', '순손익', '순수익'
            ],
            '매출원가': [
                'costofrevenue', 'cogs', 'cost of goods', '매출원가', '매출비용',
                'CostOfRevenue', 'CostOfGoodsSold', '제품매출원가', '원가'
            ],
            '총자산': [
                'totalassets', 'total assets', '총자산', '자산총계', 
                'TotalAssets', 'Assets', '자산총액', '자산'
            ]
        }
        
        # 1단계: XML 태그에서 추출
        for category, patterns in financial_patterns.items():
            for elem in root.iter():
                elem_tag = elem.tag.lower()
                elem_text = elem.text
                
                if elem_text:
                    # 숫자 정리 및 검증
                    clean_text = str(elem_text).replace(',', '').strip()
                    
                    # 📉 마이너스 값 처리
                    if '(' in clean_text and ')' in clean_text:
                        clean_text = '-' + clean_text.replace('(', '').replace(')', '')
                    
                    # 숫자인지 확인 (마이너스 포함)
                    if clean_text.replace('.', '').replace('-', '').isdigit() and len(clean_text) > 0:
                        try:
                            value = float(clean_text)
                            
                            # 태그명 패턴 매칭
                            for pattern in patterns:
                                if pattern.lower() in elem_tag:
                                    # 더 큰 절대값으로 업데이트 (중복 방지)
                                    if category not in financial_items or abs(value) > abs(financial_items[category]):
                                        financial_items[category] = value
                                    break
                            
                            if category in financial_items:
                                break
                        except:
                            continue
        
        # 2단계: 텍스트 패턴 매칭 (정규표현식)
        if len(financial_items) < 2:  # 충분한 데이터가 없으면 텍스트 검색
            for category, patterns in financial_patterns.items():
                if category in financial_items:
                    continue
                    
                for pattern in patterns:
                    # 정규표현식 패턴 개선
                    regex_patterns = [
                        pattern + r'[:\s]*([0-9,.\-()]+)',
                        r'<[^>]*' + pattern + r'[^>]*>([0-9,.\-()]+)',
                        r'\b' + pattern + r'\s*[:\-=]\s*([0-9,.\-()]+)'
                    ]
                    
                    for regex_pattern in regex_patterns:
                        try:
                            matches = re.findall(regex_pattern, content_str, re.IGNORECASE)
                            
                            if matches:
                                # 유효한 숫자값만 추출
                                valid_values = []
                                for match in matches:
                                    clean_match = str(match).replace(',', '').strip()
                                    
                                    # 📉 마이너스 값 처리
                                    if '(' in clean_match and ')' in clean_match:
                                        clean_match = '-' + clean_match.replace('(', '').replace(')', '')
                                    
                                    if clean_match.replace('.', '').replace('-', '').isdigit() and len(clean_match) > 0:
                                        try:
                                            valid_values.append(float(clean_match))
                                        except:
                                            continue
                                
                                if valid_values:
                                    # 가장 큰 절댓값 선택 (보통 최종 합계값)
                                    financial_items[category] = max(valid_values, key=abs)
                                    break
                        except Exception as e:
                            continue
                    
                    if category in financial_items:
                        break
        
        return financial_items if financial_items else None
    
    def _extract_company_info_enhanced(self, root, content_str):
        """회사 정보 추출"""
        info = {}
        
        # 회사명 추출 패턴
        name_patterns = [
            'company', 'corp', 'entity', '회사', '법인', 'name',
            'EntityRegistrantName', 'CompanyName', 'CorporationName'
        ]
        
        # XML 태그에서 찾기
        for pattern in name_patterns:
            for elem in root.iter():
                if pattern.lower() in elem.tag.lower():
                    if elem.text and elem.text.strip():
                        info['name'] = elem.text.strip()
                        break
            if 'name' in info:
                break
        
        # 텍스트에서 찾기 (정규표현식 사용)
        if 'name' not in info:
            company_patterns = [
                r'회사명[:\s]*([가-힣\w\s\(\)]+)',
                r'법인명[:\s]*([가-힣\w\s\(\)]+)',
                r'EntityRegistrantName[^>]*>([^<]+)<'
            ]
            
            for pattern in company_patterns:
                try:
                    match = re.search(pattern, content_str)
                    if match:
                        info['name'] = match.group(1).strip()
                        break
                except:
                    continue
        
        # 기간 정보
        try:
            year_match = re.search(r'20\d{2}', content_str)
            if year_match:
                info['period'] = year_match.group()
        except:
            pass
        
        return info
# ==========================
# PDF 템플릿 클래스 (개선된 버전 - 페이지 이어짐 지원)
# ==========================
class SKEnergyDocTemplate(SimpleDocTemplate):
    """📄 SK에너지 전용 PDF 템플릿 (한글 폰트 지원 + 페이지 이어짐)"""
    
    def __init__(self, filename, **kwargs):
        super().__init__(filename, **kwargs)
        self.company_name = "SK에너지"
        self.report_title = "경쟁사 분석 보고서"
        
        # 페이지 이어짐을 위한 프레임 설정
        self.frame_width = letter[0] - 100  # 좌우 여백
        self.frame_height = letter[1] - 100  # 상하 여백
        self.frame_x = 50  # 좌측 여백
        self.frame_y = 50  # 하단 여백
        
        # 메인 콘텐츠 프레임
        self.main_frame = Frame(
            self.frame_x, 
            self.frame_y, 
            self.frame_width, 
            self.frame_height,
            leftPadding=0,
            bottomPadding=0,
            rightPadding=0,
            topPadding=0
        )
        
        # 페이지 템플릿 설정
        self.page_template = PageTemplate(
            id='SKEnergyTemplate',
            frames=[self.main_frame],
            onPage=self.add_page_elements
        )
        self.addPageTemplates([self.page_template])
    
    def add_page_elements(self, canvas, doc):
        """페이지 헤더/푸터 추가 (개선된 버전)"""
        # SK 브랜드 헤더 (상단)
        canvas.setFillColor(colors.Color(227/255, 30/255, 36/255))  # SK Red
        canvas.rect(0, letter[1] - 40, letter[0], 40, fill=True, stroke=False)
        
        # 헤더 텍스트
        canvas.setFillColor(colors.white)
        canvas.setFont('Helvetica-Bold', 12)
        canvas.drawString(50, letter[1] - 25, f"SK Energy Competitive Analysis Report")
        
        # 날짜 (오른쪽 상단)
        canvas.drawRightString(letter[0] - 50, letter[1] - 25, datetime.now().strftime('%Y-%m-%d'))
        
        # 페이지 번호 (하단 중앙)
        canvas.setFillColor(colors.black)
        canvas.setFont('Helvetica', 10)
        page_num = canvas.getPageNumber()
        canvas.drawCentredString(letter[0]/2, 25, f"- {page_num} -")
        
        # DART 출처 (하단 왼쪽)
        canvas.setFont('Helvetica', 8)
        canvas.drawString(50, 10, "Data Source: DART Electronic Disclosure System")
        
        # 구분선 (헤더와 콘텐츠 사이)
        canvas.setStrokeColor(colors.Color(227/255, 30/255, 36/255))
        canvas.setLineWidth(1)
        canvas.line(50, letter[1] - 50, letter[0] - 50, letter[1] - 50)

# ==========================
# 차트 생성 함수들 (파스텔 색상 + 폰트 크기 증가)
# ==========================
def create_sk_radar_chart(chart_df):
    """SK에너지 중심 레이더 차트 (파스텔 색상 + 폰트 크기 증가)"""
    if chart_df.empty or not PLOTLY_AVAILABLE:
        return None
    
    companies = chart_df['회사'].unique()
    metrics = chart_df['지표'].unique()
    
    fig = go.Figure()
    
    for i, company in enumerate(companies):
        company_data = chart_df[chart_df['회사'] == company]
        values = company_data['수치'].tolist()
        
        if values:
            values.append(values[0])
            theta_labels = list(metrics) + [metrics[0]]
        else:
            continue
        
        # 파스텔 색상 적용
        color = get_company_color(company, companies)
        
        # SK에너지는 특별한 스타일
        if 'SK' in company:
            line_width = 5
            marker_size = 12
            name_style = f"<b>{company}</b>"  # 굵게 표시
        else:
            line_width = 3
            marker_size = 8
            name_style = company
        
        fig.add_trace(go.Scatterpolar(
            r=values,
            theta=theta_labels,
            fill='toself',
            name=name_style,
            line=dict(width=line_width, color=color),
            marker=dict(size=marker_size, color=color)
        ))
    
    max_value = chart_df['수치'].max() if not chart_df.empty else 10
    
    fig.update_layout(
        polar=dict(
            radialaxis=dict(
                visible=True,
                range=[0, max_value * 1.2],
                tickmode='linear',
                tick0=0,
                dtick=max_value * 0.2,
                tickfont=dict(size=14)
            ),
            angularaxis=dict(
                tickfont=dict(size=16)
            )
        ),
        title="🎯 SK에너지 vs 경쟁사 수익성 지표 비교",
        height=600,
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
            font=dict(size=14)
        ),
        title_font_size=20,
        font=dict(size=14)
    )
    
    return fig

def create_sk_bar_chart(chart_df):
    """SK에너지 강조 막대 차트 (파스텔 색상 + 폰트 크기 증가)"""
    if chart_df.empty or not PLOTLY_AVAILABLE:
        return None
    
    # 파스텔 색상 매핑
    companies = chart_df['회사'].unique()
    color_discrete_map = {}
    for company in companies:
        color_discrete_map[company] = get_company_color(company, companies)
    
    fig = px.bar(
        chart_df, 
        x='지표', 
        y='수치', 
        color='회사',
        title="💼 SK에너지 vs 경쟁사 수익성 지표 비교",
        height=600,
        text='수치',
        color_discrete_map=color_discrete_map
    )
    
    fig.update_traces(
        texttemplate='%{text:.2f}',
        textposition='outside',
        textfont=dict(size=14, color='black')
    )
    
    max_value = chart_df['수치'].max()
    fig.update_layout(
        yaxis=dict(
            range=[0, max_value * 1.2],
            title="수치",
            title_font_size=16,
            tickfont=dict(size=14)
        ),
        xaxis=dict(
            title="재무 지표",
            tickangle=45,
            title_font_size=16,
            tickfont=dict(size=14)
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
            font=dict(size=14)
        ),
        title_font_size=20,
        font=dict(size=14)
    )
    
    return fig

def create_sk_bubble_chart(chart_df):
    """SK 버블 차트 (파스텔 색상 + 폰트 크기 증가)"""
    if chart_df.empty or not PLOTLY_AVAILABLE:
        return None
    
    try:
        # 빈 데이터나 잘못된 데이터 처리
        if len(chart_df) < 2:
            st.warning("버블 차트를 생성하기 위해서는 최소 2개 이상의 데이터가 필요합니다.")
            return None
        
        # 수치가 0이거나 음수인 경우 처리
        chart_df_filtered = chart_df[chart_df['수치'] > 0].copy()
        
        if chart_df_filtered.empty:
            st.warning("버블 차트를 생성할 수 있는 양수 데이터가 없습니다.")
            return None
        
        # X축용 인덱스 생성
        chart_df_filtered['지표_인덱스'] = range(len(chart_df_filtered))
        
        # 파스텔 색상 매핑
        companies = chart_df_filtered['회사'].unique()
        color_discrete_map = {}
        for company in companies:
            color_discrete_map[company] = get_company_color(company, companies)
        
        fig = px.scatter(
            chart_df_filtered,
            x='지표_인덱스',
            y='수치',
            size='수치',
            color='회사',
            hover_name='지표',
            title="🔵 SK에너지 성과 버블 차트 (크기 = 값)",
            height=600,
            size_max=60,
            color_discrete_map=color_discrete_map
        )
        
        # X축 라벨을 지표명으로 변경
        fig.update_layout(
            xaxis=dict(
                tickmode='array',
                tickvals=chart_df_filtered['지표_인덱스'],
                ticktext=chart_df_filtered['지표'],
                tickangle=45,
                title="재무 지표",
                title_font_size=16,
                tickfont=dict(size=14)
            ),
            yaxis=dict(
                title="수치",
                title_font_size=16,
                tickfont=dict(size=14)
            ),
            showlegend=True,
            legend=dict(
                font=dict(size=14)
            ),
            title_font_size=20,
            font=dict(size=14)
        )
        
        return fig
        
    except Exception as e:
        st.error(f"버블 차트 생성 오류: {e}")
        return None

def create_sk_heatmap_chart(chart_df):
    """SK 히트맵 차트 (파스텔 색상 + 폰트 크기 증가)"""
    if chart_df.empty or not PLOTLY_AVAILABLE:
        return None
    
    try:
        # pivot 테이블 생성
        pivot_df = chart_df.pivot(index='지표', columns='회사', values='수치')
        
        # NaN 값 처리
        pivot_df = pivot_df.fillna(0)
        
        if pivot_df.empty:
            st.warning("히트맵을 생성할 데이터가 부족합니다.")
            return None
        
        fig = px.imshow(
            pivot_df, 
            text_auto=True,
            title="🔥 SK에너지 vs 경쟁사 수익성 히트맵",
            color_continuous_scale='RdYlGn',
            aspect="auto",
            height=600
        )
        
        # 폰트 크기 증가
        fig.update_layout(
            font=dict(size=14),
            title_font_size=20,
            xaxis=dict(
                title_font_size=16,
                tickfont=dict(size=14)
            ),
            yaxis=dict(
                title_font_size=16,
                tickfont=dict(size=14)
            )
        )
        
        # 텍스트 폰트 크기 증가
        fig.update_traces(textfont_size=12)
        
        return fig
        
    except Exception as e:
        st.error(f"히트맵 생성 오류: {e}")
        return None

def create_quarterly_trend_chart(quarterly_df):
    """분기별 추이 차트 생성 (파스텔 색상 + 폰트 크기 증가)"""
    if quarterly_df.empty or not PLOTLY_AVAILABLE:
        return None
    
    fig = go.Figure()
    
    companies = quarterly_df['회사'].unique()
    
    for company in companies:
        company_data = quarterly_df[quarterly_df['회사'] == company]
        
        # 파스텔 색상 적용
        line_color = get_company_color(company, companies)
        
        # SK에너지는 특별한 스타일
        if 'SK' in company:
            line_width = 4
            marker_size = 10
            name_style = f"<b>{company}</b>"
        else:
            line_width = 2
            marker_size = 6
            name_style = company
        
        # 매출액 추이
        if '매출액' in company_data.columns:
            fig.add_trace(go.Scatter(
                x=company_data['분기'],
                y=company_data['매출액'],
                mode='lines+markers',
                name=f"{name_style} 매출액",
                line=dict(color=line_color, width=line_width),
                marker=dict(size=marker_size, color=line_color),
                yaxis='y'
            ))
        
        # 영업이익률 추이 (보조 축)
        if '영업이익률' in company_data.columns:
            fig.add_trace(go.Scatter(
                x=company_data['분기'],
                y=company_data['영업이익률'],
                mode='lines+markers',
                name=f"{name_style} 영업이익률",
                line=dict(color=line_color, width=line_width, dash='dash'),
                marker=dict(size=marker_size, color=line_color, symbol='diamond'),
                yaxis='y2'
            ))
    
    fig.update_layout(
        title="📈 분기별 재무성과 추이 분석 (SK에너지 vs 경쟁사)",
        xaxis=dict(
            title="분기",
            title_font_size=16,
            tickfont=dict(size=14)
        ),
        yaxis=dict(
            title="매출액 (조원)", 
            side="left",
            title_font_size=16,
            tickfont=dict(size=14)
        ),
        yaxis2=dict(
            title="영업이익률 (%)", 
            side="right", 
            overlaying="y",
            title_font_size=16,
            tickfont=dict(size=14)
        ),
        height=600,
        hovermode='x unified',
        legend=dict(
            font=dict(size=14)
        ),
        title_font_size=20,
        font=dict(size=14)
    )
    
    return fig

# ==========================
# SK 중심 인사이트 테이블 생성
# ==========================
def create_sk_insight_table(merged_df, collected_companies):
    """SK에너지 중심 인사이트 테이블"""
    if merged_df.empty:
        return "<p>분석할 데이터가 없습니다.</p>"
    
    # 주요 지표 추출
    key_metrics = ['영업이익률(%)', '순이익률(%)', '매출원가율(%)', '종합수익성점수(점)']
    table_data = []
    
    for metric in key_metrics:
        metric_row = merged_df[merged_df['구분'] == metric]
        if not metric_row.empty:
            row_data = {'지표': metric}
            
            # SK에너지를 첫 번째로 배치
            sk_companies = [comp for comp in collected_companies if 'SK' in comp]
            other_companies = [comp for comp in collected_companies if 'SK' not in comp]
            ordered_companies = sk_companies + other_companies
            
            # 각 회사별 데이터 수집
            for company in ordered_companies:
                for col in metric_row.columns:
                    if company in col:
                        value = str(metric_row[col].iloc[0])
                        if value != "-":
                            if '%' in value or '점' in value:
                                try:
                                    num_val = float(value.replace('%', '').replace('점', ''))
                                    formatted_val = f"{num_val:.2f}%" if '%' in value else f"{num_val:.2f}점"
                                except:
                                    formatted_val = value
                            else:
                                formatted_val = value
                            row_data[company] = formatted_val
                        else:
                            row_data[company] = "N/A"
                        break
                else:
                    row_data[company] = "N/A"
            
            # 모든 지표에 대해 업계평균과 분석결과 설정
            if '영업이익률' in metric:
                row_data['업계평균'] = "3.50%"
                row_data['분석결과'] = "SK 영업수익성 분석완료"
            elif '순이익률' in metric:
                row_data['업계평균'] = "2.50%"
                row_data['분석결과'] = "SK 순수익성 분석완료"
            elif '매출원가율' in metric:
                row_data['업계평균'] = "95.00%"
                row_data['분석결과'] = "SK 원가효율성 분석완료"
            elif '종합수익성' in metric:
                row_data['업계평균'] = "3.00점"
                row_data['분석결과'] = "SK 종합성과 분석완료"
            else:
                row_data['업계평균'] = "N/A"
                row_data['분석결과'] = "분석중"
            
            table_data.append(row_data)
    
    if not table_data:
        return "<p>분석 가능한 지표가 없습니다.</p>"
    
    # HTML 변수 초기화
    html = f"""
    <div style='font-family: "Noto Sans KR", "Malgun Gothic", "맑은 고딕", sans-serif; margin: 20px 0;'>
        <h3 style='color: {SK_COLORS["primary"]}; margin-bottom: 20px;'>📈 SK에너지 중심 핵심 재무지표 비교표</h3>
        <table class='analysis-table'>
            <thead>
                <tr>
                    <th style='min-width: 150px;'>지표</th>
    """
    
    # SK에너지를 첫 번째로 배치한 헤더
    sk_companies = [comp for comp in collected_companies if 'SK' in comp]
    other_companies = [comp for comp in collected_companies if 'SK' not in comp]
    ordered_companies = sk_companies + other_companies
    
    for company in ordered_companies:
        if 'SK' in company:
            html += f"<th style='min-width: 120px; background: linear-gradient(135deg, {SK_COLORS['primary']} 0%, {SK_COLORS['secondary']} 100%);'><b>{company}</b></th>"
        else:
            html += f"<th style='min-width: 120px;'>{company}</th>"
    
    html += """
                    <th style='min-width: 100px;'>업계평균</th>
                    <th style='min-width: 150px;'>분석결과</th>
                </tr>
            </thead>
            <tbody>
    """
    
    # 데이터 행 추가 (SK 강조)
    for row_data in table_data:
        html += "<tr>"
        html += f"<td style='font-weight: 600; text-align: left;'>{row_data['지표']}</td>"
        
        for company in ordered_companies:
            value = row_data.get(company, "N/A")
            
            # SK 데이터는 특별 강조
            if 'SK' in company:
                html += f"<td class='sk-data-cell'><b>{value}</b></td>"
            else:
                html += f"<td>{value}</td>"
        
        # 안전한 접근으로 KeyError 방지
        html += f"<td style='color: #6c757d;'>{row_data.get('업계평균', 'N/A')}</td>"
        html += f"<td style='font-weight: 600; color: {SK_COLORS['primary']};'>{row_data.get('분석결과', '분석중')}</td>"
        html += "</tr>"
    
    html += """
            </tbody>
        </table>
    </div>
    """
    
    return html

# ==========================
# 한글 깨짐 방지 테이블 생성 함수
# ==========================
def create_sk_korean_table(df):
    """SK 테마 한글 테이블 생성"""
    if df.empty:
        return "<p>데이터가 없습니다.</p>"
    
    html = "<div style='font-family: \"Noto Sans KR\", \"Malgun Gothic\", \"맑은 고딕\", sans-serif; overflow-x: auto; margin: 20px 0;'>"
    html += "<table style='width:100%; border-collapse: collapse; background: white; box-shadow: 0 2px 4px rgba(0,0,0,0.1);'>"
    
    # 헤더 (SK 테마 색상)
    html += f"<thead><tr style='background: linear-gradient(135deg, {SK_COLORS['primary']} 0%, {SK_COLORS['secondary']} 100%); color: white;'>"
    for col in df.columns:
        if 'SK' in col:
            html += f"<th style='border: 1px solid #ddd; padding: 15px; text-align: center; font-weight: 700; font-size: 16px;'><b>{col}</b></th>"
        else:
            html += f"<th style='border: 1px solid #ddd; padding: 15px; text-align: center; font-weight: 600;'>{col}</th>"
    html += "</tr></thead>"
    
    # 데이터 행
    html += "<tbody>"
    for idx, row in df.iterrows():
        bg_color = "#f8f9fa" if idx % 2 == 0 else "white"
        html += f"<tr style='background-color: {bg_color}; transition: background-color 0.2s;' onmouseover='this.style.backgroundColor=\"#e3f2fd\"' onmouseout='this.style.backgroundColor=\"{bg_color}\"'>"
        for i, val in enumerate(row):
            align = "left" if i == 0 else "right"
            font_weight = "600" if i == 0 else "400"
            
            # SK 관련 데이터는 특별 강조
            if i > 0 and 'SK' in df.columns[i]:
                # 손실 표시 적용
                if isinstance(val, str) and '영업손실' in val:
                    html += f"<td class='sk-data-cell loss-indicator' style='border: 1px solid #ddd; padding: 12px; text-align: {align}; font-weight: 700; font-size: 15px;'><b>{val}</b></td>"
                else:
                    html += f"<td class='sk-data-cell' style='border: 1px solid #ddd; padding: 12px; text-align: {align}; font-weight: 700; font-size: 15px;'><b>{val}</b></td>"
            else:
                html += f"<td style='border: 1px solid #ddd; padding: 12px; text-align: {align}; font-weight: {font_weight};'>{val}</td>"
        html += "</tr>"
    html += "</tbody></table></div>"
    return html

# ==========================
# PDF 생성 함수 (한글 폰트 강제 등록)
# ==========================
def create_enhanced_pdf_report(merged_df, collected_companies, analysis_year, chart_images=None):
    """📄 한글 폰트 강제 등록 PDF 보고서 생성 (차트 포함)"""
    if not PDF_AVAILABLE:
        return None
        
    buffer = io.BytesIO()
    
    try:
        # 한글 폰트 강제 등록
        font_paths = [
            "C:/Windows/Fonts/malgun.ttf",     # 맑은 고딕
            "C:/Windows/Fonts/malgunbd.ttf",   # 맑은 고딕 Bold
            "C:/Windows/Fonts/gulim.ttc", 
            "/System/Library/Fonts/AppleGothic.ttf",
        ]
        
        korean_font_registered = False
        bold_font_registered = False
        
        # 일반 폰트 등록
        for font_path in font_paths:
            if 'malgunbd' not in font_path and os.path.exists(font_path):
                try:
                    pdfmetrics.registerFont(TTFont('KoreanFont', font_path))
                    korean_font_registered = True
                    break
                except:
                    continue
        
        # Bold 폰트 등록
        bold_paths = ["C:/Windows/Fonts/malgunbd.ttf"]
        for bold_path in bold_paths:
            if os.path.exists(bold_path):
                try:
                    pdfmetrics.registerFont(TTFont('KoreanFont-Bold', bold_path))
                    bold_font_registered = True
                    break
                except:
                    continue
        
        # Bold 폰트가 없으면 일반 폰트를 Bold로 사용
        if korean_font_registered and not bold_font_registered:
            for font_path in font_paths:
                if 'malgunbd' not in font_path and os.path.exists(font_path):
                    try:
                        pdfmetrics.registerFont(TTFont('KoreanFont-Bold', font_path))
                        bold_font_registered = True
                        break
                    except:
                        continue
                        
    except Exception as e:
        korean_font_registered = False
        bold_font_registered = False
    
    # SK에너지 전용 PDF 템플릿 사용
    doc = SKEnergyDocTemplate(buffer, pagesize=letter)
    styles = getSampleStyleSheet()
    story = []
    
    # 한글 폰트 스타일 적용
    sk_title_style = ParagraphStyle(
        'SKTitleStyle',
        parent=styles['Heading1'],
        fontSize=24,
        spaceAfter=20,
        spaceBefore=10,
        alignment=1,
        textColor=colors.Color(227/255, 30/255, 36/255),
        fontName='KoreanFont-Bold' if bold_font_registered else 'Helvetica-Bold'
    )
    
    sk_heading_style = ParagraphStyle(
        'SKHeadingStyle',
        parent=styles['Heading2'],
        fontSize=18,
        spaceBefore=15,
        spaceAfter=10,
        textColor=colors.Color(227/255, 30/255, 36/255),
        fontName='KoreanFont-Bold' if bold_font_registered else 'Helvetica-Bold'
    )
    
    normal_style = ParagraphStyle(
        'KoreanNormal',
        parent=styles['Normal'],
        fontSize=11,
        leading=14,
        spaceAfter=8,
        fontName='KoreanFont' if korean_font_registered else 'Helvetica'
    )
    
    # 제목
    story.append(Paragraph("<b>SK에너지 경쟁사 분석 보고서</b>", sk_title_style))
    
    # 요약 정보
    summary_text = f"""
    <b>분석 대상:</b> SK에너지 vs {', '.join([c for c in collected_companies if 'SK' not in c])}<br/>
    <b>분석 연도:</b> {analysis_year}년 | <b>생성일시:</b> {datetime.now().strftime('%Y-%m-%d %H:%M')}<br/>
    <b>분석 목적:</b> SK에너지의 경쟁우위 분석 및 전략적 인사이트 도출<br/>
    """
    story.append(Paragraph(summary_text, normal_style))
    story.append(Spacer(1, 15))
    
    # 상세 재무 데이터 테이블
    story.append(Paragraph("<b>상세 재무분석 데이터</b>", sk_heading_style))
    
    if not merged_df.empty:
        # 테이블 데이터 준비
        table_data = []
        
        # 헤더 행
        header_row = ['구분']
        sk_companies = [comp for comp in collected_companies if 'SK' in comp]
        other_companies = [comp for comp in collected_companies if 'SK' not in comp]
        ordered_companies = sk_companies + other_companies
        header_row.extend(ordered_companies)
        table_data.append(header_row)
        
        # 우선 지표 추가
        priority_metrics = [
            '매출 1조원당 영업이익(억원)',
            '매출원가율(%)',
            '순이익률(%)',
            '업계대비성과(%)',
            '영업이익률(%)',
            '원가효율성지수(점)',
            '종합수익성점수(점)',
            '판관비율(%)'
        ]
        
        # 우선 지표 추가
        for metric in priority_metrics:
            metric_row = merged_df[merged_df['구분'] == metric]
            if not metric_row.empty:
                row_data = [metric]
                for company in ordered_companies:
                    for col in metric_row.columns:
                        if company in col:
                            value = str(metric_row[col].iloc[0])
                            row_data.append(value)
                            break
                    else:
                        row_data.append("-")
                table_data.append(row_data)
        
        # PDF 테이블 생성
        if len(table_data) > 1:
            detail_table = Table(table_data, repeatRows=1)
            
            # SK 컬럼 찾기
            sk_col_index = None
            for i, company in enumerate(ordered_companies):
                if 'SK' in company:
                    sk_col_index = i + 1
                    break
            
            # 테이블 스타일 적용
            table_style = [
                ('BACKGROUND', (0, 0), (-1, 0), colors.Color(227/255, 30/255, 36/255)),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, -1), 'KoreanFont' if korean_font_registered else 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, 0), 10),
                ('FONTSIZE', (0, 1), (-1, -1), 9),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('ALIGN', (0, 1), (0, -1), 'LEFT'),
            ]
            
            # SK 컬럼 특별 강조
            if sk_col_index is not None:
                table_style.extend([
                    ('BACKGROUND', (sk_col_index, 0), (sk_col_index, 0), colors.Color(180/255, 20/255, 30/255)),
                    ('BACKGROUND', (sk_col_index, 1), (sk_col_index, -1), colors.Color(227/255, 30/255, 36/255, 0.1)),
                    ('TEXTCOLOR', (sk_col_index, 1), (sk_col_index, -1), colors.Color(227/255, 30/255, 36/255)),
                    ('FONTSIZE', (sk_col_index, 1), (sk_col_index, -1), 10),
                    ('FONTNAME', (sk_col_index, 1), (sk_col_index, -1), 'KoreanFont-Bold' if bold_font_registered else 'Helvetica-Bold'),
                ])
            
            detail_table.setStyle(TableStyle(table_style))
            story.append(detail_table)
    
    story.append(Spacer(1, 15))
    
    # 차트 이미지 추가 (새로운 기능 - 페이지 이어짐 지원)
    if chart_images and len(chart_images) > 0:
        story.append(Paragraph("<b>📈 시각화 분석</b>", sk_heading_style))
        
        for i, (chart_name, chart_bytes) in enumerate(chart_images.items()):
            if chart_bytes:
                try:
                    # 차트 이미지를 PDF에 추가 (페이지 이어짐 고려)
                    img_buffer = io.BytesIO(chart_bytes)
                    img = Image(img_buffer, width=5.5*inch, height=3.5*inch)  # 크기 조정
                    story.append(img)
                    story.append(Spacer(1, 8))
                    
                    # 차트 설명 추가
                    chart_desc = f"<i>{chart_name}</i>"
                    story.append(Paragraph(chart_desc, normal_style))
                    
                    # 마지막 차트가 아니면 페이지 구분 추가
                    if i < len(chart_images) - 1:
                        story.append(Spacer(1, 10))
                    else:
                        story.append(Spacer(1, 15))
                    
                except Exception as e:
                    # 차트 추가 실패 시 설명만 추가
                    chart_desc = f"<i>{chart_name} (이미지 로드 실패)</i>"
                    story.append(Paragraph(chart_desc, normal_style))
                    story.append(Spacer(1, 10))
    
    # DART 출처 정보 추가
    story.append(Paragraph("<b>데이터 출처</b>", sk_heading_style))
    
    dart_source_text = """
    <b>DART 전자공시시스템 출처:</b><br/>
    """
    
    for company in collected_companies:
        dart_source_text += f"• {company}: DART 전자공시시스템 - {analysis_year}년 재무제표<br/>"
    
    dart_source_text += f"""<br/>
    <b>수집 기준:</b> 금융감독원 전자공시시스템 (https://dart.fss.or.kr)<br/>
    <b>분석 일시:</b> {datetime.now().strftime('%Y년 %m월 %d일 %H시 %M분')}<br/>
    <b>보고서 타입:</b> 사업보고서, 분기보고서 등 최신 공시자료 기준<br/>
    """
    
    story.append(Paragraph(dart_source_text, normal_style))
    story.append(Spacer(1, 15))
    
    # PDF 생성
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

def save_chart_as_image(fig, filename):
    """차트를 이미지로 저장"""
    if fig is None:
        return None
    
    try:
        img_bytes = fig.to_image(format="png", width=800, height=600)
        return img_bytes
    except Exception as e:
        st.warning(f"차트 이미지 저장 실패: {e}")
        return None

# ==========================
# 파일 생성 함수들
# ==========================
def create_excel_file(merged_df, collected_companies, analysis_year):
    """Excel 파일 생성 (한글 지원)"""
    excel_buffer = io.BytesIO()
    
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        merged_df.to_excel(writer, sheet_name='SK에너지_통합분석', index=False)
        
        sk_cols = ['구분'] + [col for col in merged_df.columns if 'SK' in col]
        if len(sk_cols) > 1:
            sk_df = merged_df[sk_cols]
            sk_df.to_excel(writer, sheet_name='SK에너지_단독분석', index=False)
        
        display_df = merged_df.copy()
        display_df.to_excel(writer, sheet_name='경쟁사_비교분석', index=False)
    
    excel_buffer.seek(0)
    return excel_buffer.getvalue()

# ==========================
# SK 로고 표시 함수
# ==========================
def display_sk_header():
    """SK 브랜드 헤더 표시"""
    st.markdown(f"""
    <div class="sk-header">
        <h1 style="margin: 0; font-size: 32px;">
            ⚡ SK에너지 경쟁사 분석 대시보드
        </h1>
        <p style="margin: 10px 0 0 0; font-size: 18px; opacity: 0.9;">
            SK Energy Competitive Analysis Dashboard
        </p>
        <p style="margin: 5px 0 0 0; font-size: 14px; opacity: 0.8;">
            DART API 기반 실시간 재무분석 + RSS 뉴스 모니터링 (4가지 문제 해결 완료)
        </p>
    </div>
    """, unsafe_allow_html=True)

# ==========================
# 메인 함수
# ==========================
def main():
    # SK 브랜드 헤더 표시
    display_sk_header()
    
    # 사이드바
    with st.sidebar:
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, {SK_COLORS['primary']} 0%, {SK_COLORS['secondary']} 100%); 
                    color: white; padding: 15px; border-radius: 10px; margin-bottom: 20px; text-align: center;">
            <h3 style="margin: 0;">⚡ SK에너지 분석 설정</h3>
        </div>
        """, unsafe_allow_html=True)
        
        st.header("📊 분석 설정")
        analysis_year = st.selectbox("분석 연도", [2024, 2023, 2022], index=0)
        
        # 구글시트 연동 설정 (외부 자료 기반 손익 개선 아이디어)
        st.header("🔗 외부 데이터 연동")
        enable_google_sheets = st.checkbox("구글시트 자동화 연동", value=False, 
                                         help="Make에서 처리된 뉴스 분석 데이터를 실시간으로 로드합니다")
        
        google_sheets_data = None
        improvement_ideas = []
        
        if enable_google_sheets:
            spreadsheet_id = st.text_input("구글시트 ID", 
                                         placeholder="1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms",
                                         help="Make에서 데이터를 저장하는 구글시트의 ID를 입력하세요")
            
            if spreadsheet_id:
                sheets_connector = GoogleSheetsConnector()
                if sheets_connector.connect():
                    st.success("✅ 구글시트 연결 성공")
                    
                    # 뉴스 분석 데이터 로드
                    news_df = sheets_connector.load_news_analysis_data(spreadsheet_id)
                    if news_df is not None:
                        google_sheets_data = news_df
                        st.info(f"📰 {len(news_df)}개 뉴스 분석 데이터 로드")
                    
                    # 개선 아이디어 데이터 로드
                    ideas_df = sheets_connector.load_improvement_ideas(spreadsheet_id)
                    if ideas_df is not None:
                        improvement_ideas = ideas_df.to_dict('records')
                        st.info(f"💡 {len(improvement_ideas)}개 개선 아이디어 로드")
                else:
                    st.warning("⚠️ 구글시트 연결 실패")
        
        st.header("🏢 경쟁사 선택")
        st.info("💡 **SK에너지는 자동 포함**됩니다. 비교할 경쟁사를 선택하세요.")
        
        available_competitors = [
            'GS칼텍스', 
            '현대오일뱅크',
            'S-Oil'
        ]
        selected_competitors = st.multiselect(
            "경쟁사 선택 (SK에너지 vs 선택 기업들)",
            available_competitors,
            default=['GS칼텍스', 'S-Oil'],
            help="SK에너지와 비교할 경쟁사들을 선택하세요"
        )
        
        selected_companies = ['SK에너지'] + selected_competitors
        
        st.header("🎯 비교 기준 설정")
        comparison_metric = st.selectbox(
            "비교 기준 선택",
            ["매출 대비 비율", "절대값", "업계 평균 대비", "규모 정규화"],
            index=0,
            help="선택한 기준에 따라 표시되는 지표가 달라집니다"
        )
        st.session_state.comparison_metric = comparison_metric
    
    # 메인 탭들
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["🤖 재무분석", "📈 분기별 추이", "📰 뉴스 모니터링", "💡 손익 개선 아이디어", "📊 수동 업로드"])
    
    # === 탭 1: SK 중심 재무분석 ===
    with tab1:
        st.header("🤖 SK에너지 중심 재무분석")
        
        if not selected_competitors:
            st.info("👆 사이드바에서 비교할 경쟁사들을 선택해주세요.")
            return
        
        col1, col2 = st.columns([3, 1])
        
        with col1:
            analyze_button = st.button("🚀 SK에너지 경쟁사 분석 시작", type="primary")
        
        with col2:
            if st.session_state.analysis_results is not None:
                if st.button("🗑️ 결과 초기화", help="저장된 분석 결과를 삭제합니다"):
                    st.session_state.analysis_results = None
                    st.rerun()
        
        # 자동 분석 실행
        if analyze_button:
            dart_collector = DartAPICollector(DART_API_KEY)
            processor = SKFinancialDataProcessor()
            
            st.subheader("📡 DART API 데이터 수집")
            
            dataframes = []
            collected_companies = []
            
            # SK에너지를 먼저 처리
            with st.expander("⚡ SK에너지 데이터 수집 중", expanded=True):
                dart_df = dart_collector.get_company_financials_auto('SK에너지', str(analysis_year))
                
                if dart_df is not None and not dart_df.empty:
                    processed_df = processor.process_dart_data(dart_df, 'SK에너지')
                    
                    if processed_df is not None:
                        dataframes.append(processed_df)
                        collected_companies.append('SK에너지')
                        st.success("✅ SK에너지 데이터 처리 완료")
                        st.markdown(create_sk_korean_table(processed_df), unsafe_allow_html=True)
                    else:
                        st.error("❌ SK에너지 데이터 처리 실패")
                else:
                    st.error("❌ SK에너지 DART 데이터 수집 실패")
            
            # 경쟁사들 처리
            for company in selected_competitors:
                with st.expander(f"🏢 {company} 데이터 수집 중", expanded=True):
                    dart_df = dart_collector.get_company_financials_auto(company, str(analysis_year))
                    
                    if dart_df is not None and not dart_df.empty:
                        processed_df = processor.process_dart_data(dart_df, company)
                        
                        if processed_df is not None:
                            dataframes.append(processed_df)
                            collected_companies.append(company)
                            st.success(f"✅ {company} 데이터 처리 완료")
                            st.markdown(create_sk_korean_table(processed_df), unsafe_allow_html=True)
                        else:
                            st.error(f"❌ {company} 데이터 처리 실패")
                    else:
                        st.error(f"❌ {company} DART 데이터 수집 실패")
            
            if dataframes:
                merged_df = processor.merge_company_data(dataframes)
                
                st.session_state.analysis_results = {
                    'merged_df': merged_df,
                    'collected_companies': collected_companies,
                    'analysis_year': analysis_year,
                    'processor': processor,
                    'dart_collector': dart_collector
                }
                st.success("✅ SK에너지 경쟁사 분석 결과가 저장되었습니다!")
                st.rerun()
            else:
                st.error("❌ 수집된 데이터가 없습니다. 회사명을 확인해주세요.")
        
        # 저장된 분석 결과 표시
        if st.session_state.analysis_results is not None:
            results = st.session_state.analysis_results
            merged_df = results['merged_df']
            collected_companies = results['collected_companies']
            analysis_year = results['analysis_year']
            processor = results['processor']
            dart_collector = results.get('dart_collector')
            
            st.markdown(f"""
            <div class="analysis-card">
            <h3>📊 SK에너지 경쟁사 분석 결과</h3>
            <p><b>분석 완료:</b> SK에너지 vs {len([c for c in collected_companies if 'SK' not in c])}개 경쟁사</p>
            <p><b>분석 연도:</b> {analysis_year}년 | <b>생성일시:</b> {datetime.now().strftime('%Y-%m-%d %H:%M')}</p>
            </div>
            """, unsafe_allow_html=True)
            
            filtered_df = processor.apply_comparison_metric(merged_df, comparison_metric)
            
            st.subheader(f"📊 SK에너지 vs 경쟁사 비교 분석 - {comparison_metric} 기준")
            st.info(f"💡 현재 **{comparison_metric}** 기준으로 데이터를 표시하고 있습니다. 사이드바에서 기준을 변경할 수 있습니다.")
            
            st.write("**⚡ SK에너지 중심 경쟁사 비교 손익계산서**")
            st.markdown(create_sk_korean_table(filtered_df), unsafe_allow_html=True)
            
            # 🔗 DART 보고서 직접 링크 테이블 추가 (문제 2 해결)
            if dart_collector:
                dart_source_table = create_dart_source_table(dart_collector, collected_companies, analysis_year)
                st.markdown(dart_source_table, unsafe_allow_html=True)
            
            # 시각화 (개선된 버전 - PDF 차트 선택 기능 포함)
            chart_images = {}
            selected_charts = []
            
            if PLOTLY_AVAILABLE and len(collected_companies) > 1:
                st.subheader("📈 SK에너지 성과 시각화")
                
                ratio_data = filtered_df[filtered_df['구분'].str.contains('%|점|억원', na=False)]
                if not ratio_data.empty:
                    companies = [col for col in ratio_data.columns if col != '구분']
                    
                    chart_data = []
                    for _, row in ratio_data.iterrows():
                        for company in companies:
                            value_str = str(row[company]).replace('%', '').replace('점', '').replace('억원', '')
                            try:
                                chart_data.append({
                                    '지표': row['구분'],
                                    '회사': company,
                                    '수치': round(float(value_str), 2)
                                })
                            except:
                                pass
                    
                    if chart_data:
                        chart_df = pd.DataFrame(chart_data)
                        
                        # 모든 차트 생성 및 표시
                        st.markdown("**📊 생성된 차트들:**")
                        
                        # 차트 생성
                        all_charts = {}
                        
                        # 막대그래프
                        sk_bar_fig = create_sk_bar_chart(chart_df)
                        if sk_bar_fig:
                            st.markdown("**📊 SK 막대그래프**")
                            st.plotly_chart(sk_bar_fig, use_container_width=True)
                            all_charts["SK 막대그래프"] = sk_bar_fig
                        
                        # 레이더 차트
                        sk_radar_fig = create_sk_radar_chart(chart_df)
                        if sk_radar_fig:
                            st.markdown("**🎯 SK 레이더 차트**")
                            st.plotly_chart(sk_radar_fig, use_container_width=True)
                            all_charts["SK 레이더 차트"] = sk_radar_fig
                        
                        # 버블 차트
                        sk_bubble_fig = create_sk_bubble_chart(chart_df)
                        if sk_bubble_fig:
                            st.markdown("**🫧 SK 버블 차트**")
                            st.plotly_chart(sk_bubble_fig, use_container_width=True)
                            all_charts["SK 버블 차트"] = sk_bubble_fig
                        
                        # 히트맵
                        sk_heatmap_fig = create_sk_heatmap_chart(chart_df)
                        if sk_heatmap_fig:
                            st.markdown("**🔥 SK 히트맵**")
                            st.plotly_chart(sk_heatmap_fig, use_container_width=True)
                            all_charts["SK 히트맵"] = sk_heatmap_fig
                        
                        # PDF에 포함할 차트 선택 (개선된 버전)
                        if all_charts:
                            st.markdown("**📄 PDF에 포함할 차트 선택:**")
                            col1, col2, col3, col4 = st.columns(4)
                            
                            with col1:
                                include_bar = st.checkbox("막대그래프", value=True)
                            with col2:
                                include_radar = st.checkbox("레이더차트", value=True)
                            with col3:
                                include_bubble = st.checkbox("버블차트", value=False)
                            with col4:
                                include_heatmap = st.checkbox("히트맵", value=False)
                            
                            # 선택된 차트들을 저장할 딕셔너리
                            selected_charts = {}
                            
                            if include_bar and "SK 막대그래프" in all_charts:
                                chart_bytes = save_chart_as_image(all_charts["SK 막대그래프"], "sk_bar_chart.png")
                                if chart_bytes:
                                    selected_charts["SK 막대그래프"] = chart_bytes
                            
                            if include_radar and "SK 레이더 차트" in all_charts:
                                chart_bytes = save_chart_as_image(all_charts["SK 레이더 차트"], "sk_radar_chart.png")
                                if chart_bytes:
                                    selected_charts["SK 레이더 차트"] = chart_bytes
                            
                            if include_bubble and "SK 버블 차트" in all_charts:
                                chart_bytes = save_chart_as_image(all_charts["SK 버블 차트"], "sk_bubble_chart.png")
                                if chart_bytes:
                                    selected_charts["SK 버블 차트"] = chart_bytes
                            
                            if include_heatmap and "SK 히트맵" in all_charts:
                                chart_bytes = save_chart_as_image(all_charts["SK 히트맵"], "sk_heatmap_chart.png")
                                if chart_bytes:
                                    selected_charts["SK 히트맵"] = chart_bytes
                        
                        # 선택된 차트들을 chart_images에 저장
                        chart_images = selected_charts
                        
                        if selected_charts:
                            st.success(f"✅ {len(selected_charts)}개 차트가 PDF에 포함됩니다.")
                        else:
                            st.info("📄 PDF에 포함할 차트를 선택해주세요.")
            
            # SK 중심 인사이트 테이블 표시
            st.subheader("🧠 SK에너지 AI 인사이트")
            insight_table = create_sk_insight_table(filtered_df, collected_companies)
            st.markdown(insight_table, unsafe_allow_html=True)
            
            # 다운로드 섹션
            st.subheader("💾 SK에너지 분석 보고서 다운로드")
            
            st.info("""
            📋 **SK에너지 보고서 다운로드 (4가지 문제 해결 완료):**
            - **Excel**: SK에너지 중심 분석 데이터 (한글 지원) ⭐ 추천
            - **PDF**: SK 브랜드 보고서 (한글 폰트 + 페이지 번호 + DART 출처 링크)
            - **CSV**: 원시 데이터 분석용
            """)
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                excel_data = create_excel_file(merged_df, collected_companies, analysis_year)
                st.download_button(
                    label="📊 SK 분석 Excel ⭐",
                    data=excel_data,
                    file_name=f"SK에너지_경쟁사분석_{analysis_year}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            with col2:
                if PDF_AVAILABLE:
                    # 선택된 차트들을 PDF에 포함
                    pdf_data = create_enhanced_pdf_report(merged_df, collected_companies, analysis_year, chart_images)
                    if pdf_data:
                        chart_count = len(chart_images) if chart_images else 0
                        pdf_label = f"📄 SK 보고서 PDF ({chart_count}개 차트)"
                        st.download_button(
                            label=pdf_label,
                            data=pdf_data,
                            file_name=f"SK에너지_경쟁사분석보고서_{analysis_year}.pdf",
                            mime="application/pdf"
                        )
                else:
                    st.warning("PDF 기능은 reportlab 설치 필요")
            
            with col3:
                csv = merged_df.to_csv(index=False, encoding='utf-8-sig')
                st.download_button(
                    label="📋 원시데이터 CSV",
                    data=csv,
                    file_name=f"SK에너지_raw_data_{analysis_year}.csv",
                    mime="text/csv",
                    help="고급 분석용 원시 데이터"
                )
    
    # === 탭 2: 분기별 추이 분석 ===
    with tab2:
        st.header("📈 분기별 성과 추이 분석")
        st.info("💡 **분기별 트렌드 분석**으로 SK에너지와 경쟁사의 성과 변화를 추적합니다.")
        
        if not selected_competitors:
            st.info("👆 사이드바에서 비교할 경쟁사들을 선택해주세요.")
        else:
            col1, col2 = st.columns([2, 1])
            
            with col1:
                quarterly_year = st.selectbox("분기별 분석 연도", [2024, 2023, 2022], index=0)
            
            with col2:
                collect_quarterly = st.button("📊 분기별 데이터 수집", type="primary")
            
            if collect_quarterly:
                dart_collector = DartAPICollector(DART_API_KEY)
                quarterly_collector = QuarterlyDataCollector(dart_collector)
                
                st.subheader("📡 분기별 DART 데이터 수집")
                
                all_quarterly_data = []
                
                # SK에너지 분기별 데이터 수집
                with st.expander("⚡ SK에너지 분기별 데이터 수집", expanded=True):
                    sk_quarterly = quarterly_collector.collect_quarterly_data('SK에너지', quarterly_year)
                    if not sk_quarterly.empty:
                        all_quarterly_data.append(sk_quarterly)
                        st.success("✅ SK에너지 분기별 데이터 수집 완료")
                        st.dataframe(sk_quarterly)
                
                # 경쟁사 분기별 데이터 수집
                for company in selected_competitors:
                    with st.expander(f"🏢 {company} 분기별 데이터 수집", expanded=True):
                        comp_quarterly = quarterly_collector.collect_quarterly_data(company, quarterly_year)
                        if not comp_quarterly.empty:
                            all_quarterly_data.append(comp_quarterly)
                            st.success(f"✅ {company} 분기별 데이터 수집 완료")
                            st.dataframe(comp_quarterly)
                
                if all_quarterly_data:
                    quarterly_df = pd.concat(all_quarterly_data, ignore_index=True)
                    st.session_state.quarterly_data = quarterly_df
                    st.success("✅ 모든 분기별 데이터 수집 및 저장 완료!")
                    st.rerun()
                else:
                    st.error("❌ 분기별 데이터 수집 실패")
            
            # 저장된 분기별 데이터 표시
            if st.session_state.quarterly_data is not None:
                quarterly_df = st.session_state.quarterly_data
                
                st.subheader("📊 분기별 성과 추이 분석")
                
                if PLOTLY_AVAILABLE:
                    trend_fig = create_quarterly_trend_chart(quarterly_df)
                    if trend_fig:
                        st.plotly_chart(trend_fig, use_container_width=True)
                
                st.subheader("📋 분기별 상세 데이터")
                
                sk_data = quarterly_df[quarterly_df['회사'] == 'SK에너지']
                other_data = quarterly_df[quarterly_df['회사'] != 'SK에너지']
                sorted_quarterly = pd.concat([sk_data, other_data], ignore_index=True)
                
                st.dataframe(sorted_quarterly, use_container_width=True)
    
    # === 탭 3: SK 중심 뉴스 모니터링 ===
    with tab3:
        st.header("📰 SK에너지 중심 뉴스 모니터링")
        st.info("💡 **SK에너지 관련 뉴스**를 우선적으로 수집하고 분석합니다.")
        
        rss_collector = SKNewsRSSCollector()
        
        col1, col2 = st.columns([1, 1])
        
        with col1:
            business_type = st.selectbox("📈 사업 분야", ['정유'], disabled=True, help="정유업 중심 분석")
        
        with col2:
            auto_collect = st.button("🔄 SK 중심 뉴스 수집", type="primary")
        
        if auto_collect:
            news_df = rss_collector.collect_real_korean_news(business_type)
            
            if not news_df.empty:
                # SK 테마 워드클라우드
                st.subheader("☁️ SK에너지 키워드 클라우드")
                
                if WORDCLOUD_AVAILABLE:
                    wordcloud_img = rss_collector.create_sk_wordcloud(news_df)
                    if wordcloud_img:
                        st.image(wordcloud_img, caption="SK에너지 관련 뉴스 키워드 분석 (한글 지원)", width=800)
                    else:
                        st.warning("워드클라우드 생성에 실패했습니다.")
                else:
                    st.warning("워드클라우드 기능을 사용하려면 'pip install wordcloud matplotlib' 설치가 필요합니다.")
                
                # SK 중심 뉴스 분석 차트
                if PLOTLY_AVAILABLE:
                    col1, col2 = st.columns([1, 1])
                    
                    with col1:
                        fig = rss_collector.create_keyword_analysis(news_df)
                        if fig:
                            st.plotly_chart(fig, use_container_width=True)
                    
                    with col2:
                        company_counts = news_df['회사'].value_counts()
                        
                        # 파스텔 색상 적용
                        color_map = {}
                        companies = list(company_counts.index)
                        for company in companies:
                            color_map[company] = get_company_color(company, companies)
                        
                        fig2 = px.bar(
                            x=company_counts.values,
                            y=company_counts.index,
                            orientation='h',
                            title="🏢 회사별 뉴스 언급 빈도 (SK 강조)",
                            height=400,
                            color=company_counts.index,
                            color_discrete_map=color_map
                        )
                        # 폰트 크기 증가
                        fig2.update_layout(
                            title_font_size=18,
                            font_size=14,
                            xaxis=dict(title_font_size=16, tickfont=dict(size=14)),
                            yaxis=dict(title_font_size=16, tickfont=dict(size=14))
                        )
                        st.plotly_chart(fig2, use_container_width=True)
                
                # 뉴스 목록 표시
                st.subheader("📋 SK에너지 우선 뉴스 목록")
                
                col1, col2 = st.columns(2)
                with col1:
                    impact_filter = st.slider("🎯 최소 영향도", 1, 10, 3)
                with col2:
                    sk_priority = st.checkbox("⚡ SK 관련 뉴스 우선 표시", value=True)
                
                filtered_news = news_df[news_df['영향도'] >= impact_filter]
                if sk_priority:
                    sk_news = filtered_news[filtered_news['회사'].str.contains('SK', na=False)]
                    other_news = filtered_news[~filtered_news['회사'].str.contains('SK', na=False)]
                    filtered_news = pd.concat([sk_news, other_news], ignore_index=True)
                
                if not filtered_news.empty:
                    st.info(f"📄 총 {len(filtered_news)}개 뉴스 발견")
                    for idx, row in filtered_news.head(10).iterrows():
                        is_sk = 'SK' in str(row['회사'])
                        card_color = SK_COLORS['primary'] if is_sk else SK_COLORS['competitor']
                        
                        st.markdown(f"""
                        <div style="border: 2px solid {card_color}; padding: 10px; margin: 10px 0; border-radius: 8px;">
                            <h4 style="color: {card_color}; margin: 0 0 5px 0;">{row['제목']}</h4>
                            <p style="margin: 5px 0; font-size: 12px;">🏢 {row['회사']} | 📅 {row['날짜']} | 🎯 {row['영향도']}/10</p>
                            <a href="{row['URL']}" target="_blank">🔗 뉴스 원문</a>
                        </div>
                        """, unsafe_allow_html=True)
                else:
                    st.warning("조건에 맞는 뉴스가 없습니다.")
            else:
                st.warning("관련 뉴스를 찾을 수 없습니다.")
    
    # === 탭 4: 외부 자료 기반 손익 개선 아이디어 ===
    with tab4:
        st.header("💡 외부 자료 기반 손익 개선 아이디어")
        
        if enable_google_sheets and google_sheets_data is not None:
            st.success("✅ Make 자동화와 연동되어 외부 자료를 실시간으로 분석합니다.")
            
            # 뉴스 분석 결과 표시
            st.subheader("📰 뉴스 분석 결과")
            if not google_sheets_data.empty:
                st.dataframe(google_sheets_data, use_container_width=True)
                
                # 개선 아이디어 도출
                if improvement_ideas:
                    st.subheader("🎯 손익 개선 아이디어")
                    
                    for i, idea in enumerate(improvement_ideas, 1):
                        with st.expander(f"💡 아이디어 {i}: {idea.get('title', '제목 없음')}", expanded=True):
                            col1, col2, col3 = st.columns(3)
                            
                            with col1:
                                st.metric("영향도", idea.get('impact', '중간'))
                            with col2:
                                st.metric("구현 기간", idea.get('implementation', '중장기'))
                            with col3:
                                st.metric("키워드", idea.get('keyword', 'N/A'))
                            
                            st.markdown(f"**분석 내용:** {idea.get('content', '내용 없음')}")
                            
                            # SK 브랜드 컬러로 강조
                            if idea.get('impact') == '높음':
                                st.markdown("""
                                <div style='background: rgba(227, 30, 36, 0.1); padding: 10px; border-radius: 5px; border-left: 4px solid #E31E24;'>
                                    <strong>🔥 고우선순위 아이디어</strong><br/>
                                    이 아이디어는 높은 영향도를 가져 즉시 검토가 필요합니다.
                                </div>
                                """, unsafe_allow_html=True)
        else:
            st.info("""
            📋 **외부 자료 기반 손익 개선 아이디어 분석**
            
            이 기능을 사용하려면:
            1. 사이드바에서 "구글시트 자동화 연동" 활성화
            2. Make에서 처리된 구글시트 ID 입력
            3. 실시간으로 외부 자료 분석 결과 확인
            
            **Make 자동화 파이프라인:**
            - 뉴스 스크랩 → OpenAI 분석 → 구글시트 저장
            - 실시간 손익 개선 아이디어 도출
            """)
    
    # === 탭 5: 수동 업로드 ===
    with tab5:
        st.header("📊 수동 XBRL/XML 파일 업로드 분석")
        st.info("💡 **XBRL/XML 파싱 기능**이 개선되었습니다. 안전하게 업로드하세요.")
        
        uploaded_files = st.file_uploader(
            "XBRL/XML 파일들을 업로드하세요",
            type=['xbrl', 'xml'],
            accept_multiple_files=True,
            help="SK에너지 및 경쟁사의 XBRL/XML 파일을 업로드하면 자동으로 분석합니다."
        )
        
        if uploaded_files:
            st.success(f"✅ {len(uploaded_files)}개 파일이 업로드되었습니다.")
            
            for file in uploaded_files:
                with st.expander(f"📄 {file.name}", expanded=False):
                    st.write(f"**파일명**: {file.name}")
                    st.write(f"**파일 크기**: {file.size:,} bytes")
                    st.write(f"**파일 타입**: {file.type}")
            
            process_files = st.button("🚀 XBRL/XML 파일 분석 시작", type="primary")
            
            if process_files:
                xbrl_parser = EnhancedXBRLParser()
                parsed_data = []
                
                st.subheader("📡 XBRL/XML 파일 파싱 중...")
                
                for file in uploaded_files:
                    with st.expander(f"⚙️ {file.name} 파싱 중", expanded=True):
                        try:
                            file_content = file.read()
                            
                            # XBRL 파싱
                            parsed_result = xbrl_parser.parse_xbrl_file(file_content, file.name)
                            
                            if parsed_result:
                                parsed_data.append(parsed_result)
                                st.success(f"✅ {file.name} 파싱 완료")
                                st.json(parsed_result)
                            else:
                                st.warning(f"⚠️ {file.name} 파싱 실패 (재무 데이터 없음)")
                                
                        except Exception as e:
                            st.error(f"❌ {file.name} 처리 오류: {e}")
                
                if parsed_data:
                    st.success(f"🎉 총 {len(parsed_data)}개 파일 파싱 완료!")
                    
                    # 파싱된 데이터를 DataFrame으로 변환
                    manual_df_list = []
                    for data in parsed_data:
                        if data['data']:
                            df_data = []
                            for key, value in data['data'].items():
                                df_data.append({
                                    '구분': key,
                                    data['company_name']: f"{value:,.0f}원" if value >= 10000 else f"{value:,.2f}"
                                })
                            
                            if df_data:
                                manual_df_list.append(pd.DataFrame(df_data))
                    
                    if manual_df_list:
                        processor = SKFinancialDataProcessor()
                        manual_merged = processor.merge_company_data(manual_df_list)
                        
                        st.subheader("📊 수동 업로드 분석 결과")
                        st.markdown(create_sk_korean_table(manual_merged), unsafe_allow_html=True)
                        
                        manual_csv = manual_merged.to_csv(index=False, encoding='utf-8-sig')
                        st.download_button(
                            label="📋 수동 분석 CSV 다운로드",
                            data=manual_csv,
                            file_name=f"수동업로드_분석결과.csv",
                            mime="text/csv"
                        )
                    else:
                        st.warning("파싱된 재무 데이터가 없습니다.")
                else:
                    st.warning("파싱 가능한 파일이 없습니다. 다른 파일을 시도해보세요.")
        else:
            st.info("👆 분석할 XBRL/XML 파일을 업로드해주세요.")
            
            st.markdown(f"""
            <div class="analysis-card">
                <h4>📋 XBRL/XML 파일 업로드 가이드</h4>
                <ul style="margin: 10px 0; padding-left: 20px;">
                    <li><b>지원 형식:</b> .xbrl, .xml 파일 (한글 인코딩 지원)</li>
                    <li><b>권장 용도:</b> DART에서 직접 다운받은 재무제표 파일</li>
                    <li><b>자동 분석:</b> 업로드 즉시 SK에너지 중심 비교 분석</li>
                    <li><b>결과 제공:</b> 테이블, CSV 다운로드</li>
                </ul>
                <p style="margin: 10px 0; color: {SK_COLORS['primary']}; font-weight: 600;">
                    💡 <b>추천:</b> DART API 자동 분석이 더 정확하고 빠릅니다!
                </p>
            </div>
            """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()

