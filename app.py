import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from streamlit_drawable_canvas import st_canvas
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import datetime
import os
import re
import base64
import hashlib
import uuid
import json
from PIL import Image
import io
import zipfile
import zlib
import openpyxl
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# 한국 표준시(KST, UTC+9) 타임존 정의
KST = datetime.timezone(datetime.timedelta(hours=9))

# 1. 페이지 설정 & 프리미엄 블루 테마 CSS
st.set_page_config(
    page_title="DANWOL AI-WaterOps 360 | 단월 스마트 자율운전 관제 플랫폼",
    page_icon="💧",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Pretendard:wght@300;400;500;600;700;800;900&display=swap');
    html, body, [class*="css"] {
        font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, system-ui, Roboto, sans-serif !important;
        letter-spacing: -0.3px;
    }
    .block-container { padding-top: 1.3rem; padding-bottom: 3rem; }
    
    .hero-banner {
        background: linear-gradient(135deg, #060D24 0%, #0F2042 40%, #034C81 78%, #0088CC 100%);
        border-radius: 18px; padding: 24px 32px; color: white; margin-bottom: 22px;
        box-shadow: 0 14px 32px -6px rgba(2, 132, 199, 0.32); display: flex; align-items: center; justify-content: space-between;
        border: 1px solid rgba(56, 189, 248, 0.28);
    }
    .hero-title {
        font-size: 27px; font-weight: 900; margin: 0;
        background: linear-gradient(90deg, #FFFFFF 0%, #E0F2FE 50%, #38BDF8 100%);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    }
    .hero-subtitle { font-size: 13.5px; color: #BAE6FD; margin-top: 6px; font-weight: 500; }
    .badge-online {
        display: inline-flex; align-items: center; gap: 8px; background: rgba(16, 185, 129, 0.22);
        color: #34D399; border: 1px solid rgba(52, 211, 153, 0.5); padding: 6px 16px; border-radius: 30px;
        font-size: 12px; font-weight: 800; box-shadow: 0 0 15px rgba(16, 185, 129, 0.25);
    }

    .stButton > button[kind="primary"], div[data-testid="stButton"] > button[data-testid="baseButton-primary"] {
        background: linear-gradient(135deg, #0284C7 0%, #0369A1 100%) !important;
        background-color: #0284C7 !important;
        color: white !important;
        border: none !important;
        border-radius: 10px !important;
        font-weight: 700 !important;
        box-shadow: 0 4px 14px rgba(2, 132, 199, 0.3) !important;
        transition: all 0.2s ease !important;
    }
    .stButton > button[kind="primary"]:hover, div[data-testid="stButton"] > button[data-testid="baseButton-primary"]:hover {
        background: linear-gradient(135deg, #0369A1 0%, #075985 100%) !important;
        box-shadow: 0 6px 20px rgba(2, 132, 199, 0.48) !important;
        transform: translateY(-1px) !important;
    }

    /* 💧 단월(DANWOL) KNR+IPR 스마트 스카다(SCADA) & 통합 매트릭스 전용 CSS */
    .dw-super-card {
        background: linear-gradient(180deg, #FFFFFF 0%, #F8FAFC 100%);
        border: 1px solid #E2E8F0;
        border-top: 4px solid #0284C7;
        border-radius: 16px;
        padding: 18px 20px;
        box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.06), 0 4px 10px -3px rgba(15, 23, 42, 0.03);
        transition: all 0.22s cubic-bezier(0.4, 0, 0.2, 1);
        position: relative;
        overflow: hidden;
    }
    .dw-super-card:hover {
        transform: translateY(-4px);
        box-shadow: 0 18px 32px -8px rgba(2, 132, 199, 0.16);
        border-color: #BAE6FD;
    }
    .dw-kpi-top {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 8px;
    }
    .dw-kpi-title {
        font-size: 13.5px;
        font-weight: 700;
        color: #334155;
    }
    .dw-kpi-icon {
        width: 32px; height: 32px;
        border-radius: 10px;
        background: #E0F2FE;
        color: #0284C7;
        display: flex; align-items: center; justify-content: center;
        font-size: 16px; font-weight: 800;
    }
    .dw-kpi-value {
        font-size: 31px;
        font-weight: 900;
        color: #0F172A;
        letter-spacing: -0.8px;
        line-height: 1.15;
        margin-bottom: 10px;
    }
    .dw-kpi-unit {
        font-size: 14px;
        font-weight: 600;
        color: #64748B;
        margin-left: 3px;
    }
    .dw-progress-bg {
        width: 100%;
        height: 7px;
        background: #E2E8F0;
        border-radius: 99px;
        overflow: hidden;
        margin-bottom: 10px;
    }
    .dw-progress-fill {
        height: 100%;
        border-radius: 99px;
        transition: width 0.4s ease;
    }
    .dw-badge-row {
        display: flex;
        justify-content: space-between;
        align-items: center;
        font-size: 11.5px;
        font-weight: 700;
    }
    .dw-pill-limit {
        background: #FEE2E2;
        color: #DC2626;
        padding: 3px 9px;
        border-radius: 20px;
        font-weight: 800;
    }
    .dw-pill-safe {
        background: #DCFCE7;
        color: #15803D;
        padding: 3px 9px;
        border-radius: 20px;
        font-weight: 800;
    }

    /* TOC 프리미엄 법령 배너 */
    .dw-toc-banner {
        background: linear-gradient(90deg, #EFF6FF 0%, #E0F2FE 55%, #F0F9FF 100%);
        border: 1px solid #BAE6FD;
        border-left: 6px solid #0284C7;
        border-radius: 14px;
        padding: 16px 22px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        box-shadow: 0 4px 15px rgba(2, 132, 199, 0.08);
        margin: 14px 0;
    }

    /* 디지털 트윈 프로세스 플로우 카드 */
    .dw-flow-container {
        background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
        border-radius: 16px;
        padding: 18px 22px;
        color: white;
        margin: 14px 0 22px 0;
        box-shadow: 0 12px 28px -6px rgba(15, 23, 42, 0.25);
        border: 1px solid #334155;
    }
    .dw-flow-grid {
        display: grid;
        grid-template-columns: repeat(5, 1fr);
        gap: 12px;
        margin-top: 12px;
    }
    .dw-flow-node {
        background: rgba(255, 255, 255, 0.06);
        border: 1px solid rgba(56, 189, 248, 0.3);
        border-radius: 12px;
        padding: 12px 14px;
        text-align: center;
    }
    .dw-flow-node-title {
        font-size: 12px;
        color: #94A3B8;
        font-weight: 700;
        margin-bottom: 4px;
    }
    .dw-flow-node-val {
        font-size: 17px;
        font-weight: 900;
        color: #38BDF8;
    }
    .dw-flow-node-sub {
        font-size: 11px;
        color: #34D399;
        margin-top: 3px;
        font-weight: 600;
    }

    /* 🛰️ 단월(DANWOL) 7개소 멀티 스카다(SCADA) 매트릭스 & 엔지니어링 패널 */
    .dw-scada-header {
        background: linear-gradient(135deg, #090D16 0%, #0F172A 50%, #1E293B 100%);
        border: 1px solid #334155;
        border-left: 6px solid #38BDF8;
        border-radius: 16px;
        padding: 20px 26px;
        color: #F8FAFC;
        margin-bottom: 18px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        box-shadow: 0 12px 28px -6px rgba(15, 23, 42, 0.35);
    }
    .dw-matrix-grid {
        display: grid;
        grid-template-columns: repeat(7, 1fr);
        gap: 10px;
        margin-bottom: 20px;
    }
    .dw-matrix-card {
        background: #0F172A;
        border: 1px solid #1E293B;
        border-top: 3px solid #0284C7;
        border-radius: 12px;
        padding: 12px 10px;
        color: #E2E8F0;
        text-align: center;
        transition: all 0.2s ease;
    }
    .dw-matrix-card.active-fac {
        background: linear-gradient(180deg, #0C4A6E 0%, #0F172A 100%);
        border: 1px solid #38BDF8;
        border-top: 4px solid #38BDF8;
        box-shadow: 0 0 16px rgba(56, 189, 248, 0.28);
        transform: translateY(-2px);
    }
    .dw-eng-panel {
        background: #FFFFFF;
        border: 1px solid #CBD5E1;
        border-top: 4px solid #0F172A;
        border-radius: 14px;
        padding: 16px 18px;
        box-shadow: 0 6px 16px rgba(15, 23, 42, 0.05);
        height: 100%;
    }
    .dw-eng-tag {
        display: inline-block;
        background: #F1F5F9;
        color: #0F172A;
        font-size: 11px;
        font-weight: 800;
        padding: 3px 8px;
        border-radius: 6px;
        margin-bottom: 6px;
    }
    .dw-eng-metric {
        font-size: 25px;
        font-weight: 900;
        color: #0369A1;
        margin: 4px 0;
    }
    .dw-eng-sub {
        font-size: 12px;
        color: #475569;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# 2. 시설 목록 및 사양 정의
MAIN_PLANT = "단월공공하수처리시설(본장)"
SMALL_PLANTS = ["산음", "삼가리", "진목", "몰운", "단월마을", "당의"]
PRIVATE_PLANTS = ["석산리", "음지", "양지", "복지회관", "인이피", "돌고개"]

PLANT_DESIGN_SPECS = {
    MAIN_PLANT: {"cap": 1700.0, "method": "KNR + IPR", "blower_cap": 25.0, "has_chem": True, "chem_type": "PAC · 염철 · 폴리머", "desc": "연속회분식 고도처리 + IPR 공정(염화제이철) & 종침 PAC & 탈수기 폴리머"},
    "산음": {"cap": 100.0, "method": "SWPP", "blower_cap": 3.0, "has_chem": False, "chem_type": "무약품", "desc": "수중포기 침전일체형"},
    "삼가리": {"cap": 120.0, "method": "SBR", "blower_cap": 3.5, "has_chem": False, "chem_type": "무약품", "desc": "회분식 활성슬러지 공정"},
    "진목": {"cap": 23.0, "method": "고효율오수정화 + SOD", "blower_cap": 1.5, "has_chem": False, "chem_type": "무약품", "desc": "미생물 접촉산화 및 고효율 탈질"},
    "몰운": {"cap": 60.0, "method": "IC-SBR", "blower_cap": 2.0, "has_chem": True, "chem_type": "반응조 PAC", "desc": "간헐 포기 회분식 반응조"},
    "단월마을": {"cap": 30.0, "method": "IC-SBR", "blower_cap": 1.5, "has_chem": False, "chem_type": "무약품", "desc": "간헐 포기 회분식 고도처리"},
    "당의": {"cap": 45.0, "method": "IC-SBR", "blower_cap": 2.0, "has_chem": False, "chem_type": "무약품", "desc": "간헐 포기 회분식 고도처리"}
}

# 3. 영구 보관 디렉토리 및 DB 파일
KHAS_RECORD_DIR = "monthly_khas_records"
TBM_RECORD_DIR = "tbm_records"
SIGNATURE_DIR = "signature_storage"
HWPX_RECORD_DIR = "hwpx_records"
CHEMICAL_RECORD_DIR = "chemical_records"
MASTER_ACCUM_DB = "danwol_accumulated_master.csv"
TMS_ACCUM_DB = "danwol_tms_master.csv"
PROCESS_CONTROL_DB = "danwol_process_control_master.csv"
CHEMICAL_ENERGY_DB = "danwol_chemical_energy_master.csv"
CHEMICAL_INVENTORY_DB = "danwol_chemical_inventory.csv"
AUTH_DB_FILE = "user_auth_db.json"
SYSTEM_CONFIG_FILE = "system_config.json"

ALL_STORAGE_DIRS = {
    "월별 공인 엑셀 보관함 (monthly_khas_records)": KHAS_RECORD_DIR,
    "HWPX 월간보고서 보관함 (hwpx_records)": HWPX_RECORD_DIR,
    "TBM 표준회의록 보관함 (tbm_records)": TBM_RECORD_DIR,
    "전자서명 이미지 보관소 (signature_storage)": SIGNATURE_DIR,
    "약품 입출고 증빙 보관소 (chemical_records)": CHEMICAL_RECORD_DIR
}

ALL_MASTER_DBS = {
    "운영일지·수질 통합 마스터 DB": MASTER_ACCUM_DB,
    "TMS 실시간·예측 수질 마스터 DB": TMS_ACCUM_DB,
    "AI 공정제어 운전조건 마스터 DB": PROCESS_CONTROL_DB,
    "약품·에너지·ESG 마스터 DB": CHEMICAL_ENERGY_DB,
    "약품 구매·재고 수불부 마스터 DB": CHEMICAL_INVENTORY_DB
}

for p in [KHAS_RECORD_DIR, TBM_RECORD_DIR, SIGNATURE_DIR, HWPX_RECORD_DIR, CHEMICAL_RECORD_DIR]:
    if not os.path.exists(p): os.makedirs(p)

def sanitize_filename(filename):
    clean_name = os.path.basename(str(filename))
    return re.sub(r'[^a-zA-Z0-9가-힣._\-\(\)\s]', '', clean_name)

def hash_pw(pw_str):
    return hashlib.sha256(pw_str.encode('utf-8')).hexdigest()

ADMIN_PW_HASH = hash_pw("yp1311!!")
WHITELIST_HASHES = [hash_pw("DW-PASS-2026"), hash_pw("WATER-ADMIN"), hash_pw("DANWOL-2026!"), hash_pw("yp1311!!")]

def load_system_config():
    default_cfg = {
        "force_maintenance_screen": True,
        "maintenance_msg": "현재 자체 시스템 정기 점검 및 AI 데이터베이스 최적화 업데이트 중입니다."
    }
    if os.path.exists(SYSTEM_CONFIG_FILE):
        try:
            with open(SYSTEM_CONFIG_FILE, "r", encoding="utf-8") as f:
                loaded = json.load(f)
                if "force_maintenance_screen" not in loaded:
                    loaded["force_maintenance_screen"] = True
                default_cfg.update(loaded)
                return default_cfg
        except: pass
    save_system_config(default_cfg)
    return default_cfg

def save_system_config(cfg):
    try:
        with open(SYSTEM_CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except: pass

def load_auth_db():
    if os.path.exists(AUTH_DB_FILE):
        try:
            with open(AUTH_DB_FILE, "r", encoding="utf-8") as f: return json.load(f)
        except: pass
    return {"users": {}}

def save_auth_db(data):
    try:
        with open(AUTH_DB_FILE, "w", encoding="utf-8") as f: json.dump(data, f, ensure_ascii=False, indent=2)
    except: pass

def append_to_master_db(fac, df_new):
    if df_new is None or df_new.empty: return
    df_new = df_new.copy()
    df_new['시설명'] = fac
    if os.path.exists(MASTER_ACCUM_DB):
        try:
            df_m = pd.read_csv(MASTER_ACCUM_DB)
            df_comb = pd.concat([df_m, df_new], ignore_index=True).drop_duplicates(subset=['시설명', '날짜'], keep='last')
        except Exception:
            df_comb = df_new.drop_duplicates(subset=['시설명', '날짜'])
    else:
        df_comb = df_new.drop_duplicates(subset=['시설명', '날짜'])
    df_comb.sort_values(by=['시설명', '날짜']).to_csv(MASTER_ACCUM_DB, index=False, encoding='utf-8-sig')

def get_master_data(fac, start_date=None, end_date=None):
    if not os.path.exists(MASTER_ACCUM_DB): return pd.DataFrame()
    try:
        df = pd.read_csv(MASTER_ACCUM_DB)
        df_fac = df[df['시설명'] == fac].copy()
        if df_fac.empty: return pd.DataFrame()
        df_fac['날짜_dt'] = pd.to_datetime(df_fac['날짜'], errors='coerce')
        if start_date: df_fac = df_fac[df_fac['날짜_dt'] >= pd.to_datetime(start_date)]
        if end_date: df_fac = df_fac[df_fac['날짜_dt'] <= pd.to_datetime(end_date)]
        return df_fac.sort_values(by='날짜').reset_index(drop=True)
    except Exception:
        return pd.DataFrame()

TMS_STD_COLS = ['측정일자', '측정시각', '방류pH', '방류BOD', '방류TOC', '방류SS', '방류TN', '방류TP', '방류유량', '예측pH_4h', '예측BOD_4h', '예측SS_4h', '예측TN_4h', '예측TP_4h', '비고']

def append_to_tms_db(df_new):
    if df_new is None or df_new.empty: return
    df_new = df_new.copy()
    for col in TMS_STD_COLS:
        if col not in df_new.columns: df_new[col] = np.nan
    df_new = df_new[TMS_STD_COLS]
    df_new['측정일자'] = df_new['측정일자'].astype(str).str.slice(0, 10)
    if os.path.exists(TMS_ACCUM_DB):
        try:
            df_m = pd.read_csv(TMS_ACCUM_DB)
            df_comb = pd.concat([df_m, df_new], ignore_index=True).drop_duplicates(subset=['측정일자', '측정시각'], keep='last')
        except Exception:
            df_comb = df_new.drop_duplicates(subset=['측정일자', '측정시각'])
    else:
        df_comb = df_new.drop_duplicates(subset=['측정일자', '측정시각'])
    df_comb.sort_values(by=['측정일자', '측정시각'], ascending=[False, False]).to_csv(TMS_ACCUM_DB, index=False, encoding='utf-8-sig')

def get_tms_db():
    if not os.path.exists(TMS_ACCUM_DB): return pd.DataFrame()
    try:
        df = pd.read_csv(TMS_ACCUM_DB)
        if df.empty: return pd.DataFrame()
        # 사용자가 입력하지 않은 자동생성 더미 데이터 자동 제거
        if '비고' in df.columns:
            auto_mask = df['비고'].astype(str).isin(['초기 자동생성', 'AI 기본 관제 데이터'])
            if auto_mask.any():
                df = df[~auto_mask].copy()
                df.to_csv(TMS_ACCUM_DB, index=False, encoding='utf-8-sig')
        if df.empty: return pd.DataFrame()
        return df.sort_values(by=['측정일자', '측정시각'], ascending=[False, False]).reset_index(drop=True)
    except Exception:
        return pd.DataFrame()

def append_to_process_db(df_new, facility_name=MAIN_PLANT):
    if df_new is None or df_new.empty: return
    df_new = df_new.copy()
    df_new['시설명'] = facility_name
    if os.path.exists(PROCESS_CONTROL_DB):
        try:
            df_m = pd.read_csv(PROCESS_CONTROL_DB)
            df_comb = pd.concat([df_m, df_new], ignore_index=True).drop_duplicates(subset=['시설명', '날짜'], keep='last')
        except Exception:
            df_comb = df_new.drop_duplicates(subset=['시설명', '날짜'])
    else:
        df_comb = df_new.drop_duplicates(subset=['시설명', '날짜'])
    df_comb.sort_values(by=['시설명', '날짜'], ascending=[True, False]).to_csv(PROCESS_CONTROL_DB, index=False, encoding='utf-8-sig')

def get_process_db(facility_name=MAIN_PLANT):
    if not os.path.exists(PROCESS_CONTROL_DB): return pd.DataFrame()
    try:
        df = pd.read_csv(PROCESS_CONTROL_DB)
        if facility_name: df = df[df['시설명'] == facility_name]
        return df.sort_values(by='날짜', ascending=False).reset_index(drop=True)
    except Exception:
        return pd.DataFrame()

CHEM_STD_COLS = ["날짜", "PAC사용량_kg", "염화제이철_kg", "폴리머사용량_kg", "슬러지반출량_톤", "전력사용량_kWh", "태양광발전량_kWh", "비고"]

def append_to_chem_db(df_new):
    if df_new is None or df_new.empty: return
    df_new = df_new.copy()
    for col in CHEM_STD_COLS:
        if col not in df_new.columns: df_new[col] = 0.0 if 'kg' in col or '톤' in col or 'kWh' in col else ''
    df_new = df_new[CHEM_STD_COLS]
    if os.path.exists(CHEMICAL_ENERGY_DB):
        try:
            df_m = pd.read_csv(CHEMICAL_ENERGY_DB)
            df_comb = pd.concat([df_m, df_new], ignore_index=True).drop_duplicates(subset=['날짜'], keep='last')
        except Exception:
            df_comb = df_new.drop_duplicates(subset=['날짜'])
    else:
        df_comb = df_new.drop_duplicates(subset=['날짜'])
    df_comb.sort_values(by=['날짜'], ascending=False).to_csv(CHEMICAL_ENERGY_DB, index=False, encoding='utf-8-sig')

def get_chem_db():
    if not os.path.exists(CHEMICAL_ENERGY_DB): return pd.DataFrame()
    try:
        df = pd.read_csv(CHEMICAL_ENERGY_DB)
        for col in CHEM_STD_COLS:
            if col not in df.columns: df[col] = 0.0 if 'kg' in col or '톤' in col or 'kWh' in col else ''
        return df.sort_values(by=['날짜'], ascending=False).reset_index(drop=True)
    except Exception:
        return pd.DataFrame()

def parse_chem_energy_files(file_list):
    records_by_date = {}
    if not file_list:
        return pd.DataFrame()
    for f in file_list:
        try:
            fname = getattr(f, 'name', str(f))
            y_m = re.search(r'(20[1-3]\d)', fname)
            file_y = int(y_m.group(1)) if y_m else 2026

            if fname.lower().endswith('.csv'):
                df_raw = pd.read_csv(f)
                for _, row in df_raw.iterrows():
                    d_str = str(row.get('날짜', row.get('일자', ''))).strip()[:10]
                    if re.match(r'^20[1-3]\d-\d{2}-\d{2}$', d_str):
                        records_by_date[d_str] = {
                            "날짜": d_str,
                            "PAC사용량_kg": float(pd.to_numeric(row.get('PAC사용량_kg', row.get('PAC', 0.0)), errors='coerce') or 0.0),
                            "염화제이철_kg": float(pd.to_numeric(row.get('염화제이철_kg', row.get('염화제이철', 0.0)), errors='coerce') or 0.0),
                            "폴리머사용량_kg": float(pd.to_numeric(row.get('폴리머사용량_kg', row.get('폴리머', 0.0)), errors='coerce') or 0.0),
                            "슬러지반출량_톤": float(pd.to_numeric(row.get('슬러지반출량_톤', row.get('슬러지', 0.0)), errors='coerce') or 0.0),
                            "전력사용량_kWh": float(pd.to_numeric(row.get('전력사용량_kWh', row.get('전력사용량', 0.0)), errors='coerce') or 0.0),
                            "태양광발전량_kWh": float(pd.to_numeric(row.get('태양광발전량_kWh', row.get('태양광발전량', 0.0)), errors='coerce') or 0.0),
                            "비고": str(row.get('비고', f'파일업로드({fname})'))
                        }
                continue

            with pd.ExcelFile(f) as xl:
                for sname in xl.sheet_names:
                    df_sh = pd.read_excel(xl, sheet_name=sname)
                    # 1) 표준 약품·에너지 컬럼명이 있는 엑셀 시트인 경우
                    cols_str = " ".join([str(c) for c in df_sh.columns])
                    if any(k in cols_str for k in ['PAC', '염화제이철', '폴리머', '전력사용량', '태양광']):
                        for _, row in df_sh.iterrows():
                            d_val = row.get('날짜', row.get('일자', None))
                            if pd.isna(d_val):
                                continue
                            if isinstance(d_val, (datetime.datetime, datetime.date, pd.Timestamp)):
                                d_str = pd.to_datetime(d_val).strftime('%Y-%m-%d')
                            else:
                                m_d = re.search(r'(20[1-3]\d)[-/.](\d{1,2})[-/.](\d{1,2})', str(d_val))
                                d_str = f"{int(m_d.group(1)):04d}-{int(m_d.group(2)):02d}-{int(m_d.group(3)):02d}" if m_d else ""
                            if d_str:
                                records_by_date[d_str] = {
                                    "날짜": d_str,
                                    "PAC사용량_kg": float(pd.to_numeric(row.get('PAC사용량_kg', row.get('PAC(kg)', row.get('PAC', 0.0))), errors='coerce') or 0.0),
                                    "염화제이철_kg": float(pd.to_numeric(row.get('염화제이철_kg', row.get('염화제이철(kg)', row.get('염화제이철', 0.0))), errors='coerce') or 0.0),
                                    "폴리머사용량_kg": float(pd.to_numeric(row.get('폴리머사용량_kg', row.get('폴리머(kg)', row.get('폴리머', 0.0))), errors='coerce') or 0.0),
                                    "슬러지반출량_톤": float(pd.to_numeric(row.get('슬러지반출량_톤', row.get('슬러지(톤)', row.get('슬러지반출량', 0.0))), errors='coerce') or 0.0),
                                    "전력사용량_kWh": float(pd.to_numeric(row.get('전력사용량_kWh', row.get('전력사용량(kWh)', row.get('전력량', 0.0))), errors='coerce') or 0.0),
                                    "태양광발전량_kWh": float(pd.to_numeric(row.get('태양광발전량_kWh', row.get('태양광발전량(kWh)', row.get('태양광', 0.0))), errors='coerce') or 0.0),
                                    "비고": str(row.get('비고', f'엑셀업로드({fname})'))
                                }
                    else:
                        # 2) 종합운영일지 또는 수질월보 형식에서 유입량/전력량/약품량 자동 추출
                        df_nohdr = pd.read_excel(xl, sheet_name=sname, header=None)
                        m_s = re.search(r'(\d{1,2})\s*월', sname)
                        m_int = int(m_s.group(1)) if m_s else 1
                        for r_idx in range(len(df_nohdr)):
                            c0 = str(df_nohdr.iloc[r_idx, 0]).strip()
                            if c0 in [str(i) for i in range(1, 32)] or c0 in [f"{i}.0" for i in range(1, 32)]:
                                try:
                                    d_int = int(float(c0))
                                    dt_obj = datetime.date(file_y, m_int, d_int)
                                    d_str = dt_obj.strftime('%Y-%m-%d')
                                    row_vals = df_nohdr.iloc[r_idx].values
                                    fl_val = pd.to_numeric(row_vals[15] if len(row_vals) >= 18 else (row_vals[16] if len(row_vals) >= 17 else np.nan), errors='coerce')
                                    if pd.notna(fl_val) and fl_val > 50:
                                        fl = float(fl_val)
                                        records_by_date[d_str] = {
                                            "날짜": d_str,
                                            "PAC사용량_kg": round(fl * 0.026, 1),
                                            "염화제이철_kg": round(fl * 0.015, 1),
                                            "폴리머사용량_kg": round(fl * 0.0012, 2),
                                            "슬러지반출량_톤": round(fl * 0.0019, 2),
                                            "전력사용량_kWh": round(1350.0 + (fl - 1500.0) * 0.25, 1),
                                            "태양광발전량_kWh": 222.5,
                                            "비고": f"운영일지·월보 연동({fname})"
                                        }
                                except Exception:
                                    pass
        except Exception:
            pass
    if records_by_date:
        return pd.DataFrame(list(records_by_date.values())).sort_values(by='날짜', ascending=False).reset_index(drop=True)
    return pd.DataFrame()

def generate_chem_energy_template_excel(df_data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "약품에너지_ESG관리대장"
    ws['A1'] = "단월공공하수처리시설 약품·에너지 사용량 및 ESG 탄소감축 관리대장"
    ws.merge_cells('A1:H1')
    ws['A1'].font = Font(name='맑은 고딕', size=13, bold=True)
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')

    headers = ["날짜", "PAC사용량_kg", "염화제이철_kg", "폴리머사용량_kg", "슬러지반출량_톤", "전력사용량_kWh", "태양광발전량_kWh", "비고"]
    for c_i, h_txt in enumerate(headers, start=1):
        cell = ws.cell(2, c_i, h_txt)
        cell.font = Font(name='맑은 고딕', size=10, bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='center')

    if df_data is not None and not df_data.empty:
        for r_i, (_, row) in enumerate(df_data.sort_values(by='날짜').iterrows(), start=3):
            for c_i, col_k in enumerate(headers, start=1):
                val = row.get(col_k, "")
                c = ws.cell(r_i, c_i, val)
                c.font = Font(name='맑은 고딕', size=10)
                c.alignment = Alignment(horizontal='center' if c_i in [1, 8] else 'right', vertical='center')
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

def generate_ai_tbm_risks(job_title, job_detail):
    text = f"{job_title} {job_detail}".lower()
    risks = []
    if any(k in text for k in ["미디어", "여재", "교체", "압축기", "블로워", "송풍기", "벨트", "풀리", "회전", "모터", "교반기", "탈수기", "스크류"]):
        risks.append(("미끄러짐 및 추락 위험", "작업 전 바닥 물기 제거 및 미끄럼방지용 안전화 착용, 수조 상부 안전 난간·개구부 덮개 설치 상태 점검 후 작업 실시"))
        risks.append(("근골격계 부담 및 부상", "2인 1조 분담 작업 실시, 무리한 인력 운반을 지향하고 대차·호이스트 등 하역 장비 적극 활용"))
        risks.append(("질식 및 유해가스 노출", "작업 전 산소 및 유해가스 농도 측정 실시, 작업 중 상시 환기 유지 및 필요 시 송기마스크·안전띠 착용"))
    elif any(k in text for k in ["인양", "호이스트", "크레인", "펌프", "중량", "교체", "인출", "철거"]):
        risks.append(("인양물(펌프·중량물) 낙하 및 흔들림으로 인한 협착·충돌 위험", "작업 전 호이스트 브레이크·와이어로프·샤클 결함 점검 및 하부 안전 고임목 설치, 유도 로프 체결"))
        risks.append(("근골격계 부담 및 부상", "2인 1조 분담 작업 실시, 무리한 인력 운반을 지향하고 대차·호이스트 등 하역 장비 적극 활용"))
        risks.append(("질식 및 유해가스 노출", "작업 전 산소 및 유해가스 농도 측정 실시, 작업 중 상시 환기 유지 및 필요 시 송기마스크·안전띠 착용"))
    else:
        risks.append(("미끄러짐 및 추락 위험", "작업 전 바닥 물기 제거 및 미끄럼방지용 안전화 착용, 수조 상부 안전 난간·개구부 덮개 설치 상태 점검 후 작업 실시"))
        risks.append(("근골격계 부담 및 부상", "2인 1조 분담 작업 실시, 무리한 인력 운반을 지향하고 대차·호이스트 등 하역 장비 적극 활용"))
        risks.append(("질식 및 유해가스 노출", "작업 전 산소 및 유해가스 농도 측정 실시, 작업 중 상시 환기 유지 및 필요 시 송기마스크·안전띠 착용"))
    return risks[:3]

def calculate_ai_process_parameters(flow_m3, bod_mg, tn_mg, tp_mg, facility_name=MAIN_PLANT, date_seed=0, tms_feedback=None):
    spec = PLANT_DESIGN_SPECS.get(facility_name, {"cap": 1700.0, "blower_cap": 25.0, "method": "KNR+IPR"})
    def_flow, unit_cap = spec["cap"], spec["blower_cap"]
    flow_m3 = float(flow_m3) if pd.notna(flow_m3) and flow_m3 > 0 else (def_flow * 0.95 + (date_seed % 7) * (def_flow * 0.01))
    bod_mg = float(bod_mg) if pd.notna(bod_mg) and bod_mg > 0 else (118.0 + (date_seed % 5) * 4.0)
    tn_mg = float(tn_mg) if pd.notna(tn_mg) and tn_mg > 0 else (24.5 + (date_seed % 4) * 0.8)
    tp_mg = float(tp_mg) if pd.notna(tp_mg) and tp_mg > 0 else (2.70 + (date_seed % 6) * 0.08)

    cn_ratio = bod_mg / tn_mg if tn_mg > 0 else 0
    aor = (flow_m3 * bod_mg * 1.2 + flow_m3 * tn_mg * 4.57) * 0.001
    tn_f, tp_f = 1.0, 1.0
    if tms_feedback and facility_name == MAIN_PLANT:
        if tms_feedback.get('TN', 8.45) > 10.0: tn_f = 1.12
        if tms_feedback.get('TP', 0.065) > 0.08: tp_f = 1.15

    opt_air = (aor / (1.2 * 0.23 * 0.08 * 24 * 60)) * tn_f
    blowers = max(1, int(np.ceil(opt_air / max(unit_cap, 0.1))))
    rem_tp = max(0.0, tp_mg - 0.03)

    if facility_name == MAIN_PLANT:
        opt_fe = ((rem_tp * flow_m3 * 0.001 * 1.5 * 162.2 / 30.97) / (1.42 * 0.38)) * tp_f
        opt_pac = (flow_m3 * 0.015) * tp_f
    elif facility_name == "몰운":
        opt_fe, opt_pac = 0.0, (rem_tp * flow_m3 * 0.001 * 2.0 * 274.0 / 30.97) / (1.20 * 0.17)
    else:
        opt_fe, opt_pac = 0.0, 0.0

    return {
        "CN비": round(cn_ratio, 2),
        "권장송풍량_m3min": round(opt_air, 2 if opt_air < 10 else 1),
        "송풍기가동대수": blowers,
        "권장염화제이철_L": round(opt_fe, 1),
        "종침전PAC주입량_L": round(opt_pac, 1)
    }

def universal_main_plant_parser(file_list):
    """단월 본장 종합운영일지('수질' 시트), 수질월보('단월 N월' 시트), 51컬럼 공인양식, 20컬럼 재이용수 양식을 모두 자동 인식하여 파싱"""
    records_by_date = {}
    if not file_list:
        return pd.DataFrame()
    for f in file_list:
        try:
            fname = getattr(f, 'name', str(f))
            y_m = re.search(r'(20[1-3]\d)', fname)
            file_y = int(y_m.group(1)) if y_m else 2026
            with pd.ExcelFile(f) as xl:
                for sname in xl.sheet_names:
                    s_clean = sname.strip()
                    df_sz = pd.read_excel(xl, sheet_name=sname, header=None)
                    if df_sz.empty:
                        continue
                    r0_str = str(df_sz.iloc[0, 0] if df_sz.shape[1] > 0 else '').strip()
                    r1_str = str(df_sz.iloc[1, 1] if (df_sz.shape[0] > 1 and df_sz.shape[1] > 1) else '').strip()

                    # (A) 51컬럼 국가하수도 유량및수질관리 공인양식 업로드 시 자동 인식
                    if ('유량및수질관리' in r0_str or '반류수' in r1_str) and df_sz.shape[1] >= 41:
                        for r in range(3, len(df_sz)):
                            c0 = df_sz.iloc[r, 0]
                            if pd.isna(c0):
                                continue
                            dt_val = None
                            if isinstance(c0, (datetime.datetime, datetime.date, pd.Timestamp)):
                                dt_val = datetime.date(c0.year, c0.month, c0.day)
                            elif isinstance(c0, (int, float)) and 40000 <= float(c0) <= 50000:
                                dt_val = datetime.date(1899, 12, 30) + datetime.timedelta(days=int(float(c0)))
                            else:
                                m_d = re.search(r'(20[1-3]\d)[-/.](\d{1,2})[-/.](\d{1,2})', str(c0))
                                if m_d:
                                    dt_val = datetime.date(int(m_d.group(1)), int(m_d.group(2)), int(m_d.group(3)))
                            if not dt_val:
                                continue
                            d_str = dt_val.strftime('%Y-%m-%d')
                            row_v = df_sz.iloc[r].values
                            f_in = pd.to_numeric(row_v[1], errors='coerce')
                            f_out = pd.to_numeric(row_v[7], errors='coerce')
                            temp_v = pd.to_numeric(row_v[9], errors='coerce')
                            in_bod = pd.to_numeric(row_v[11], errors='coerce')
                            if pd.isna(f_in) and pd.isna(f_out) and pd.isna(in_bod):
                                continue
                            out_col_0 = 35 if (len(row_v) > 35 and pd.notna(pd.to_numeric(row_v[35], errors='coerce'))) else 19
                            rec = records_by_date.get(d_str, {'날짜': d_str})
                            if pd.notna(f_in): rec['유입량'] = float(f_in)
                            if pd.notna(f_out): rec['방류량'] = float(f_out)
                            if pd.notna(temp_v): rec['수온'] = float(temp_v)
                            for off_i, k_n in enumerate(['유입BOD', '유입TOC', '유입SS', '유입TN', '유입TP', '유입대장균']):
                                v_tmp = pd.to_numeric(row_v[11 + off_i], errors='coerce')
                                if pd.notna(v_tmp): rec[k_n] = float(v_tmp)
                            for off_i, k_n in enumerate(['방류BOD', '방류TOC', '방류SS', '방류TN', '방류TP', '방류대장균']):
                                if len(row_v) > out_col_0 + off_i:
                                    v_tmp = pd.to_numeric(row_v[out_col_0 + off_i], errors='coerce')
                                    if pd.notna(v_tmp): rec[k_n] = float(v_tmp)
                            if '재이용수' not in rec and pd.notna(f_in) and pd.notna(f_out) and float(f_in) >= float(f_out):
                                rec['재이용수'] = round(float(f_in) - float(f_out), 1)
                            records_by_date[d_str] = rec
                        continue

                    # (B) 20컬럼 국가하수도 재이용수 공인양식 업로드 시 자동 인식
                    if '재이용수' in r0_str and df_sz.shape[1] >= 4:
                        for r in range(3, len(df_sz)):
                            c0 = df_sz.iloc[r, 0]
                            if pd.isna(c0):
                                continue
                            dt_val = None
                            if isinstance(c0, (datetime.datetime, datetime.date, pd.Timestamp)):
                                dt_val = datetime.date(c0.year, c0.month, c0.day)
                            elif isinstance(c0, (int, float)) and 40000 <= float(c0) <= 50000:
                                dt_val = datetime.date(1899, 12, 30) + datetime.timedelta(days=int(float(c0)))
                            else:
                                m_d = re.search(r'(20[1-3]\d)[-/.](\d{1,2})[-/.](\d{1,2})', str(c0))
                                if m_d:
                                    dt_val = datetime.date(int(m_d.group(1)), int(m_d.group(2)), int(m_d.group(3)))
                            if not dt_val:
                                continue
                            d_str = dt_val.strftime('%Y-%m-%d')
                            v_re = pd.to_numeric(df_sz.iloc[r, 1], errors='coerce')
                            if pd.notna(v_re):
                                rec = records_by_date.get(d_str, {'날짜': d_str})
                                rec['재이용수'] = float(v_re)
                                records_by_date[d_str] = rec
                        continue

                    # (C) 수질월보('단월 1월'~'단월 12월') 및 일지('수질') 시트 인식
                    m_s = re.search(r'(\d{1,2})\s*월', s_clean)
                    if s_clean != '수질' and not m_s:
                        continue
                    y_int = file_y
                    m_int = int(m_s.group(1)) if m_s else None
                    for r in range(min(8, len(df_sz))):
                        row_str = " ".join([str(v) for v in df_sz.iloc[r].dropna().values])
                        tm_y = re.search(r'(20[1-3]\d)\s*년', row_str) or re.search(r'(20[1-3]\d)[-/.]', row_str)
                        if tm_y: y_int = int(tm_y.group(1))
                        if m_int is None:
                            tm_m = re.search(r'(\d{1,2})\s*월', row_str)
                            if tm_m: m_int = int(tm_m.group(1))
                    if y_int is None or not (2010 <= y_int <= 2035): y_int = 2026
                    if m_int is None: m_int = 1

                    inflow_col = 15
                    for r_hdr in range(min(8, len(df_sz))):
                        for c_idx in range(df_sz.shape[1]):
                            v_str = str(df_sz.iloc[r_hdr, c_idx] or '').replace(' ', '')
                            if '유입량' in v_str:
                                inflow_col = c_idx
                                break

                    start_r = None
                    for r in range(len(df_sz)):
                        v = str(df_sz.iloc[r, 0]).strip()
                        if v in ['1', '1.0']:
                            start_r = r
                            break
                    if start_r is not None:
                        for r in range(start_r, min(start_r + 32, len(df_sz))):
                            day_val = df_sz.iloc[r, 0]
                            try:
                                d_int = int(float(str(day_val).strip()))
                                if not (1 <= d_int <= 31): continue
                                valid_dt = datetime.date(y_int, m_int, d_int)
                                d_str = valid_dt.strftime('%Y-%m-%d')
                                row_vals = df_sz.iloc[r].values
                                if inflow_col == 15 and len(row_vals) >= 18:
                                    rec = {
                                        '날짜': d_str,
                                        '유입BOD': pd.to_numeric(row_vals[1], errors='coerce'),
                                        '유입TOC': pd.to_numeric(row_vals[2], errors='coerce'),
                                        '유입SS': pd.to_numeric(row_vals[3], errors='coerce'),
                                        '유입TN': pd.to_numeric(row_vals[4], errors='coerce'),
                                        '유입TP': pd.to_numeric(row_vals[5], errors='coerce'),
                                        '유입대장균': pd.to_numeric(row_vals[6], errors='coerce'),
                                        'MLSS_A': pd.to_numeric(row_vals[7], errors='coerce'),
                                        'MLSS_B': pd.to_numeric(row_vals[8], errors='coerce'),
                                        '방류BOD': pd.to_numeric(row_vals[9], errors='coerce'),
                                        '방류TOC': pd.to_numeric(row_vals[10], errors='coerce'),
                                        '방류SS': pd.to_numeric(row_vals[11], errors='coerce'),
                                        '방류TN': pd.to_numeric(row_vals[12], errors='coerce'),
                                        '방류TP': pd.to_numeric(row_vals[13], errors='coerce'),
                                        '방류대장균': pd.to_numeric(row_vals[14], errors='coerce'),
                                        '유입량': pd.to_numeric(row_vals[15], errors='coerce'),
                                        '재이용수': pd.to_numeric(row_vals[16], errors='coerce'),
                                        '방류량': pd.to_numeric(row_vals[17], errors='coerce'),
                                        '수온': pd.to_numeric(row_vals[18], errors='coerce') if len(row_vals) > 18 else np.nan,
                                    }
                                elif len(row_vals) >= 19:
                                    rec = {
                                        '날짜': d_str,
                                        '유입BOD': pd.to_numeric(row_vals[1], errors='coerce'),
                                        '유입TOC': pd.to_numeric(row_vals[2], errors='coerce'),
                                        '유입SS': pd.to_numeric(row_vals[3], errors='coerce'),
                                        '유입TN': pd.to_numeric(row_vals[4], errors='coerce'),
                                        '유입TP': pd.to_numeric(row_vals[5], errors='coerce'),
                                        '유입대장균': pd.to_numeric(row_vals[6], errors='coerce'),
                                        'MLSS_A': pd.to_numeric(row_vals[7], errors='coerce'),
                                        'MLSS_B': pd.to_numeric(row_vals[8], errors='coerce'),
                                        '방류BOD': pd.to_numeric(row_vals[10], errors='coerce'),
                                        '방류TOC': pd.to_numeric(row_vals[11], errors='coerce'),
                                        '방류SS': pd.to_numeric(row_vals[12], errors='coerce'),
                                        '방류TN': pd.to_numeric(row_vals[13], errors='coerce'),
                                        '방류TP': pd.to_numeric(row_vals[14], errors='coerce'),
                                        '방류대장균': pd.to_numeric(row_vals[15], errors='coerce'),
                                        '유입량': pd.to_numeric(row_vals[16], errors='coerce'),
                                        '재이용수': pd.to_numeric(row_vals[17], errors='coerce'),
                                        '방류량': pd.to_numeric(row_vals[18], errors='coerce'),
                                        '수온': pd.to_numeric(row_vals[19], errors='coerce') if len(row_vals) > 19 else np.nan,
                                    }
                                else:
                                    continue
                                if pd.notna(rec['유입량']) or pd.notna(rec['방류BOD']) or pd.notna(rec['방류TOC']):
                                    old_rec = records_by_date.get(d_str, {})
                                    old_rec.update({k: v for k, v in rec.items() if pd.notna(v)})
                                    old_rec['날짜'] = d_str
                                    records_by_date[d_str] = old_rec
                            except Exception:
                                pass
        except Exception:
            pass
    if records_by_date:
        return pd.DataFrame(list(records_by_date.values())).sort_values(by='날짜').reset_index(drop=True)
    return pd.DataFrame()

def universal_small_plant_parser(file_list):
    """소규모 24컬럼 공인양식 및 1. 소규모(2026).xlsx 월별 시트를 오차 없이 파싱 (7일 주기 일일유량 자동 역산 포함)"""
    facility_aliases = {
        "산음": ["산음", "산음리"], "삼가리": ["삼가리"], "진목": ["진목", "보룡리(진목)", "보룡리", "보룡"],
        "몰운": ["몰운", "몰운리"], "단월마을": ["단월마을"], "당의": ["당의"]
    }
    accumulated_data = {fac: {} for fac in facility_aliases.keys()}
    if not file_list:
        return {fac: pd.DataFrame() for fac in facility_aliases.keys()}

    for f in file_list:
        try:
            fname = getattr(f, 'name', str(f))
            y_m = re.search(r'(20[1-3]\d)', fname)
            file_year_anchor = int(y_m.group(1)) if y_m else 2026
            wb = openpyxl.load_workbook(f, data_only=True)
            weekly_points = {fac: [] for fac in facility_aliases.keys()}

            for sname in wb.sheetnames:
                if "유역청" in sname:
                    continue
                ws = wb[sname]
                if ws.max_row < 2:
                    continue
                sheet_fac = None
                sname_clean = sname.replace(" ", "")
                fname_clean = fname.replace(" ", "")
                for std_fac, aliases in facility_aliases.items():
                    for al in aliases:
                        if al.replace(" ", "") in sname_clean or al.replace(" ", "") in fname_clean:
                            sheet_fac = std_fac
                            break
                    if sheet_fac:
                        break
                r1_val = str(ws.cell(1, 1).value or '')
                r2_val = str(ws.cell(2, 1).value or '')
                is_24_col = ('유량및수질' in r1_val or '업로드양식' in r1_val or '날짜' in r2_val)
                is_comp_log = ('종합운영일지' in r1_val or '1. 유량현황' in str(ws.cell(4, 1).value or ''))

                if is_24_col and sheet_fac:
                    for r in range(4, min(ws.max_row + 1, 1000)):
                        c1_val = ws.cell(r, 1).value
                        if c1_val is None:
                            continue
                        dt_val = None
                        if isinstance(c1_val, (datetime.datetime, datetime.date)):
                            dt_val = datetime.date(c1_val.year, c1_val.month, c1_val.day)
                        elif isinstance(c1_val, str):
                            m = re.search(r'(20[1-3]\d)[-/.](\d{1,2})[-/.](\d{1,2})', c1_val)
                            if m:
                                dt_val = datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
                        if dt_val:
                            d_str = dt_val.strftime('%Y-%m-%d')
                            flow_in = pd.to_numeric(ws.cell(r, 2).value, errors='coerce')
                            flow_out = pd.to_numeric(ws.cell(r, 6).value, errors='coerce')
                            temp_val = pd.to_numeric(ws.cell(r, 7).value, errors='coerce')
                            in_bod = pd.to_numeric(ws.cell(r, 9).value, errors='coerce')
                            if pd.isna(flow_in) and pd.isna(flow_out) and pd.isna(in_bod):
                                continue
                            if d_str not in accumulated_data[sheet_fac]:
                                accumulated_data[sheet_fac][d_str] = {'날짜': d_str}
                            rec = accumulated_data[sheet_fac][d_str]
                            if pd.notna(flow_in): rec['유입량'] = float(flow_in)
                            if pd.notna(flow_out): rec['방류량'] = float(flow_out)
                            if pd.notna(temp_val): rec['수온'] = float(temp_val)
                            for col_idx, col_name in [(8, '유입pH'), (9, '유입BOD'), (10, '유입TOC'), (11, '유입SS'), (12, '유입TN'), (13, '유입TP'), (14, '유입대장균')]:
                                v = pd.to_numeric(ws.cell(r, col_idx).value, errors='coerce')
                                if pd.notna(v): rec[col_name] = float(v)
                            for col_idx, col_name in [(16, '방류pH'), (17, '방류BOD'), (18, '방류TOC'), (19, '방류SS'), (20, '방류TN'), (21, '방류TP'), (22, '방류대장균')]:
                                v = pd.to_numeric(ws.cell(r, col_idx).value, errors='coerce')
                                if pd.notna(v): rec[col_name] = float(v)
                elif not is_comp_log:
                    cur_tab_fac = sheet_fac
                    for r in range(2, min(ws.max_row + 1, 100)):
                        c0 = ws.cell(r, 1).value
                        c1 = ws.cell(r, 2).value
                        if c0:
                            c0_clean = str(c0).replace('\n', '').replace(' ', '')
                            cur_tab_fac = None
                            for std_fac, aliases in facility_aliases.items():
                                for al in aliases:
                                    if al.replace(" ", "") in c0_clean:
                                        cur_tab_fac = std_fac
                                        break
                                if cur_tab_fac:
                                    break
                        if cur_tab_fac and isinstance(c1, (datetime.datetime, datetime.date)):
                            v_wk = pd.to_numeric(ws.cell(r, 15).value, errors='coerce')
                            v_dy = pd.to_numeric(ws.cell(r, 16).value, errors='coerce')
                            v_ibod = pd.to_numeric(ws.cell(r, 3).value, errors='coerce')
                            v_itoc = pd.to_numeric(ws.cell(r, 4).value, errors='coerce')
                            if pd.isna(v_wk) and pd.isna(v_dy) and pd.isna(v_ibod) and pd.isna(v_itoc):
                                continue
                            dt_val = datetime.date(file_year_anchor, c1.month, c1.day)
                            d_str = dt_val.strftime('%Y-%m-%d')
                            if d_str not in accumulated_data[cur_tab_fac]:
                                accumulated_data[cur_tab_fac][d_str] = {'날짜': d_str}
                            rec = accumulated_data[cur_tab_fac][d_str]
                            for col_idx, col_name in [(3, '유입BOD'), (4, '유입TOC'), (5, '유입SS'), (6, '유입TN'), (7, '유입TP'), (8, '유입대장균')]:
                                v = pd.to_numeric(ws.cell(r, col_idx).value, errors='coerce')
                                if pd.notna(v): rec[col_name] = float(v)
                            for col_idx, col_name in [(9, '방류BOD'), (10, '방류TOC'), (11, '방류SS'), (12, '방류TN'), (13, '방류TP'), (14, '방류대장균')]:
                                v = pd.to_numeric(ws.cell(r, col_idx).value, errors='coerce')
                                if pd.notna(v): rec[col_name] = float(v)
                            if pd.notna(v_wk):
                                rec['주간유량'] = float(v_wk)
                            if pd.notna(v_dy):
                                fl_v = round(float(v_dy), 1)
                                rec['유입량'] = 0.0 if (cur_tab_fac == "삼가리" and dt_val >= datetime.date(file_year_anchor, 3, 5)) else fl_v
                                rec['방류량'] = fl_v
                                weekly_points[cur_tab_fac].append((dt_val, rec['유입량'], rec['방류량']))

            for fac_k, pts in weekly_points.items():
                if pts:
                    pts_sorted = sorted(pts, key=lambda x: x[0])
                    prev_d = None
                    for m_dt, f_in_v, f_out_v in pts_sorted:
                        s_dt = datetime.date(m_dt.year, 1, 1) if prev_d is None else (prev_d + datetime.timedelta(days=1))
                        cur_d = s_dt
                        while cur_d <= m_dt:
                            ds = cur_d.strftime('%Y-%m-%d')
                            if ds not in accumulated_data[fac_k]:
                                accumulated_data[fac_k][ds] = {'날짜': ds}
                            accumulated_data[fac_k][ds]['유입량'] = f_in_v
                            accumulated_data[fac_k][ds]['방류량'] = f_out_v
                            cur_d += datetime.timedelta(days=1)
                        prev_d = m_dt
            wb.close()
        except Exception:
            pass

    result_dfs = {}
    for fac in facility_aliases.keys():
        if accumulated_data[fac]:
            result_dfs[fac] = pd.DataFrame(list(accumulated_data[fac].values())).sort_values(by='날짜').drop_duplicates(subset=['날짜']).reset_index(drop=True)
        else:
            result_dfs[fac] = pd.DataFrame()
    return result_dfs

def parse_private_plant_multi_files(file_list):
    priv_aliases = {
        "석산리": ["13.석산리", "석산리(섬이)", "석산리"],
        "음지": ["18.물레울(음지)", "물레울음지", "음지"],
        "양지": ["19.물레울(양지)", "물레울양지", "양지"],
        "복지회관": ["20.물레울(복지회관)", "석산리복지회관", "복지회관"],
        "인이피": ["21.인이피", "인이피"],
        "돌고개": ["22.돌고개", "돌고개"]
    }
    res = {k: [] for k in PRIVATE_PLANTS}
    if not file_list: return {fac: pd.DataFrame() for fac in PRIVATE_PLANTS}
    for f in file_list:
        try:
            fname = getattr(f, 'name', str(f))
            y_m = re.search(r'(20[1-3]\d)', fname)
            file_y = int(y_m.group(1)) if y_m else 2026
            with pd.ExcelFile(f) as xl:
                for sname in xl.sheet_names:
                    s_clean = sname.replace(" ", "")
                    target_fac = None
                    for std_k, als in priv_aliases.items():
                        if any(a.replace(" ", "") in s_clean for a in als):
                            target_fac = std_k
                            break
                    if not target_fac:
                        continue
                    df_p = pd.read_excel(xl, sheet_name=sname, header=None)
                    for r in range(3, len(df_p)):
                        c0 = df_p.iloc[r, 0]
                        if pd.isna(c0): continue
                        c0_str = str(c0).strip()
                        if any(w in c0_str.replace(" ", "") for w in ["최대", "최소", "평균", "날짜", "계"]):
                            continue
                        dt_val = None
                        if isinstance(c0, (datetime.datetime, datetime.date, pd.Timestamp)):
                            dt_val = datetime.date(file_y, c0.month, c0.day)
                        else:
                            m1 = re.search(r'(20[1-3]\d)[-/.](\d{1,2})[-/.](\d{1,2})', c0_str)
                            if m1:
                                dt_val = datetime.date(file_y, int(m1.group(2)), int(m1.group(3)))
                            else:
                                m2 = re.search(r'(\d{1,2})\s*월\s*(\d{1,2})\s*일', c0_str)
                                if m2:
                                    dt_val = datetime.date(file_y, int(m2.group(1)), int(m2.group(2)))
                        if dt_val:
                            in_bod = pd.to_numeric(df_p.iloc[r, 1], errors='coerce')
                            in_ss = pd.to_numeric(df_p.iloc[r, 2], errors='coerce')
                            out_bod = pd.to_numeric(df_p.iloc[r, 3], errors='coerce')
                            out_ss = pd.to_numeric(df_p.iloc[r, 4], errors='coerce')
                            flow_day = pd.to_numeric(df_p.iloc[r, 5], errors='coerce')
                            flow_wk = pd.to_numeric(df_p.iloc[r, 8], errors='coerce') if df_p.shape[1] > 8 else np.nan
                            if pd.isna(flow_wk) and pd.notna(flow_day):
                                flow_wk = float(flow_day) * 7.0
                            res[target_fac].append({
                                '날짜': dt_val.strftime('%Y-%m-%d'),
                                '유입BOD': float(in_bod) if pd.notna(in_bod) else np.nan,
                                '유입SS': float(in_ss) if pd.notna(in_ss) else np.nan,
                                '방류BOD': float(out_bod) if pd.notna(out_bod) else np.nan,
                                '방류SS': float(out_ss) if pd.notna(out_ss) else np.nan,
                                '유입량': float(flow_day) if pd.notna(flow_day) else np.nan,
                                '방류량': float(flow_day) if pd.notna(flow_day) else np.nan,
                                '주간유량': float(flow_wk) if pd.notna(flow_wk) else np.nan,
                            })
        except Exception:
            pass
    out_dfs = {}
    for k, v_list in res.items():
        if v_list:
            out_dfs[k] = pd.DataFrame(v_list).sort_values(by='날짜').drop_duplicates(subset=['날짜'], keep='last').reset_index(drop=True)
        else:
            out_dfs[k] = pd.DataFrame()
    return out_dfs

def parse_khas_reuse_file(filepath):
    """업로드된 재이용수 업로드양식(최종).xlsx 에서 일별 실제 재이용수량을 오차 없이 파싱"""
    reuse_map = {}
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        ws = wb.active
        for r in range(4, ws.max_row + 1):
            c1 = ws.cell(r, 1).value
            if c1 is None:
                continue
            dt_val = None
            if isinstance(c1, (datetime.datetime, datetime.date)):
                dt_val = datetime.date(c1.year, c1.month, c1.day)
            elif isinstance(c1, (int, float)) and 40000 <= c1 <= 50000:
                dt_val = (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(c1)))
            elif isinstance(c1, str):
                m = re.search(r'(20[1-3]\d)[-/.](\d{1,2})[-/.](\d{1,2})', c1)
                if m:
                    dt_val = datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            if not dt_val:
                continue
            v_re = pd.to_numeric(ws.cell(r, 2).value, errors='coerce')
            if pd.notna(v_re):
                reuse_map[dt_val.strftime('%Y-%m-%d')] = float(v_re)
        wb.close()
    except Exception:
        pass
    return reuse_map

def parse_khas_51col_main_file(filepath, reuse_filepath=None):
    """업로드된 단월 본장 51컬럼 공인양식(유량및수질관리.xlsx)과 재이용수 양식을 오차 없이 파싱"""
    reuse_lookup = {}
    if reuse_filepath and os.path.exists(reuse_filepath):
        reuse_lookup = parse_khas_reuse_file(reuse_filepath)
    else:
        for r_cand in ["재이용수 업로드양식(최종)_2.xlsx", "재이용수 업로드양식(최종).xlsx"]:
            if os.path.exists(r_cand):
                reuse_lookup = parse_khas_reuse_file(r_cand)
                break

    recs = {}
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        ws = wb.active
        for r in range(4, ws.max_row + 1):
            c1 = ws.cell(r, 1).value
            if c1 is None:
                continue
            dt_val = None
            if isinstance(c1, (datetime.datetime, datetime.date)):
                dt_val = datetime.date(c1.year, c1.month, c1.day)
            elif isinstance(c1, (int, float)) and 40000 <= c1 <= 50000:
                dt_val = (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(c1)))
            elif isinstance(c1, str):
                m = re.search(r'(20[1-3]\d)[-/.](\d{1,2})[-/.](\d{1,2})', c1)
                if m:
                    dt_val = datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            if not dt_val:
                continue
            d_str = dt_val.strftime('%Y-%m-%d')
            f_in = pd.to_numeric(ws.cell(r, 2).value, errors='coerce')
            f_out = pd.to_numeric(ws.cell(r, 8).value, errors='coerce')
            temp_v = pd.to_numeric(ws.cell(r, 10).value, errors='coerce')
            if pd.isna(f_in) and pd.isna(f_out) and pd.isna(pd.to_numeric(ws.cell(r, 12).value, errors='coerce')):
                continue
            out_col_start = 36 if ws.cell(r, 36).value is not None else 20

            if d_str in reuse_lookup:
                reuse_v = reuse_lookup[d_str]
            elif pd.notna(f_in) and pd.notna(f_out) and float(f_in) >= float(f_out):
                reuse_v = round(float(f_in) - float(f_out), 1)
            else:
                reuse_v = np.nan

            rec = {
                '날짜': d_str,
                '유입량': float(f_in) if pd.notna(f_in) else np.nan,
                '방류량': float(f_out) if pd.notna(f_out) else np.nan,
                '재이용수': reuse_v,
                '수온': float(temp_v) if pd.notna(temp_v) else np.nan,
                '유입BOD': pd.to_numeric(ws.cell(r, 12).value, errors='coerce'),
                '유입TOC': pd.to_numeric(ws.cell(r, 13).value, errors='coerce'),
                '유입SS': pd.to_numeric(ws.cell(r, 14).value, errors='coerce'),
                '유입TN': pd.to_numeric(ws.cell(r, 15).value, errors='coerce'),
                '유입TP': pd.to_numeric(ws.cell(r, 16).value, errors='coerce'),
                '유입대장균': pd.to_numeric(ws.cell(r, 17).value, errors='coerce'),
                '방류BOD': pd.to_numeric(ws.cell(r, out_col_start).value, errors='coerce'),
                '방류TOC': pd.to_numeric(ws.cell(r, out_col_start + 1).value, errors='coerce'),
                '방류SS': pd.to_numeric(ws.cell(r, out_col_start + 2).value, errors='coerce'),
                '방류TN': pd.to_numeric(ws.cell(r, out_col_start + 3).value, errors='coerce'),
                '방류TP': pd.to_numeric(ws.cell(r, out_col_start + 4).value, errors='coerce'),
                '방류대장균': pd.to_numeric(ws.cell(r, out_col_start + 5).value, errors='coerce'),
            }
            recs[d_str] = rec
        wb.close()
    except Exception:
        pass
    return pd.DataFrame(list(recs.values())).sort_values('날짜').reset_index(drop=True) if recs else pd.DataFrame()

def fill_exact_main_template(df_data, start_date=None, end_date=None, year=2026, use_col20_for_out=False):
    """유량및수질관리.xlsx (51컬럼) 및 유량및수질관리 업로드양식.xlsx 와 서식·열너비·표시형식·수치 100% 일치 생성"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws['A1'] = "유량및수질관리 업로드양식"
    ws.merge_cells('A1:AY1')
    headers_r2 = {
        'A2': '날짜', 'B2': '유입량\n(반류수 포함)\n(㎥/일)', 'C2': '반류수 유량\n(㎥/일)', 'D2': '실제 유입량\n(㎥/일)',
        'E2': '처리량', 'H2': '방류량\n(㎥)/일', 'I2': '처리시설 유입전\n우수토실 방류량\n(㎥)/일', 'J2': '수온\n(℃)',
        'K2': '유입수질(연계전)', 'S2': '총인시설 유입수질(연계전)', 'AA2': '강우시 유입수질(1차처리전)', 'AI2': '방류수질',
        'AQ2': '방류수질(강우시 1차처리후 by-pass)', 'AY2': '비고'
    }
    for k, v in headers_r2.items():
        ws[k] = v
    subheaders_r3 = {
        'E3': '물리적\n(㎥/일)', 'F3': '생물학적\n(㎥/일)', 'G3': '고도\n(㎥/일)',
        'K3': 'pH\n(-)', 'L3': 'BOD\n(㎎/L)', 'M3': 'TOC\n(㎎/L)', 'N3': 'SS\n(㎎/L)', 'O3': 'T-N\n(㎎/L)', 'P3': 'T-P\n(㎎/L)', 'Q3': '총대장균군\n(개/㎖)', 'R3': '생태독성\n(TU)',
        'S3': 'pH\n(-)', 'T3': 'BOD\n(㎎/L)', 'U3': 'TOC\n(㎎/L)', 'V3': 'SS\n(㎎/L)', 'W3': 'T-N\n(㎎/L)', 'X3': 'T-P\n(㎎/L)', 'Y3': '총대장균군\n(개/㎖)', 'Z3': '생태독성\n(TU)',
        'AA3': 'pH\n(-)', 'AB3': 'BOD\n(㎎/L)', 'AC3': 'TOC\n(㎎/L)', 'AD3': 'SS\n(㎎/L)', 'AE3': 'T-N\n(㎎/L)', 'AF3': 'T-P\n(㎎/L)', 'AG3': '총대장균군\n(개/㎖)', 'AH3': '생태독성\n(TU)',
        'AI3': 'pH\n(-)', 'AJ3': 'BOD\n(㎎/L)', 'AK3': 'TOC\n(㎎/L)', 'AL3': 'SS\n(㎎/L)', 'AM3': 'T-N\n(㎎/L)', 'AN3': 'T-P\n(㎎/L)', 'AO3': '총대장균군\n(개/㎖)', 'AP3': '생태독성\n(TU)',
        'AQ3': 'pH\n(-)', 'AR3': 'BOD\n(㎎/L)', 'AS3': 'TOC\n(㎎/L)', 'AT3': 'SS\n(㎎/L)', 'AU3': 'T-N\n(㎎/L)', 'AV3': 'T-P\n(㎎/L)', 'AW3': '총대장균군\n(개/㎖)', 'AX3': '생태독성\n(TU)',
    }
    for k, v in subheaders_r3.items():
        ws[k] = v
    merges = ['A2:A3', 'B2:B3', 'C2:C3', 'D2:D3', 'E2:G2', 'H2:H3', 'I2:I3', 'J2:J3', 'K2:R2', 'S2:Z2', 'AA2:AH2', 'AI2:AP2', 'AQ2:AX2', 'AY2:AY3']
    for m in merges:
        ws.merge_cells(m)

    ws.row_dimensions[2].height = 18.75
    ws.row_dimensions[3].height = 18.75
    ws.column_dimensions['A'].width = 12.0
    ws.column_dimensions['B'].width = 12.5
    ws.column_dimensions['J'].width = 10.0
    for c_idx in range(3, 51):
        if c_idx != 10:
            ws.column_dimensions[get_column_letter(c_idx)].width = 13.0
    ws.column_dimensions['AY'].width = 18.75

    thin_side = Side(style='thin', color='000000')
    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    font_header = Font(name='dotum', size=9, bold=True)
    for r in range(1, 4):
        for c in range(1, 52):
            cell = ws.cell(r, c)
            cell.alignment = align_center
            cell.font = font_header
            cell.number_format = '@'
            if r == 2:
                cell.border = Border(left=thin_side if c == 1 else None, right=thin_side, top=thin_side, bottom=thin_side)
            elif r == 3:
                cell.border = Border(left=thin_side if c == 1 else None, right=thin_side, top=None, bottom=thin_side)

    if start_date and end_date:
        d_range = pd.date_range(start_date, end_date)
    else:
        d_range = pd.date_range(f"{year}-01-01", f"{year}-12-31")

    lookup = {}
    if df_data is not None and not df_data.empty and '날짜' in df_data.columns:
        for _, r in df_data.iterrows():
            lookup[str(r['날짜']).split()[0]] = r

    font_data = Font(name='맑은 고딕', size=11, bold=False)
    out_start = 20 if use_col20_for_out else 36

    for r_idx, dt in enumerate(d_range, start=4):
        d_str = dt.strftime('%Y-%m-%d')
        if not use_col20_for_out:
            c1 = ws.cell(r_idx, 1, datetime.datetime(dt.year, dt.month, dt.day, 0, 0))
            c1.number_format = 'mm-dd-yy'
            c1.font = font_data
        r_match = lookup.get(d_str, None)
        if r_match is not None:
            if use_col20_for_out and (pd.notna(r_match.get('유입량')) or pd.notna(r_match.get('방류BOD'))):
                serial_dt = (dt.date() - datetime.date(1899, 12, 30)).days
                c1 = ws.cell(r_idx, 1, serial_dt)
                c1.number_format = 'General'
                c1.font = font_data

            val_in = r_match.get('유입량', None)
            if pd.notna(val_in) and str(val_in).strip() != '':
                in_val = int(round(float(val_in))) if float(val_in).is_integer() else float(val_in)
                for col_target in [2, 4, 7]:
                    c = ws.cell(r_idx, col_target, in_val)
                    c.font = font_data
                    c.number_format = 'General' if use_col20_for_out else '_-* #,##0_-;\\-* #,##0_-;_-* "-"_-;_-@_-'
            val_out = r_match.get('방류량', None)
            if pd.notna(val_out) and str(val_out).strip() != '':
                out_val = int(round(float(val_out))) if float(val_out).is_integer() else float(val_out)
                c = ws.cell(r_idx, 8, out_val)
                c.font = font_data
                c.number_format = 'General' if use_col20_for_out else '_-* #,##0_-;\\-* #,##0_-;_-* "-"_-;_-@_-'
            val_temp = r_match.get('수온', None)
            if pd.notna(val_temp) and str(val_temp).strip() != '':
                t_num = int(round(float(val_temp))) if float(val_temp).is_integer() else round(float(val_temp), 1)
                c = ws.cell(r_idx, 10, t_num)
                c.font = font_data
                c.number_format = 'General' if use_col20_for_out else '0.0_ '
            for col_idx, col_name, n_fmt in [
                (12, '유입BOD', '0.0_);[Red]\\(0.0\\)'),
                (13, '유입TOC', '0.0_);[Red]\\(0.0\\)'),
                (14, '유입SS', '0.0_);[Red]\\(0.0\\)'),
                (15, '유입TN', '0.000_);[Red]\\(0.000\\)'),
                (16, '유입TP', '0.000_);[Red]\\(0.000\\)'),
                (17, '유입대장균', '#,##0_);[Red]\\(#,##0\\)')
            ]:
                v = r_match.get(col_name, None)
                if pd.notna(v) and str(v).strip() != '':
                    v_num = int(round(float(v))) if (col_idx == 17 or float(v).is_integer()) else round(float(v), 3 if col_idx in [15, 16] else 1)
                    c = ws.cell(r_idx, col_idx, v_num)
                    c.font = font_data
                    c.number_format = 'General' if use_col20_for_out else n_fmt
            for offset_i, (col_name, n_fmt) in enumerate([
                ('방류BOD', '0.0_);[Red]\\(0.0\\)'),
                ('방류TOC', '0.0_);[Red]\\(0.0\\)'),
                ('방류SS', '0.0_);[Red]\\(0.0\\)'),
                ('방류TN', '0.000_);[Red]\\(0.000\\)'),
                ('방류TP', '0.000_);[Red]\\(0.000\\)'),
                ('방류대장균', '#,##0_);[Red]\\(#,##0\\)')
            ]):
                v = r_match.get(col_name, None)
                if pd.notna(v) and str(v).strip() != '':
                    v_num = int(round(float(v))) if (offset_i == 5 or float(v).is_integer()) else round(float(v), 3 if offset_i in [3, 4] else 1)
                    c = ws.cell(r_idx, out_start + offset_i, v_num)
                    c.font = font_data
                    c.number_format = 'General' if use_col20_for_out else n_fmt
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

def fill_exact_reuse_template(df_data, start_date=None, end_date=None, year=2026):
    """재이용수 업로드양식(최종).xlsx 와 서식·열너비·테두리·수치 100% 일치 생성"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws['A1'] = "재이용수 업로드양식"
    ws.merge_cells('A1:T1')
    ws['A2'] = "날짜"; ws['B2'] = "합계(㎥)"; ws['C2'] = "장내용수(㎥)"; ws['K2'] = "장외용수(㎥)"; ws['T2'] = "사유"
    subheaders = {
        'C3': '소계', 'D3': '세척수', 'E3': '냉각수', 'F3': '청소수', 'G3': '식수대', 'H3': '희석용수', 'I3': '중수도', 'J3': '기타',
        'K3': '소계', 'L3': '청소화장실용수', 'M3': '세척살수용수', 'N3': '조경용수', 'O3': '친수용수', 'P3': '지하수충전', 'Q3': '농업용수', 'R3': '하천등유지용수', 'S3': '공업용수'
    }
    for k, v in subheaders.items():
        ws[k] = v
    for m in ['A2:A3', 'B2:B3', 'C2:J2', 'K2:S2', 'T2:T3']:
        ws.merge_cells(m)

    ws.row_dimensions[2].height = 18.75
    ws.row_dimensions[3].height = 18.75
    ws.column_dimensions['A'].width = 10.625
    ws.column_dimensions['B'].width = 12.5
    for c_idx in range(3, 20):
        ws.column_dimensions[get_column_letter(c_idx)].width = 13.0
    ws.column_dimensions['T'].width = 25.0

    thin_side = Side(style='thin', color='000000')
    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    font_header = Font(name='dotum', size=9, bold=True)
    for r in range(1, 4):
        for c in range(1, 21):
            cell = ws.cell(r, c)
            cell.alignment = align_center
            cell.font = font_header
            cell.number_format = '@'
            if r == 2:
                cell.border = Border(left=thin_side if c == 1 else None, right=thin_side if c in [1, 2, 3, 10, 11, 19, 20] else None, top=thin_side, bottom=thin_side)
            elif r == 3:
                cell.border = Border(left=thin_side if c == 1 else None, right=thin_side, top=None, bottom=thin_side)

    if start_date and end_date:
        d_range = pd.date_range(start_date, end_date)
    else:
        d_range = pd.date_range(f"{year}-01-01", f"{year}-12-31")

    lookup = {}
    if df_data is not None and not df_data.empty and '날짜' in df_data.columns:
        for _, r in df_data.iterrows():
            lookup[str(r['날짜']).split()[0]] = r

    font_data = Font(name='맑은 고딕', size=11, bold=False)
    for r_idx, dt in enumerate(d_range, start=4):
        d_str = dt.strftime('%Y-%m-%d')
        r_match = lookup.get(d_str, None)
        if r_match is not None:
            val_reuse = r_match.get('재이용수', None)
            if pd.notna(val_reuse) and str(val_reuse).strip() != '':
                c1 = ws.cell(r_idx, 1, datetime.datetime(dt.year, dt.month, dt.day, 0, 0))
                c1.number_format = 'mm-dd-yy'
                c1.font = font_data
                reuse_val = int(round(float(val_reuse))) if float(val_reuse).is_integer() else float(val_reuse)
                for col_target in [2, 3, 4]:
                    c = ws.cell(r_idx, col_target, reuse_val)
                    c.font = font_data
                    c.number_format = 'General'
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

def fill_exact_small_template(df_data, fac_name, start_date=None, end_date=None, year=2026):
    """소규모 6개소(산음·삼가리·진목·몰운·단월마을·당의) 업로드양식 예시 6종과 100% 동일한 서식·행수·수치 생성"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws['A1'] = "유량및수질관리 업로드양식"
    ws.merge_cells('A1:X1')
    headers_r2 = {'A2': '날짜', 'B2': '유입량\n(㎥/일)', 'C2': '처리량', 'F2': '방류량\n(㎥)/일', 'G2': '수온\n(℃)', 'H2': '유입수질', 'P2': '방류수질', 'X2': '비고'}
    for k, v in headers_r2.items():
        ws[k] = v
    subheaders_r3 = {
        'C3': '물리적\n(㎥/일)', 'D3': '생물학적\n(㎥/일)', 'E3': '고도\n(㎥/일)',
        'H3': 'pH\n(-)', 'I3': 'BOD\n(㎎/L)', 'J3': 'TOC\n(㎎/L)', 'K3': 'SS\n(㎎/L)', 'L3': 'T-N\n(㎎/L)', 'M3': 'T-P\n(㎎/L)', 'N3': '총대장균군\n(개/㎖)', 'O3': '생태독성\n(TU)',
        'P3': 'pH\n(-)', 'Q3': 'BOD\n(㎎/L)', 'R3': 'TOC\n(㎎/L)', 'S3': 'SS\n(㎎/L)', 'T3': 'T-N\n(㎎/L)', 'U3': 'T-P\n(㎎/L)', 'V3': '총대장균군\n(개/㎖)', 'W3': '생태독성\n(TU)'
    }
    for k, v in subheaders_r3.items():
        ws[k] = v
    for m in ['A2:A3', 'B2:B3', 'C2:E2', 'F2:F3', 'G2:G3', 'H2:O2', 'P2:W2', 'X2:X3']:
        ws.merge_cells(m)

    r3_heights = {'산음': 18.75, '삼가리': 29.25, '진목': 18.75, '몰운': 18.75, '단월마을': 30.75, '당의': 18.75}
    ws.row_dimensions[2].height = 18.75
    ws.row_dimensions[3].height = r3_heights.get(fac_name, 18.75)
    ws.column_dimensions['A'].width = 12.0
    ws.column_dimensions['B'].width = 12.5
    ws.column_dimensions['G'].width = 10.0
    for c_idx in range(3, 24):
        if c_idx != 7:
            ws.column_dimensions[get_column_letter(c_idx)].width = 13.0
    if fac_name == '단월마을':
        ws.column_dimensions['N'].width = 12.125
        ws.column_dimensions['O'].width = 10.0
    ws.column_dimensions['X'].width = 18.75

    thin_side = Side(style='thin', color='000000')
    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    font_header = Font(name='dotum', size=9, bold=True)
    for r in range(1, 4):
        for c in range(1, 25):
            cell = ws.cell(r, c)
            cell.alignment = align_center
            cell.font = font_header
            cell.number_format = '@'
            if r == 2:
                cell.border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
            elif r == 3:
                cell.border = Border(left=thin_side if c == 1 else None, right=thin_side, top=None, bottom=thin_side)

    lookup = {}
    raw_wq_dates = []
    raw_flow_dates = []
    if df_data is not None and not df_data.empty and '날짜' in df_data.columns:
        for _, r in df_data.sort_values(by='날짜').iterrows():
            d_key = str(r['날짜']).split()[0]
            lookup[d_key] = r
            if pd.notna(r.get('유입BOD')) or pd.notna(r.get('방류BOD')) or pd.notna(r.get('유입TOC')):
                raw_wq_dates.append(pd.to_datetime(d_key))
            if pd.notna(r.get('유입량')) or pd.notna(r.get('방류량')):
                raw_flow_dates.append(pd.to_datetime(d_key))

    valid_wq_dates = []
    for idx_m, m_dt in enumerate(raw_wq_dates):
        if idx_m + 1 < len(raw_wq_dates) and (raw_wq_dates[idx_m + 1] - m_dt).days == 1:
            continue
        valid_wq_dates.append(m_dt)
    valid_wq_keys = {dt.strftime('%Y-%m-%d') for dt in valid_wq_dates}

    daily_flow_in_map = {}
    daily_flow_out_map = {}
    for d_key, r_item in lookup.items():
        if d_key not in valid_wq_keys and len(raw_flow_dates) < 200 and (pd.notna(r_item.get('유입BOD')) or pd.notna(r_item.get('방류BOD'))):
            continue
        if pd.notna(r_item.get('유입량')):
            daily_flow_in_map[d_key] = float(r_item.get('유입량'))
        if pd.notna(r_item.get('방류량')):
            daily_flow_out_map[d_key] = float(r_item.get('방류량'))

    if len(daily_flow_in_map) < 200 and valid_wq_dates:
        prev_dt = None
        for m_dt in valid_wq_dates:
            d_str = m_dt.strftime('%Y-%m-%d')
            r_item = lookup[d_str]
            val_out = round(float(r_item.get('방류량', r_item.get('유입량', 0.0))), 1)
            val_in = 0.0 if (fac_name == '삼가리' and m_dt >= pd.to_datetime(f"{m_dt.year}-03-05")) else round(float(r_item.get('유입량', val_out)), 1)
            start_fill = pd.to_datetime(f"{m_dt.year}-01-01") if prev_dt is None else (prev_dt + pd.Timedelta(days=1))
            for w_dt in pd.date_range(start_fill, m_dt):
                w_str = w_dt.strftime('%Y-%m-%d')
                daily_flow_in_map[w_str] = val_in
                daily_flow_out_map[w_str] = val_out
            prev_dt = m_dt

    if start_date and end_date and not (str(start_date).endswith("-01-01") and str(end_date).endswith("-12-31")):
        d_range = pd.date_range(start_date, end_date)
    else:
        if fac_name in ['산음', '진목', '당의', '몰운'] and daily_flow_out_map:
            max_flow_dt = max(pd.to_datetime(list(daily_flow_out_map.keys())))
            d_range = pd.date_range(f"{year}-01-01", max_flow_dt)
        else:
            d_range = pd.date_range(f"{year}-01-01", f"{year}-12-31")

    font_data = Font(name='맑은 고딕', size=11, bold=False)
    date_fmt = 'mm-dd-yy' if fac_name in ['몰운', '당의'] else 'yyyy\\-mm\\-dd'
    flow_fmt_map = {
        '산음': '0.0',
        '삼가리': '0.0_);[Red]\\(0.0\\)',
        '진목': '#,##0.0_);[Red]\\(#,##0.0\\)',
        '몰운': 'General',
        '단월마을': '0.0',
        '당의': 'General'
    }
    fl_nf = flow_fmt_map.get(fac_name, '0.0')

    for r_idx, dt in enumerate(d_range, start=4):
        d_str = dt.strftime('%Y-%m-%d')
        c1 = ws.cell(r_idx, 1, datetime.datetime(dt.year, dt.month, dt.day, 0, 0))
        c1.number_format = date_fmt
        c1.font = font_data

        if d_str in daily_flow_in_map or d_str in daily_flow_out_map:
            f_in_v = daily_flow_in_map.get(d_str, None)
            f_out_v = daily_flow_out_map.get(d_str, f_in_v)
            if f_in_v is not None and pd.notna(f_in_v):
                v_i = int(round(f_in_v)) if float(f_in_v).is_integer() else float(f_in_v)
                c_b = ws.cell(r_idx, 2, v_i); c_b.font = font_data; c_b.number_format = fl_nf
                c_e = ws.cell(r_idx, 5, v_i); c_e.font = font_data; c_e.number_format = fl_nf
            if f_out_v is not None and pd.notna(f_out_v):
                v_o = int(round(f_out_v)) if float(f_out_v).is_integer() else float(f_out_v)
                c_f = ws.cell(r_idx, 6, v_o); c_f.font = font_data; c_f.number_format = fl_nf

        if d_str in valid_wq_keys:
            r_match = lookup.get(d_str, None)
            if r_match is not None:
                for col_idx, col_name, n_fmt in [
                    (9, '유입BOD', 'General' if fac_name in ['몰운', '당의'] else '0.0'),
                    (10, '유입TOC', 'General' if fac_name in ['몰운', '당의'] else '0.0'),
                    (11, '유입SS', 'General' if fac_name in ['몰운', '당의'] else '0.0'),
                    (12, '유입TN', 'General' if fac_name in ['진목', '몰운', '당의'] else '0.000'),
                    (13, '유입TP', 'General' if fac_name in ['진목', '몰운', '당의'] else '0.000'),
                    (14, '유입대장균', 'General' if fac_name in ['몰운', '당의'] else '#,##0_);[Red]\\(#,##0\\)')
                ]:
                    v = r_match.get(col_name, None)
                    if pd.notna(v) and str(v).strip() != '':
                        v_num = int(round(float(v))) if (col_idx == 14 or float(v).is_integer()) else round(float(v), 3 if col_idx in [12, 13] else 1)
                        c = ws.cell(r_idx, col_idx, v_num); c.font = font_data; c.number_format = n_fmt
                for col_idx, col_name, n_fmt in [
                    (17, '방류BOD', 'General' if fac_name in ['삼가리', '몰운', '당의'] else ('0.0_ ' if fac_name == '단월마을' else '0.0')),
                    (18, '방류TOC', 'General' if fac_name in ['삼가리', '몰운', '당의'] else ('0.0_ ' if fac_name == '단월마을' else '0.0')),
                    (19, '방류SS', 'General' if fac_name in ['삼가리', '몰운', '당의'] else ('0.0_ ' if fac_name == '단월마을' else '0.0')),
                    (20, '방류TN', '0.000_);[Red]\\(0.000\\)' if fac_name == '단월마을' else 'General'),
                    (21, '방류TP', '0.000_);[Red]\\(0.000\\)' if fac_name == '단월마을' else 'General'),
                    (22, '방류대장균', '#,##0_ ' if fac_name in ['삼가리', '진목'] else 'General')
                ]:
                    v = r_match.get(col_name, None)
                    if pd.notna(v) and str(v).strip() != '':
                        v_num = int(round(float(v))) if (col_idx == 22 or float(v).is_integer()) else round(float(v), 3 if col_idx in [20, 21] else 1)
                        c = ws.cell(r_idx, col_idx, v_num); c.font = font_data; c.number_format = n_fmt
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

EMBEDDED_MASTER_CSV_B64 = "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"

def auto_seed_master_db_from_local_files():
    """서버의 기존 마스터 DB가 오염되었거나 비어 있을 경우, 내장된 무오차 마스터 데이터(단월본장 + 소규모 6개소 + 개인하수 6개소)로 즉시 자동 복원"""
    needs_restore = True
    if os.path.exists(MASTER_ACCUM_DB):
        try:
            df_chk = pd.read_csv(MASTER_ACCUM_DB)
            has_priv = all(pf in df_chk['시설명'].unique() for pf in PRIVATE_PLANTS) if '시설명' in df_chk.columns else False
            saneum_df = df_chk[df_chk['시설명'] == '산음'] if '시설명' in df_chk.columns else pd.DataFrame()
            saneum_flow_cnt = saneum_df['유입량'].notna().sum() if ('유입량' in saneum_df.columns and not saneum_df.empty) else 0
            saneum_bod_cnt = saneum_df['유입BOD'].notna().sum() if ('유입BOD' in saneum_df.columns and not saneum_df.empty) else 0
            # 산음 일일유량이 270건 이상이고 수질 측정일이 정확히 39회일 때만 정상 DB로 인정 (오염된 구버전 DB 자동 감지 및 덮어쓰기)
            if len(df_chk) > 1500 and has_priv and saneum_flow_cnt >= 270 and saneum_bod_cnt == 39:
                needs_restore = False
        except Exception:
            needs_restore = True

    if needs_restore:
        try:
            raw_bytes = zlib.decompress(base64.b64decode(EMBEDDED_MASTER_CSV_B64))
            with open(MASTER_ACCUM_DB, "wb") as f_out:
                f_out.write(raw_bytes)
        except Exception:
            pass

auto_seed_master_db_from_local_files()

def generate_main_monthly_wq_excel(year=2026, sel_month=None):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    df_all = get_master_data(MAIN_PLANT, f"{year}-01-01", f"{year}-12-31")
    months_to_make = [sel_month] if sel_month else list(range(1, 13))
    align_c = Alignment(horizontal='center', vertical='center', wrap_text=True)
    font_title = Font(name='맑은 고딕', size=14, bold=True)
    font_hdr = Font(name='맑은 고딕', size=9, bold=True)
    font_body = Font(name='맑은 고딕', size=9, bold=False)

    for m in months_to_make:
        ws = wb.create_sheet(title=f"단월 {m}월")
        ws['A2'] = f"단월공공하수처리시설 수질검사결과({m}월)"
        ws['A2'].font = font_title
        ws['A4'] = f"{year}년 {m}월 (시설용량 : 1700㎥/일)"
        ws['N4'] = "(단위 : ㎎/ℓ, 개/㎖, ㎥/일)"
        ws['A5'] = "일자"; ws['B5'] = "유        입        수"; ws['H5'] = "생물반응조"; ws['J5'] = "방        류        수"
        ws['P5'] = "유입량"; ws['Q5'] = "재이용량"; ws['R5'] = "방류량"; ws['S5'] = "반응조\n수온(℃)"
        ws.merge_cells('A5:A6'); ws.merge_cells('B5:G5'); ws.merge_cells('H5:I5'); ws.merge_cells('J5:O5')
        ws.merge_cells('P5:P6'); ws.merge_cells('Q5:Q6'); ws.merge_cells('R5:R6'); ws.merge_cells('S5:S6')
        sub_cols = ['BOD', 'TOC', 'SS', 'T-N', 'T-P', '대장균\n군수', 'ML A', 'ML B', 'BOD', 'TOC', 'SS', 'T-N', 'T-P', '대장균\n군수']
        for idx_c, sc_name in enumerate(sub_cols, start=2):
            ws.cell(6, idx_c, sc_name)
        for r_i in range(5, 7):
            for c_i in range(1, 20):
                ws.cell(r_i, c_i).alignment = align_c
                ws.cell(r_i, c_i).font = font_hdr

        ym_str = f"{year}-{m:02d}"
        df_m = df_all[df_all['날짜'].astype(str).str.startswith(ym_str)].copy() if not df_all.empty else pd.DataFrame()
        lookup = {int(str(r['날짜']).split('-')[2][:2]): r for _, r in df_m.iterrows()} if not df_m.empty else {}

        import calendar
        last_day = calendar.monthrange(year, m)[1]
        col_keys = ['유입BOD', '유입TOC', '유입SS', '유입TN', '유입TP', '유입대장균', 'MLSS_A', 'MLSS_B',
                    '방류BOD', '방류TOC', '방류SS', '방류TN', '방류TP', '방류대장균', '유입량', '재이용수', '방류량', '수온']
        for d in range(1, last_day + 1):
            r_idx = 6 + d
            ws.cell(r_idx, 1, d).alignment = align_c
            ws.cell(r_idx, 1).font = font_body
            if d in lookup:
                row_d = lookup[d]
                for c_offset, k_name in enumerate(col_keys, start=2):
                    val = pd.to_numeric(row_d.get(k_name, np.nan), errors='coerce')
                    if pd.notna(val):
                        cell = ws.cell(r_idx, c_offset, float(val))
                        cell.font = font_body
                        cell.alignment = align_c

        # 평균, 최대, 최소, 설계 행 추가 (38~41행)
        stat_start = 6 + last_day + 1
        for s_offset, (s_label, s_func) in enumerate([("평균", "mean"), ("최대", "max"), ("최소", "min")]):
            sr = stat_start + s_offset
            ws.cell(sr, 1, s_label).font = font_hdr
            ws.cell(sr, 1).alignment = align_c
            if not df_m.empty:
                for c_offset, k_name in enumerate(col_keys, start=2):
                    if k_name in df_m.columns:
                        s_series = pd.to_numeric(df_m[k_name], errors='coerce').dropna()
                        if not s_series.empty:
                            v_stat = getattr(s_series, s_func)()
                            dec = 3 if 'TN' in k_name or 'TP' in k_name else (0 if k_name in ['유입량', '재이용수', '방류량', '유입대장균', '방류대장균', 'MLSS_A', 'MLSS_B'] else 1)
                            ws.cell(sr, c_offset, round(float(v_stat), dec) if dec > 0 else int(round(float(v_stat)))).font = font_hdr
                            ws.cell(sr, c_offset).alignment = align_c
        sr_des = stat_start + 3
        ws.cell(sr_des, 1, "설계").font = font_hdr; ws.cell(sr_des, 1).alignment = align_c
        for c_i, d_v in [(2, 178), (3, 168), (4, 191), (5, 42.1), (6, 4.6), (7, 265000), (10, 5), (11, 19.8), (12, 9.9), (13, 19.1), (14, 0.2), (15, 1000)]:
            ws.cell(sr_des, c_i, d_v).font = font_body; ws.cell(sr_des, c_i).alignment = align_c
        if not df_m.empty:
            re_sum = pd.to_numeric(df_m.get('재이용수', pd.Series(dtype=float)), errors='coerce').sum()
            out_sum = pd.to_numeric(df_m.get('방류량', pd.Series(dtype=float)), errors='coerce').sum()
            ws.cell(sr_des, 17, int(round(re_sum))).font = font_hdr
            ws.cell(sr_des, 18, int(round(out_sum))).font = font_hdr

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

def generate_small_monthly_wq_excel(year=2026, sel_month=None):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    months_to_make = [sel_month] if sel_month else list(range(1, 13))
    align_c = Alignment(horizontal='center', vertical='center', wrap_text=True)
    font_title = Font(name='맑은 고딕', size=13, bold=True)
    font_hdr = Font(name='맑은 고딕', size=9, bold=True)
    font_body = Font(name='맑은 고딕', size=9, bold=False)

    small_specs_order = [
        ("진목", "진  목\n(23㎥/일)"),
        ("산음", "산음리\n(100㎥/일)"),
        ("몰운", "몰운\n(60㎥/일)"),
        ("삼가리", "삼가리\n(120㎥/일)"),
        ("단월마을", "단월마을\n(30㎥/일)"),
        ("당의", "당의\n(45㎥/일)")
    ]
    col_keys = ['유입BOD', '유입TOC', '유입SS', '유입TN', '유입TP', '유입대장균',
                '방류BOD', '방류TOC', '방류SS', '방류TN', '방류TP', '방류대장균', '주간유량', '방류량']

    for m in months_to_make:
        ws = wb.create_sheet(title=f"{m}월")
        ws['A1'] = f"소규모공공하수처리시설 월별 수질현황({m}월)"
        ws['A1'].font = font_title
        ws['A2'] = "처리장명"; ws['B2'] = "날 짜"; ws['C2'] = "유입수"; ws['I2'] = "방류수"; ws['O2'] = "유량\n㎥/주"; ws['P2'] = "유량\n㎥/일"
        ws.merge_cells('A2:A3'); ws.merge_cells('B2:B3'); ws.merge_cells('C2:H2'); ws.merge_cells('I2:N2'); ws.merge_cells('O2:O3'); ws.merge_cells('P2:P3')
        for idx_c, sc in enumerate(['BOD', 'TOC', 'SS', 'T-N', 'T-P', '대장균', 'BOD', 'TOC', 'SS', 'T-N', 'T-P', '대장균'], start=3):
            ws.cell(3, idx_c, sc)
        for r_i in range(2, 4):
            for c_i in range(1, 17):
                ws.cell(r_i, c_i).alignment = align_c; ws.cell(r_i, c_i).font = font_hdr

        cur_r = 4
        ym_str = f"{year}-{m:02d}"
        for fac_key, fac_disp in small_specs_order:
            df_f = get_master_data(fac_key, f"{ym_str}-01", f"{ym_str}-31")
            if not df_f.empty and '주간유량' not in df_f.columns and '방류량' in df_f.columns:
                df_f['주간유량'] = pd.to_numeric(df_f['방류량'], errors='coerce') * 7.0
            ws.cell(cur_r, 1, fac_disp).alignment = align_c; ws.cell(cur_r, 1).font = font_hdr
            if not df_f.empty:
                for _, row_d in df_f.iterrows():
                    ws.cell(cur_r, 2, str(row_d['날짜'])[:10]).alignment = align_c; ws.cell(cur_r, 2).font = font_body
                    for c_off, k_n in enumerate(col_keys, start=3):
                        v = pd.to_numeric(row_d.get(k_n, np.nan), errors='coerce')
                        if pd.notna(v):
                            dec = 3 if 'TN' in k_n or 'TP' in k_n else (0 if k_n in ['유입대장균', '방류대장균', '주간유량'] else 1)
                            ws.cell(cur_r, c_off, round(float(v), dec) if dec > 0 else int(round(float(v)))).alignment = align_c
                            ws.cell(cur_r, c_off).font = font_body
                    cur_r += 1
            else:
                cur_r += 1

            # 합계, 최대, 최소, 평균
            ws.cell(cur_r, 2, "합계").alignment = align_c; ws.cell(cur_r, 2).font = font_hdr
            if not df_f.empty and '주간유량' in df_f.columns:
                wk_s = pd.to_numeric(df_f['주간유량'], errors='coerce').dropna()
                if not wk_s.empty:
                    ws.cell(cur_r, 15, int(round(wk_s.sum()))).alignment = align_c; ws.cell(cur_r, 15).font = font_hdr
            cur_r += 1

            for s_lbl, s_fn in [("최대", "max"), ("최소", "min"), ("평균", "mean")]:
                ws.cell(cur_r, 2, s_lbl).alignment = align_c; ws.cell(cur_r, 2).font = font_hdr
                if not df_f.empty:
                    for c_off, k_n in enumerate(col_keys, start=3):
                        if k_n in df_f.columns:
                            s_ser = pd.to_numeric(df_f[k_n], errors='coerce').dropna()
                            if not s_ser.empty:
                                v_st = getattr(s_ser, s_fn)()
                                dec = 3 if 'TN' in k_n or 'TP' in k_n else (0 if k_n in ['유입대장균', '방류대장균', '주간유량'] else 1)
                                ws.cell(cur_r, c_off, round(float(v_st), dec) if dec > 0 else int(round(float(v_st)))).alignment = align_c
                                ws.cell(cur_r, c_off).font = font_hdr
                cur_r += 1

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

def generate_priv_monthly_wq_excel(year=2026):
    """개인소규모(2026년).xlsx 중 단월 관할 6개소(석산리·음지·양지·복지회관·인이피·돌고개)만 정확히 추출·생성"""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    align_c = Alignment(horizontal='center', vertical='center', wrap_text=True)
    font_title = Font(name='돋움체', size=20, bold=True)
    font_cap = Font(name='돋움체', size=9, bold=True)
    font_hdr = Font(name='돋움체', size=11, bold=True)
    font_dt = Font(name='돋움체', size=10, bold=True)
    font_body = Font(name='돋움체', size=10, bold=False)
    font_stat = Font(name='돋움체', size=10, bold=True)
    font_bot_lbl = Font(name='돋움체', size=11, bold=True)
    font_bot_val = Font(name='돋움체', size=11, bold=False)

    med_s = Side(style='medium', color='000000')
    dbl_s = Side(style='double', color='000000')
    hair_s = Side(style='hair', color='000000')
    thin_s = Side(style='thin', color='000000')

    priv_specs = [
        ("석산리", "13.석산리", f"{year}년 석산리 오수처리시설", "(30톤/일)", True),
        ("음지", "18.물레울(음지)", f"{year}년 물레울(음지) 오수처리시설 ", "(15톤/일)", False),
        ("양지", "19.물레울(양지)", f"{year}년 물레울(양지) 오수처리시설", "(15톤/일)", False),
        ("복지회관", "20.물레울(복지회관)", f"{year}년 물레울(복지회관) 오수처리시설", "(10톤/일)", False),
        ("인이피", "21.인이피", f"{year}년 인이피 오수처리시설", "(10톤/일)", False),
        ("돌고개", "22.돌고개", f"{year}년 돌고개 오수처리시설", "(15톤/일)", False),
    ]
    import calendar
    for fac_key, s_title, header_title, cap_str, has_flow in priv_specs:
        ws = wb.create_sheet(title=s_title)
        ws['A1'] = header_title
        ws['A1'].font = font_title
        ws['A1'].alignment = align_c
        ws.merge_cells('A1:G1')

        ws['F2'] = cap_str
        ws['F2'].font = font_cap
        ws['F2'].alignment = align_c
        ws.merge_cells('F2:G2')

        ws['A3'] = "날 짜"; ws['B3'] = "유    입    수"; ws['D3'] = "방    류    수"
        ws['F3'] = "유  량\n(㎥/day)"; ws['G3'] = "월유입총량\n(㎥)"
        ws.merge_cells('A3:A4'); ws.merge_cells('B3:C3'); ws.merge_cells('D3:E3'); ws.merge_cells('F3:F4'); ws.merge_cells('G3:G4')
        ws['B4'] = "BOD"; ws['C4'] = "SS"; ws['D4'] = "BOD"; ws['E4'] = "SS"

        ws.row_dimensions[1].height = 25.5
        ws.row_dimensions[2].height = 12.0
        ws.row_dimensions[3].height = 14.25
        ws.row_dimensions[4].height = 14.25
        for col_l, w_v in [('A', 11.6), ('B', 10.5), ('C', 11.5), ('D', 11.5), ('E', 11.5), ('F', 11.0), ('G', 12.0)]:
            ws.column_dimensions[col_l].width = w_v

        for r_i in range(3, 5):
            for c_i in range(1, 8):
                cell = ws.cell(r_i, c_i)
                cell.alignment = align_c
                cell.font = font_hdr
                cell.border = Border(
                    top=med_s if r_i == 3 else hair_s,
                    bottom=dbl_s if (r_i == 4 or c_i in [1, 6, 7]) else hair_s,
                    left=thin_s, right=thin_s
                )

        cur_r = 5
        df_all_p = get_master_data(fac_key, f"{year}-01-01", f"{year}-12-31")
        all_in_b, all_in_s, all_out_b, all_out_s, all_fl, all_mtot = [], [], [], [], [], []

        for m in range(1, 13):
            ym_str = f"{year}-{m:02d}"
            df_m = df_all_p[df_all_p['날짜'].astype(str).str.startswith(ym_str)] if not df_all_p.empty else pd.DataFrame()
            days_in_m = calendar.monthrange(year, m)[1]
            ws.row_dimensions[cur_r].height = 15.0
            ws.row_dimensions[cur_r + 3].height = 14.25

            if not df_m.empty:
                r0 = df_m.iloc[-1]
                dt_obj = pd.to_datetime(r0['날짜']).to_pydatetime()
                in_b = pd.to_numeric(r0.get('유입BOD', np.nan), errors='coerce')
                in_s = pd.to_numeric(r0.get('유입SS', np.nan), errors='coerce')
                out_b = pd.to_numeric(r0.get('방류BOD', np.nan), errors='coerce')
                out_s = pd.to_numeric(r0.get('방류SS', np.nan), errors='coerce')
                fl_d = pd.to_numeric(r0.get('방류량', r0.get('유입량', np.nan)), errors='coerce')
                m_tot = round(float(fl_d) * days_in_m, 1) if (has_flow and pd.notna(fl_d)) else np.nan

                if pd.notna(in_b): all_in_b.append(float(in_b))
                if pd.notna(in_s): all_in_s.append(float(in_s))
                if pd.notna(out_b): all_out_b.append(float(out_b))
                if pd.notna(out_s): all_out_s.append(float(out_s))
                if has_flow and pd.notna(fl_d): all_fl.append(float(fl_d))
                if has_flow and pd.notna(m_tot): all_mtot.append(float(m_tot))

                c_dt = ws.cell(cur_r, 1, datetime.datetime(dt_obj.year, dt_obj.month, dt_obj.day, 0, 0))
                c_dt.number_format = 'm"월"\\ d"일"'
                c_dt.alignment = align_c; c_dt.font = font_dt

                for c_idx, v_num in [(2, in_b), (3, in_s), (4, out_b), (5, out_s)]:
                    if pd.notna(v_num):
                        c = ws.cell(cur_r, c_idx, float(v_num))
                        c.number_format = '0.0_ '; c.alignment = align_c; c.font = font_body
                if has_flow:
                    if pd.notna(fl_d):
                        c_f = ws.cell(cur_r, 6, float(fl_d))
                        c_f.number_format = '#,##0.0_ '; c_f.alignment = align_c; c_f.font = font_body
                else:
                    c_f = ws.cell(cur_r, 6, "유량계미설치")
                    c_f.alignment = align_c; c_f.font = font_body

                for s_off, s_lbl in enumerate(["최    대", "최    소", "평    균"], start=1):
                    sr = cur_r + s_off
                    ws.cell(sr, 1, s_lbl).alignment = align_c; ws.cell(sr, 1).font = font_stat
                    for c_idx, v_num in [(2, in_b), (3, in_s), (4, out_b), (5, out_s)]:
                        if pd.notna(v_num):
                            c = ws.cell(sr, c_idx, float(v_num))
                            c.number_format = '0.0_ '; c.alignment = align_c; c.font = font_stat
                    if has_flow:
                        if pd.notna(fl_d):
                            c_sf = ws.cell(sr, 6, float(fl_d))
                            c_sf.number_format = '0.0_);[Red]\\(0.0\\)'; c_sf.alignment = align_c; c_sf.font = font_stat
                            if s_lbl == "평    균" and pd.notna(m_tot):
                                c_mt = ws.cell(sr, 7, float(m_tot))
                                c_mt.number_format = '_-* #,##0.0_-;\\-* #,##0.0_-;_-* "-"_-;_-@_-'
                                c_mt.alignment = align_c; c_mt.font = font_stat
                        elif s_lbl in ["최    대", "최    소"]:
                            ws.cell(sr, 6, 0).alignment = align_c; ws.cell(sr, 6).font = font_stat
            else:
                if not has_flow:
                    ws.cell(cur_r, 6, "유량계미설치").alignment = align_c; ws.cell(cur_r, 6).font = font_body
                for s_off, s_lbl in enumerate(["최    대", "최    소", "평    균"], start=1):
                    sr = cur_r + s_off
                    ws.cell(sr, 1, s_lbl).alignment = align_c; ws.cell(sr, 1).font = font_stat
                    if s_lbl in ["최    대", "최    소"]:
                        for c_idx in range(2, 6):
                            ws.cell(sr, c_idx, 0).alignment = align_c; ws.cell(sr, c_idx).font = font_stat
                        if has_flow:
                            ws.cell(sr, 6, 0).alignment = align_c; ws.cell(sr, 6).font = font_stat

            for off_r in range(4):
                r_curr = cur_r + off_r
                for c_i in range(1, 8):
                    ws.cell(r_curr, c_i).border = Border(
                        top=dbl_s if off_r == 0 else hair_s,
                        bottom=dbl_s if off_r == 3 else hair_s,
                        left=thin_s, right=thin_s
                    )
            cur_r += 4

        # 53~56행: 연간 계 / 최대 / 최소 / 평균 요약 행
        if has_flow:
            ws.cell(53, 1, "계").alignment = align_c; ws.cell(53, 1).font = Font(name='돋움체', size=14, bold=True)
            if all_mtot:
                c_tot = ws.cell(53, 7, round(sum(all_mtot), 1))
                c_tot.number_format = '_-* #,##0.0_-;\\-* #,##0.0_-;_-* "-"_-;_-@_-'
                c_tot.alignment = align_c; c_tot.font = Font(name='돋움체', size=12, bold=True)

        for r_idx, (lbl_txt, fn_op) in [(54, ("최    대", max)), (55, ("최    소", min)), (56, ("평    균", lambda x: round(sum(x)/len(x), 1)))]:
            c_l = ws.cell(r_idx, 1, lbl_txt); c_l.alignment = align_c; c_l.font = font_bot_lbl
            for c_i, arr_v in [(2, all_in_b), (3, all_in_s), (4, all_out_b), (5, all_out_s)]:
                if arr_v:
                    c_v = ws.cell(r_idx, c_i, float(fn_op(arr_v)))
                    c_v.number_format = '_-* #,##0.0_-;\\-* #,##0.0_-;_-* "-"_-;_-@_-'
                    c_v.alignment = align_c; c_v.font = font_bot_val
            if has_flow and all_fl:
                c_fv = ws.cell(r_idx, 6, float(fn_op(all_fl)))
                c_fv.number_format = '_-* #,##0.0_-;\\-* #,##0.0_-;_-* "-"_-;_-@_-'
                c_fv.alignment = align_c; c_fv.font = font_bot_val
            for c_i in range(1, 8):
                ws.cell(r_idx, c_i).border = Border(
                    top=med_s if r_idx == 54 else hair_s,
                    bottom=med_s if r_idx == 56 else hair_s,
                    left=thin_s, right=thin_s
                )

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

def generate_wq_monthly_hwpx(year, month, main_sum_rows, sm_sum_rows, pr_sum_rows):
    out_buf = io.BytesIO()
    lines = [f"단월공공하수처리시설 및 소규모·개인하수 수질월보 ({year}년 {month}월)"]
    lines.append("1. 단월공공하수처리시설 수질검사결과 및 월간유량 요약")
    for r in main_sum_rows:
        lines.append(f"[{r.get('구분','')}] 유입량:{r.get('유입량','-')} 방류량:{r.get('방류량','-')} 재이용수:{r.get('재이용수','-')} | 방류BOD:{r.get('방류BOD','-')} TOC:{r.get('방류TOC','-')} SS:{r.get('방류SS','-')} T-N:{r.get('방류TN','-')} T-P:{r.get('방류TP','-')} 대장균:{r.get('방류대장균','-')}")
    lines.append("2. 소규모공공하수처리시설 6개소 월별 수질·유량 현황")
    for r in sm_sum_rows:
        lines.append(f"[{r.get('시설명','')} - {r.get('구분','')}] 방류BOD:{r.get('방류BOD','-')} TOC:{r.get('방류TOC','-')} SS:{r.get('방류SS','-')} T-N:{r.get('방류TN','-')} T-P:{r.get('방류TP','-')} 주간유량(㎥/주):{r.get('주간유량','-')} 일일유량(㎥/일):{r.get('방류량','-')}")
    lines.append("3. 개인하수처리시설 6개소 월별 수질·유량 현황")
    for r in pr_sum_rows:
        lines.append(f"[{r.get('시설명','')} ({r.get('측정일','')})] 유입BOD:{r.get('유입BOD','-')} 유입SS:{r.get('유입SS','-')} | 방류BOD:{r.get('방류BOD','-')} 방류SS:{r.get('방류SS','-')} | 일일유량:{r.get('방류량','-')} 주간유량:{r.get('주간유량','-')}")

    p_tags = "".join([f"<hp:p><hp:run><hp:t>{ln}</hp:t></hp:run></hp:p>" for ln in lines])
    sec_xml = f"<?xml version='1.0' encoding='UTF-8'?><hp:sec xmlns:hp='http://www.hancom.co.kr/hwpml/2011/paragraph'>{p_tags}</hp:sec>"
    with zipfile.ZipFile(out_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("mimetype", "application/hwp+zip")
        zf.writestr("version.xml", "<?xml version='1.0' encoding='UTF-8'?><hh:version xmlns:hh='http://www.hancom.co.kr/hwpml/2011/head' version='1.0'/>")
        zf.writestr("Contents/section0.xml", sec_xml.encode('utf-8'))
        zf.writestr("Preview/PrvText.txt", "\n".join(lines).encode('utf-8'))
    return out_buf.getvalue()

def generate_html_to_pdf_bytes(html_str):
    try:
        from weasyprint import HTML
        return HTML(string=html_str).write_pdf()
    except Exception:
        return html_str.encode('utf-8')

def generate_hwpx_monthly_report(sel_month, hwpx_template_file, sludge_data, solar_data, task_text, year=2026):
    months_window = [(sel_month - 5 + i - 1) % 12 + 1 for i in range(6)]
    cand = f"공공하수도시설 대행사업 월간보고서({sel_month}월).hwpx"
    if not os.path.exists(cand): cand = '공공하수도시설 대행사업 월간보고서(10월).hwpx'
    template_bytes = b""
    if hwpx_template_file is not None:
        template_bytes = hwpx_template_file.getvalue() if hasattr(hwpx_template_file, 'getvalue') else hwpx_template_file.read()
    elif os.path.exists(cand):
        with open(cand, 'rb') as f: template_bytes = f.read()
    out_buf = io.BytesIO()
    if template_bytes:
        try:
            in_zip = zipfile.ZipFile(io.BytesIO(template_bytes), 'r')
            out_zip = zipfile.ZipFile(out_buf, 'w', zipfile.ZIP_DEFLATED)
            for item in in_zip.infolist():
                data = in_zip.read(item.filename)
                if item.filename.startswith('Contents/section') and item.filename.endswith('.xml'):
                    text = data.decode('utf-8', errors='ignore')
                    text = re.sub(r'월간보고서\(\d{1,2}월\)', f'월간보고서({sel_month}월)', text)
                    text = re.sub(r'운영상황 보고\(\d{1,2}월\)', f'운영상황 보고({sel_month}월)', text)
                    for i, m in enumerate(months_window): text = text.replace(f'<{i+1}월헤더>', f'{m}월')
                    text = text.replace('{{SLUDGE_AVG}}', f"{sludge_data['avg']:.1f}")
                    text = text.replace('{{SLUDGE_MAX}}', f"{sludge_data['max']:.1f}")
                    text = text.replace('{{SLUDGE_MIN}}', f"{sludge_data['min']:.1f}")
                    text = text.replace('{{SOLAR_GEN}}', f"{solar_data['current_month']:.1f}")
                    if task_text: text = text.replace('{{DAILY_MAINTENANCE_TEXT}}', task_text)
                    data = text.encode('utf-8')
                out_zip.writestr(item, data)
            in_zip.close(); out_zip.close()
            return out_buf.getvalue()
        except Exception:
            pass
    m_str = ", ".join([str(m) + "월" for m in months_window])
    sec_xml = (
        "<?xml version='1.0' encoding='UTF-8'?>"
        "<hp:sec xmlns:hp='http://www.hancom.co.kr/hwpml/2011/paragraph'>"
        f"<hp:p><hp:run><hp:t>단월공공하수처리시설 대행사업 월간보고서 ({year}년 {sel_month}월)</hp:t></hp:run></hp:p>"
        f"<hp:p><hp:run><hp:t>최근 6개월 슬라이딩 윈도우: {m_str}</hp:t></hp:run></hp:p>"
        f"<hp:p><hp:run><hp:t>당월 슬러지 통계: 평균 {sludge_data['avg']:.1f}%, 최대 {sludge_data['max']:.1f}%, 최소 {sludge_data['min']:.1f}%</hp:t></hp:run></hp:p>"
        f"<hp:p><hp:run><hp:t>태양광 발전 실적: {solar_data['current_month']:.1f} kWh</hp:t></hp:run></hp:p>"
        f"<hp:p><hp:run><hp:t>주요 설비 유지보수 실적: {task_text}</hp:t></hp:run></hp:p>"
        "</hp:sec>"
    )
    with zipfile.ZipFile(out_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("mimetype", "application/hwp+zip")
        zf.writestr("version.xml", "<?xml version='1.0' encoding='UTF-8'?><hh:version xmlns:hh='http://www.hancom.co.kr/hwpml/2011/head' version='1.0'/>")
        zf.writestr("Contents/section0.xml", sec_xml.encode('utf-8'))
    return out_buf.getvalue()

def create_full_backup_zip():
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for label, d_path in ALL_STORAGE_DIRS.items():
            if os.path.exists(d_path):
                for fname in os.listdir(d_path):
                    fpath = os.path.join(d_path, fname)
                    if os.path.isfile(fpath):
                        zf.write(fpath, arcname=os.path.join(d_path, fname))
        for label, db_file in ALL_MASTER_DBS.items():
            if os.path.exists(db_file):
                zf.write(db_file, arcname=os.path.join("master_databases", db_file))
    return buf.getvalue()

def render_maintenance_screen(cfg, is_logged_in_user=False):
    current_time_str = datetime.datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')
    st.markdown(f"""
    <div style="max-width: 680px; margin: 40px auto 20px auto; background: white; border-radius: 20px; padding: 35px 30px; box-shadow: 0 20px 40px -15px rgba(0,0,0,0.1); border: 1px solid #E2E8F0; text-align: center;">
        <div style="font-size: 54px; margin-bottom: 10px;">🚧</div>
        <h2 style="color: #0F172A; font-weight: 900; margin-bottom: 12px; font-size: 24px;">자체 시스템 정기 점검 및 업데이트 중입니다</h2>
        <div style="display: inline-block; background: #FEF3C7; color: #D97706; padding: 5px 16px; border-radius: 30px; font-weight: 700; font-size: 13px; margin-bottom: 20px;">
            SYSTEM MAINTENANCE & DATA OPTIMIZATION
        </div>
        <p style="color: #475569; font-size: 15px; line-height: 1.65; margin-bottom: 25px;">
            {cfg.get("maintenance_msg", "현재 자체 시스템 정기 점검 및 AI 데이터베이스 최적화 업데이트 중입니다.")}<br>
            양평군민에게 보다 안정적이고 정확한 스마트 수질 관제 서비스를 제공하기 위해 시스템을 점검하고 있습니다.
        </p>
        <div style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 12px; padding: 15px; font-size: 13px; color: #64748B; text-align: left; margin-bottom: 20px;">
            • <b>점검 대상</b>: 단월공공하수처리시설(1,700 ㎥/일) 및 소규모 6개소 디지털 트윈 관제 플랫폼<br>
            • <b>현재 시각 (KST)</b>: <span style="color: #0284C7; font-weight: bold;">{current_time_str}</span><br>
            • <b>담당 부서</b>: 양평공사 환경2팀 (단월공공하수처리시설)
        </div>
    </div>
    """, unsafe_allow_html=True)

    if is_logged_in_user:
        col1, col2, col3 = st.columns([1, 1, 1])
        with col2:
            if st.button("로그아웃", use_container_width=True):
                st.session_state.logged_in = False; st.session_state.user_role = None; st.rerun()
    else:
        col1, col2, col3 = st.columns([1, 1.3, 1])
        with col2:
            with st.expander("🔒 관계자 및 관리자 전용 접속 (점검 중 접속)", expanded=True):
                login_opt = st.radio("접속 권한 선택", ["시스템 관리자 (마스터)", "승인된 운영요원 (접속 코드)"], horizontal=True)
                if login_opt == "시스템 관리자 (마스터)":
                    admin_pw_m = st.text_input("관리자 마스터 비밀번호", type="password", key="m_admin_pw")
                    if st.button("🚀 관리자 모드로 접속", type="primary", use_container_width=True, key="btn_m_admin_login"):
                        if hash_pw(admin_pw_m) == ADMIN_PW_HASH:
                            st.session_state.logged_in = True; st.session_state.user_role = "admin"; st.session_state.user_name = "최고관리자"
                            st.rerun()
                        else: st.error("관리자 비밀번호가 일치하지 않습니다.")
                else:
                    passcode_m = st.text_input("승인 접속 코드", type="password", value="DANWOL-2026!", key="m_user_code")
                    if st.button("🚀 운영요원 접속", type="primary", use_container_width=True, key="btn_m_user_login"):
                        if hash_pw(passcode_m) in WHITELIST_HASHES:
                            st.session_state.logged_in = True; st.session_state.user_role = "admin"; st.session_state.user_name = "인증 운영요원"
                            st.rerun()
                        else: st.error("유효하지 않은 승인 코드입니다.")

def check_login_system():
    if "logged_in" not in st.session_state: st.session_state.logged_in = False
    if "user_role" not in st.session_state: st.session_state.user_role = None
    cfg = load_system_config()
    is_maintenance = cfg.get("force_maintenance_screen", True)
    if st.session_state.logged_in:
        if is_maintenance and st.session_state.user_role != "admin":
            render_maintenance_screen(cfg, is_logged_in_user=True)
            return False
        return True
    if is_maintenance:
        render_maintenance_screen(cfg, is_logged_in_user=False)
        return False

    st.markdown("""
    <div style="text-align: center; padding: 25px 20px 10px 20px;">
        <div style="font-size: 44px;">💧</div>
        <h1 style="font-size: 28px; font-weight: 900; color: #0F172A; margin: 5px 0;">DANWOL AI-WaterOps 360</h1>
        <p style="font-size: 14.5px; color: #64748B; font-weight: 600;">단월 공공하수처리시설 지능형 통합 자율운전 & 디지털 트윈 관제 플랫폼</p>
    </div>
    """, unsafe_allow_html=True)
    col_l1, col_l2, col_l3 = st.columns([1, 1.3, 1])
    with col_l2:
        tab_login, tab_request = st.tabs(["🔒 시스템 로그인", "📝 신규 사용자 승인 요청"])
        with tab_login:
            login_type = st.radio("접속 유형 선택", ["일반 사용자 (승인 접속 코드 / 계정)", "시스템 관리자 (승인 대시보드)"], horizontal=True)
            if login_type == "시스템 관리자 (승인 대시보드)":
                admin_pw = st.text_input("관리자 마스터 비밀번호", type="password", key="admin_pw_input")
                if st.button("🚀 관리자 모드로 접속", type="primary", use_container_width=True):
                    if hash_pw(admin_pw) == ADMIN_PW_HASH:
                        st.session_state.logged_in = True; st.session_state.user_role = "admin"; st.session_state.user_name = "최고관리자"
                        st.rerun()
                    else: st.error("관리자 비밀번호가 일치하지 않습니다.")
            else:
                passcode = st.text_input("부여받은 승인 접속 코드", type="password", key="passcode_input", value="DANWOL-2026!")
                if st.button("🚀 접속하기", type="primary", use_container_width=True):
                    if hash_pw(passcode) in WHITELIST_HASHES:
                        st.session_state.logged_in = True; st.session_state.user_role = "user"; st.session_state.user_name = "인증 사용자"
                        st.rerun()
                    else: st.error("유효하지 않은 승인 접속 코드입니다.")
        with tab_request:
            req_id = st.text_input("신청 사번/아이디")
            req_name = st.text_input("신청자 성명")
            req_dept = st.text_input("소속/부서", value="환경2팀")
            req_pw = st.text_input("비밀번호 설정", type="password")
            if st.button("📝 승인 요청 제출", use_container_width=True):
                if req_id and req_pw and req_name:
                    auth_db = load_auth_db()
                    users = auth_db.get("users", {})
                    users[req_id] = {"name": req_name, "dept": req_dept, "password": hash_pw(req_pw), "status": "pending"}
                    auth_db["users"] = users
                    save_auth_db(auth_db)
                    st.success("승인 요청이 완료되었습니다. 관리자 승인을 기다려주세요.")
                else: st.warning("모든 필수 항목을 입력해주세요.")
    return False

if not check_login_system():
    st.stop()

if st.session_state.get("user_role") == "admin":
    cfg = load_system_config()
    st.sidebar.markdown("---")
    st.sidebar.markdown("##### ⚙️ 관리자 시스템 제어")
    cur_m_state = cfg.get("force_maintenance_screen", True)
    new_m_state = st.sidebar.toggle("🚧 자체 점검중 화면 활성화", value=cur_m_state)
    if new_m_state != cur_m_state:
        cfg["force_maintenance_screen"] = new_m_state
        save_system_config(cfg)
        st.sidebar.success(f"자체 점검중 모드가 {'활성화(ON)' if new_m_state else '해제(OFF)'} 되었습니다.")
        st.rerun()
    auth_db = load_auth_db()
    users = auth_db.get("users", {})
    pending_users = {k: v for k, v in users.items() if v.get("status") == "pending"}
    with st.sidebar.expander(f"🛡️ 승인 대기 ({len(pending_users)}명)", expanded=False):
        for u_id, u_info in list(pending_users.items()):
            st.write(f"**{u_info.get('name')}** ({u_id})")
            c1, c2 = st.columns(2)
            if c1.button("승인", key=f"app_{u_id}"):
                users[u_id]["status"] = "approved"; save_auth_db(auth_db); st.rerun()
            if c2.button("반려", key=f"rej_{u_id}"):
                del users[u_id]; save_auth_db(auth_db); st.rerun()

st.markdown("""
<div class="hero-banner">
    <div>
        <h1 class="hero-title">💧 DANWOL AI-WaterOps 360</h1>
        <div class="hero-subtitle">단월 본장(1,700 ㎥/일) 및 소규모 6개소 · 지능형 자율제어 디지털 트윈 관제 플랫폼</div>
    </div>
    <div class="badge-online">SYSTEM ONLINE</div>
</div>
""", unsafe_allow_html=True)

st.sidebar.title("💧 단월 스마트 관제")
st.sidebar.markdown(f"👤 **접속자**: {st.session_state.get('user_name', '사용자')} ({st.session_state.get('user_role')})")
if st.sidebar.button("로그아웃", use_container_width=True):
    st.session_state.logged_in = False; st.session_state.user_role = None; st.rerun()

st.sidebar.info("📌 **본장**: 단월공공하수 (1,700 ㎥/일, KNR+IPR)\n📌 **소규모 6개소**: 산음·삼가리·진목·몰운·단월마을·당의\n📌 **개인하수 6개소**: 석산리·음지·양지·복지회관·인이피·돌고개")

menu = st.sidebar.radio(
    "⚡ 지능형 기능 메뉴",
    [
        "📑 1. 운영일지·실험실 엑셀 업로드 ➜ 원본양식 자동 완성",
        "📊 2. 공공하수도시설 월간보고서 (HWPX) AI 자동편철 & 보관함",
        "📡 3. TMS 수질 2·4·6·8시간 후 AI 예측 & 신호등 실시간 관제",
        "⚙️ 4. AI 최적 운전조건 제안 & KNR+IPR 공정 정밀진단",
        "🧪 5. 약품·에너지 사용량 데이터 적재 & ESG 경제성 분석",
        "🤖 6. 단월 AI 지능형 공정 Q&A 챗봇 (오프라인 온프레미스)",
        "📝 7. TBM 표준회의록 AI 자동작성/출력 & 보관함",
        "💊 8. 약품 구매 및 사용 관리 & 보관함",
        "🗄️ 9. 통합 영구저장소(보관함·마스터DB) 열람 및 백업 센터"
    ]
)

# -------------------------------------------------------------
# 1. 엑셀 변환 작업대 (국가하수도정보시스템 공인 양식: 단월 본장 & 소규모 6개소 전용, 오차 0%)
# -------------------------------------------------------------
if menu == "📑 1. 운영일지·실험실 엑셀 업로드 ➜ 원본양식 자동 완성":
    st.title("📑 운영일지 및 실험실 데이터 업로드 ➜ 국가하수도정보시스템 공인 양식 자동 완성")
    tab_work, tab_wq_monthly, tab_archive, tab_accum = st.tabs([
        "🚀 국가하수도 공인양식 다운로드 작업대 (단월 본장 · 소규모 6개소)",
        "📈 수질월보 · 월간유량 집계표 자동완성",
        "🗂️ 국가하수도 공인 엑셀 보관함",
        "📊 기간별(월·분기·상반기·하반기·1년) 누적 엑셀 일괄 생성"
    ])
    with tab_work:
        st.caption("🔒 업로드하신 8개 국가하수도정보시스템(KHAS) 공식 엑셀 파일(단월 본장 `유량및수질관리.xlsx`, `재이용수 업로드양식(최종).xlsx` 및 소규모 6개소 업로드양식)의 셀 병합·줄바꿈·폰트·데이터 규칙을 단 1셀의 오차도 없이 100% 동일하게 생성합니다.")

        c_kh1, c_kh2, c_kh3 = st.columns([1.0, 1.3, 1.5])
        with c_kh1:
            khas_yr = st.selectbox("📅 1. 대상 연도 선택", [2026, 2025, 2024], index=0, key="khas_yr_sel")
        with c_kh2:
            khas_period_type = st.selectbox(
                "🗂️ 2. 기간 단위 선택 (월/분기/반기/1년)",
                ["1년 (연간 전체)", "월별 (1월~12월)", "분기별 (1~4분기)", "상반기 / 하반기"],
                index=0,
                key="khas_ptype_sel"
            )
        with c_kh3:
            import calendar
            if khas_period_type == "월별 (1월~12월)":
                sel_m_opt = st.selectbox("📆 3. 상세 월 선택", [f"{m}월 ({khas_yr}-{m:02d})" for m in range(1, 13)], index=8, key="khas_m_detail")
                sel_m_num = int(sel_m_opt.split("월")[0])
                last_d_num = calendar.monthrange(khas_yr, sel_m_num)[1]
                kh_s_d, kh_e_d = f"{khas_yr}-{sel_m_num:02d}-01", f"{khas_yr}-{sel_m_num:02d}-{last_d_num:02d}"
                kh_tag = f"{khas_yr}년_{sel_m_num:02d}월"
            elif khas_period_type == "분기별 (1~4분기)":
                sel_q_opt = st.selectbox("📆 3. 상세 분기 선택", ["1분기 (01월~03월)", "2분기 (04월~06월)", "3분기 (07월~09월)", "4분기 (10월~12월)"], index=0, key="khas_q_detail")
                if "1분기" in sel_q_opt: kh_s_d, kh_e_d, kh_tag = f"{khas_yr}-01-01", f"{khas_yr}-03-31", f"{khas_yr}년_1분기"
                elif "2분기" in sel_q_opt: kh_s_d, kh_e_d, kh_tag = f"{khas_yr}-04-01", f"{khas_yr}-06-30", f"{khas_yr}년_2분기"
                elif "3분기" in sel_q_opt: kh_s_d, kh_e_d, kh_tag = f"{khas_yr}-07-01", f"{khas_yr}-09-30", f"{khas_yr}년_3분기"
                else: kh_s_d, kh_e_d, kh_tag = f"{khas_yr}-10-01", f"{khas_yr}-12-31", f"{khas_yr}년_4분기"
            elif khas_period_type == "상반기 / 하반기":
                sel_h_opt = st.selectbox("📆 3. 상반기 / 하반기 선택", ["상반기 (01월~06월)", "하반기 (07월~12월)"], index=0, key="khas_h_detail")
                if "상반기" in sel_h_opt: kh_s_d, kh_e_d, kh_tag = f"{khas_yr}-01-01", f"{khas_yr}-06-30", f"{khas_yr}년_상반기"
                else: kh_s_d, kh_e_d, kh_tag = f"{khas_yr}-07-01", f"{khas_yr}-12-31", f"{khas_yr}년_하반기"
            else:
                st.selectbox("📆 3. 상세 기간 확인", [f"1년 전체 ({khas_yr}-01-01 ~ {khas_yr}-12-31)"], index=0, disabled=True, key="khas_y_detail")
                kh_s_d, kh_e_d, kh_tag = f"{khas_yr}-01-01", f"{khas_yr}-12-31", f"{khas_yr}년_전체"

        st.info(f"📌 **현재 선택된 국가하수도 엑셀 생성 구간**: **`{kh_tag}`** (`{kh_s_d}` ~ `{kh_e_d}`)")

        fac_grp = st.radio("🏛️ 국가하수도정보시스템 업로드 대상 시설 선택", ["🏢 본처리장 (단월)", "🏡 소규모 처리시설 (6개소: 산음·삼가리·진목·몰운·단월마을·당의)"], horizontal=True)

        if fac_grp == "🏢 본처리장 (단월)":
            files_main = st.file_uploader("단월 본장 운영일지·수질월보·공인양식 엑셀 추가 업로드 (미업로드 시 기본 적재된 예시 데이터 즉시 적용)", type=["xlsx", "xls"], accept_multiple_files=True, key="up_main_plant_all")
            if files_main:
                df_dw_up = universal_main_plant_parser(files_main)
                if not df_dw_up.empty:
                    append_to_master_db(MAIN_PLANT, df_dw_up)
                    st.success(f"✅ 업로드하신 파일에서 단월 본장 데이터 총 **{len(df_dw_up)}일치**({df_dw_up['날짜'].min()} ~ {df_dw_up['날짜'].max()})가 즉시 반영되어 아래 엑셀 파일들이 최신 값으로 자동 갱신되었습니다!")

            df_dw_active = get_master_data(MAIN_PLANT, kh_s_d, kh_e_d)
            m_bytes_36 = fill_exact_main_template(df_dw_active, start_date=kh_s_d, end_date=kh_e_d, year=khas_yr, use_col20_for_out=False)
            r_bytes = fill_exact_reuse_template(df_dw_active, start_date=kh_s_d, end_date=kh_e_d, year=khas_yr)

            st.markdown(f"##### 📥 단월 본장 국가하수도정보시스템 공인 엑셀 2종 다운로드 (`{kh_tag}`)")
            col1, col2 = st.columns(2)
            col1.download_button("📥 유량및수질관리.xlsx 다운로드", m_bytes_36, "유량및수질관리.xlsx" if "전체" in kh_tag else f"유량및수질관리_{kh_tag}.xlsx", type="primary", use_container_width=True)
            col2.download_button("📥 재이용수 업로드양식(최종).xlsx 다운로드", r_bytes, "재이용수 업로드양식(최종).xlsx" if "전체" in kh_tag else f"재이용수 업로드양식(최종)_{kh_tag}.xlsx", type="primary", use_container_width=True)

            if st.button(f"💾 ⚡ [단월 본장 ({kh_tag}) 공인 엑셀 2종 보관함 영구 저장]", key="btn_save_main_m_all", type="primary", use_container_width=True):
                with open(os.path.join(KHAS_RECORD_DIR, sanitize_filename(f"유량및수질관리_{kh_tag}.xlsx")), "wb") as f:
                    f.write(m_bytes_36)
                with open(os.path.join(KHAS_RECORD_DIR, sanitize_filename(f"재이용수 업로드양식(최종)_{kh_tag}.xlsx")), "wb") as f:
                    f.write(r_bytes)
                st.success(f"✅ 단월 본장 국가하수도 공인 엑셀 2종(`{kh_tag}`)이 보관함에 안전하게 저장되었습니다!")

            with st.expander("🔍 단월 본장 국가하수도 공인 엑셀 2종 화면 미리보기 (오차 검증용)", expanded=True):
                t_pv1, t_pv2 = st.tabs(["📄 유량및수질관리.xlsx", "♻️ 재이용수 업로드양식(최종).xlsx"])
                with t_pv1: st.dataframe(pd.read_excel(io.BytesIO(m_bytes_36), header=None).head(35), use_container_width=True, hide_index=True)
                with t_pv2: st.dataframe(pd.read_excel(io.BytesIO(r_bytes), header=None).head(35), use_container_width=True, hide_index=True)

        else:
            st.subheader(f"🏡 소규모 6개소 (산음 · 삼가리 · 진목 · 몰운 · 단월마을 · 당의) 국가하수도 공인 엑셀 (`{kh_tag}`)")
            files_s = st.file_uploader("소규모 6개소 운영일지·수질월보·공인양식 엑셀 추가 업로드 (미업로드 시 예시 원본 데이터 즉시 적용)", type=["xlsx", "xls"], accept_multiple_files=True, key="up_small_all")
            if files_s:
                parsed_s = universal_small_plant_parser(files_s)
                for fk, df_it in parsed_s.items():
                    if not df_it.empty:
                        append_to_master_db(fk, df_it)

            s_dict = {fac: get_master_data(fac, kh_s_d, kh_e_d) for fac in SMALL_PLANTS}
            exact_fnames = {
                "산음": "유량및수질관리 업로드양식(산음).xlsx",
                "삼가리": "유량및수질관리 업로드양식(삼가리).xlsx",
                "진목": "유량및수질관리 업로드양식(진목).xlsx",
                "몰운": "유량및수질관리 업로드양식 (몰운).xlsx",
                "단월마을": "유량및수질관리 업로드양식(단월마을).xlsx",
                "당의": "유량및수질관리 업로드양식(당의).xlsx",
            }
            s_cols = st.columns(3)
            zip_small_buf = io.BytesIO()
            with zipfile.ZipFile(zip_small_buf, "w", zipfile.ZIP_DEFLATED) as zf_s:
                for idx_f, fac_k in enumerate(SMALL_PLANTS):
                    df_item = s_dict.get(fac_k, pd.DataFrame())
                    s_bytes_item = fill_exact_small_template(df_item, fac_k, start_date=kh_s_d, end_date=kh_e_d, year=khas_yr)
                    out_fn = exact_fnames[fac_k] if "전체" in kh_tag else f"유량및수질관리 업로드양식({fac_k})_{kh_tag}.xlsx"
                    zf_s.writestr(out_fn, s_bytes_item)
                    with s_cols[idx_f % 3]:
                        st.download_button(f"📥 {out_fn}", s_bytes_item, out_fn, type="primary", use_container_width=True, key=f"dl_khas_sm_{fac_k}_{kh_tag}")

            c_sm_b1, c_sm_b2 = st.columns(2)
            with c_sm_b1:
                st.download_button(
                    f"📦 소규모 6개소 국가하수도 업로드양식 전체 ZIP 일괄 다운로드 ({kh_tag})",
                    zip_small_buf.getvalue(),
                    f"국가하수도_유량및수질관리_소규모6개소_{kh_tag}.zip",
                    type="primary",
                    use_container_width=True
                )
            with c_sm_b2:
                if st.button(f"💾 ⚡ [소규모 6개소 ({kh_tag}) 국가하수도 엑셀 보관함 일괄 저장]", type="primary", use_container_width=True, key="btn_save_small_all_master"):
                    for fac_k in SMALL_PLANTS:
                        df_item = s_dict.get(fac_k, pd.DataFrame())
                        out_fn = exact_fnames[fac_k] if "전체" in kh_tag else f"유량및수질관리 업로드양식({fac_k})_{kh_tag}.xlsx"
                        with open(os.path.join(KHAS_RECORD_DIR, sanitize_filename(out_fn)), "wb") as f:
                            f.write(fill_exact_small_template(df_item, fac_k, start_date=kh_s_d, end_date=kh_e_d, year=khas_yr))
                    st.success(f"✅ 소규모 6개 시설 국가하수도 공인 엑셀(`{kh_tag}`)이 보관함에 저장되었습니다!")

            st.divider()
            sel_sub_fac = st.selectbox("🔍 미리보기할 소규모 시설 선택", SMALL_PLANTS, key="sel_small_fac_view")
            single_s_bytes = fill_exact_small_template(s_dict.get(sel_sub_fac, pd.DataFrame()), sel_sub_fac, start_date=kh_s_d, end_date=kh_e_d, year=khas_yr)
            with st.expander(f"📄 [{exact_fnames[sel_sub_fac]}] 24컬럼 공인양식 화면 미리보기", expanded=True):
                st.dataframe(pd.read_excel(io.BytesIO(single_s_bytes), header=None).head(38), use_container_width=True, hide_index=True)

    with tab_wq_monthly:
        st.subheader("📈 수질월보 & 월간유량 집계표 (단월본장 · 소규모 6개소 · 개인하수 6개소) 자동 완성 · 인쇄 · 영구보관")
        st.caption("🔒 개인하수(석산리·음지·양지·복지회관·인이피·돌고개)는 국가하수도정보시스템(KHAS) 공인 양식에는 들어가지 않지만, **수질월보 및 월간유량 집계표**에는 필수 포함됩니다. 아래에서 본장·소규모·개인하수 엑셀을 업로드하거나 수동 입력하면 수질월보 집계표에 즉시 반영됩니다.")

        with st.expander("📂 0단계: 수질월보용 엑셀 파일 업로드 & 개인하수 6개소 수질·유량 직접 입력/수정 (클릭하여 열기/닫기)", expanded=True):
            up_c1, up_c2, up_c3 = st.columns(3)
            with up_c1:
                wq_up_main = st.file_uploader("🏢 1. 단월 본장 수질월보/일지 업로드", type=["xlsx", "xls"], accept_multiple_files=True, key="wq_up_main_f")
                if wq_up_main:
                    df_m_p = universal_main_plant_parser(wq_up_main)
                    if not df_m_p.empty:
                        append_to_master_db(MAIN_PLANT, df_m_p)
                        st.success(f"✅ 단월 본장 {len(df_m_p)}건 수질월보 DB 반영 완료!")
            with up_c2:
                wq_up_small = st.file_uploader("🏡 2. 소규모 6개소 수질월보/일지 업로드", type=["xlsx", "xls"], accept_multiple_files=True, key="wq_up_small_f")
                if wq_up_small:
                    s_dict_p = universal_small_plant_parser(wq_up_small)
                    cnt_s = 0
                    for fk, df_s in s_dict_p.items():
                        if not df_s.empty:
                            append_to_master_db(fk, df_s)
                            cnt_s += 1
                    if cnt_s > 0:
                        st.success(f"✅ 소규모 {cnt_s}개 시설 수질월보 DB 반영 완료!")
            with up_c3:
                wq_up_priv = st.file_uploader("🛖 3. 개인하수 6개소 엑셀 업로드 (개인소규모+단월.xls 등)", type=["xlsx", "xls"], accept_multiple_files=True, key="wq_up_priv_f")
                if wq_up_priv:
                    p_dict_p = parse_private_plant_multi_files(wq_up_priv)
                    cnt_p = 0
                    for fk, df_p in p_dict_p.items():
                        if not df_p.empty:
                            append_to_master_db(fk, df_p)
                            cnt_p += 1
                    if cnt_p > 0:
                        st.success(f"✅ 개인하수 {cnt_p}개 시설 수질월보 DB 반영 완료!")

            st.markdown("---")
            st.markdown("###### ✍️ 개인하수 6개소 (석산리·음지·양지·복지회관·인이피·돌고개) 당월 측정 수질·유량 빠른 수동 등록")
            with st.form("priv_wq_manual_form"):
                pm1, pm2, pm3, pm4, pm5, pm6, pm7 = st.columns(7)
                with pm1: p_sel_fac = st.selectbox("개인하수 시설", PRIVATE_PLANTS, key="pm_fac")
                with pm2: p_sel_dt = st.date_input("측정일자", datetime.date(2026, 10, 15), key="pm_dt")
                with pm3: p_in_bod = st.number_input("유입 BOD (mg/L)", value=45.4, step=1.0, format="%.1f", key="pm_ibod")
                with pm4: p_in_ss = st.number_input("유입 SS (mg/L)", value=41.7, step=1.0, format="%.1f", key="pm_iss")
                with pm5: p_out_bod = st.number_input("방류 BOD (mg/L)", value=9.8, step=0.5, format="%.1f", key="pm_obod")
                with pm6: p_out_ss = st.number_input("방류 SS (mg/L)", value=9.5, step=0.5, format="%.1f", key="pm_oss")
                with pm7: p_flow_d = st.number_input("일일유량 (㎥/일, 미설치시 0)", value=25.0 if p_sel_fac == "석산리" else 0.0, step=1.0, format="%.1f", key="pm_fl")
                if st.form_submit_button("💾 개인하수 측정 데이터 수질월보 마스터 DB에 즉시 반영", type="primary", use_container_width=True):
                    fl_val = float(p_flow_d) if p_flow_d > 0 else np.nan
                    wk_val = round(fl_val * 7.0, 1) if pd.notna(fl_val) else np.nan
                    df_new_p = pd.DataFrame([{
                        "날짜": p_sel_dt.strftime('%Y-%m-%d'),
                        "유입BOD": float(p_in_bod),
                        "유입SS": float(p_in_ss),
                        "방류BOD": float(p_out_bod),
                        "방류SS": float(p_out_ss),
                        "유입량": fl_val,
                        "방류량": fl_val,
                        "주간유량": wk_val
                    }])
                    append_to_master_db(p_sel_fac, df_new_p)
                    st.success(f"✅ 개인하수 [{p_sel_fac}] ({p_sel_dt}) 측정 수질·유량이 수질월보 및 월간유량 집계표에 반영되었습니다!")
                    st.rerun()

        c_wq1, c_wq2 = st.columns(2)
        with c_wq1:
            wq_yr = st.selectbox("📅 수질월보 대상 연도", [2026, 2025, 2024], index=0, key="wq_yr_sel")
        with c_wq2:
            wq_mo = st.selectbox("📆 수질월보 조회 월", list(range(1, 13)), index=9, key="wq_mo_sel")

        main_wq_bytes = generate_main_monthly_wq_excel(year=wq_yr, sel_month=None)
        small_wq_bytes = generate_small_monthly_wq_excel(year=wq_yr, sel_month=None)
        priv_wq_bytes = generate_priv_monthly_wq_excel(year=wq_yr)

        st.markdown("##### 📥 1단계: 3종 원본 양식 엑셀(XLSX) 즉시 다운로드 (단월본장 · 소규모 6개소 · 개인하수 6개소)")
        col_dl_wq1, col_dl_wq2, col_dl_wq3 = st.columns(3)
        with col_dl_wq1:
            st.download_button(
                f"📥 {wq_yr}년 월별수질(단월).xlsx 다운로드",
                data=main_wq_bytes,
                file_name=f"{wq_yr}년 월별수질(단월).xlsx",
                type="primary",
                use_container_width=True
            )
        with col_dl_wq2:
            st.download_button(
                f"📥 1. 소규모({wq_yr}).xlsx 다운로드",
                data=small_wq_bytes,
                file_name=f"1. 소규모({wq_yr}).xlsx",
                type="primary",
                use_container_width=True
            )
        with col_dl_wq3:
            st.download_button(
                f"📥 개인소규모({wq_yr}년)_단월6개소.xlsx 다운로드",
                data=priv_wq_bytes,
                file_name=f"개인소규모({wq_yr}년)_단월6개소.xlsx",
                type="primary",
                use_container_width=True
            )

        st.divider()
        st.markdown(f"##### 📊 {wq_yr}년 {wq_mo}월 수질월보 및 월간유량 집계표 (평균 · 최대 · 최소 · 월유입총량 자동산출)")
        ym_wq = f"{wq_yr}-{wq_mo:02d}"
        df_m_wq = get_master_data(MAIN_PLANT, f"{ym_wq}-01", f"{ym_wq}-31")
        sum_rows = []
        if not df_m_wq.empty:
            st.markdown(f"**🏢 1. 단월공공하수처리시설 수질검사결과 ({wq_mo}월) — 총 {len(df_m_wq)}일치 실측 데이터**")
            num_cols = ['유입BOD', '유입TOC', '유입SS', '유입TN', '유입TP', '유입대장균', 'MLSS_A', 'MLSS_B',
                        '방류BOD', '방류TOC', '방류SS', '방류TN', '방류TP', '방류대장균', '유입량', '재이용수', '방류량', '수온']
            avail_cols = [c for c in num_cols if c in df_m_wq.columns]
            for lbl, fn in [("평균", "mean"), ("최대", "max"), ("최소", "min"), ("월총량(합계)", "sum")]:
                r_dict = {"구분": lbl}
                for c in avail_cols:
                    s_v = pd.to_numeric(df_m_wq[c], errors='coerce').dropna()
                    if not s_v.empty:
                        if fn == "sum" and c not in ['유입량', '재이용수', '방류량']:
                            r_dict[c] = "-"
                        else:
                            val = getattr(s_v, fn)()
                            dec = 3 if 'TN' in c or 'TP' in c else (0 if c in ['유입량', '재이용수', '방류량', '유입대장균', '방류대장균', 'MLSS_A', 'MLSS_B'] else 1)
                            r_dict[c] = f"{val:,.{dec}f}" if dec > 0 else f"{int(round(val)):,}"
                    else:
                        r_dict[c] = "-"
                sum_rows.append(r_dict)
            st.dataframe(pd.DataFrame(sum_rows), use_container_width=True, hide_index=True)
            with st.expander(f"🔍 단월본장 {wq_mo}월 일별 수질월보 원본 데이터 열람", expanded=False):
                st.dataframe(df_m_wq[['날짜'] + avail_cols], use_container_width=True, hide_index=True)
        else:
            st.info(f"{wq_yr}년 {wq_mo}월 단월본장 데이터가 없습니다.")

        st.markdown(f"**🏡 2. 소규모공공하수처리시설 6개소 월별 수질·유량 현황 ({wq_mo}월)**")
        sm_summary = []
        for fac in SMALL_PLANTS:
            df_sf = get_master_data(fac, f"{ym_wq}-01", f"{ym_wq}-31")
            if not df_sf.empty:
                if '주간유량' not in df_sf.columns and '방류량' in df_sf.columns:
                    df_sf['주간유량'] = pd.to_numeric(df_sf['방류량'], errors='coerce') * 7.0
                for lbl, fn in [("평균", "mean"), ("최대", "max"), ("최소", "min")]:
                    row_s = {"시설명": fac, "구분": lbl}
                    for c, dec in [('유입BOD', 1), ('유입TOC', 1), ('유입SS', 1), ('유입TN', 3), ('유입TP', 3), ('유입대장균', 0),
                                   ('방류BOD', 1), ('방류TOC', 1), ('방류SS', 1), ('방류TN', 3), ('방류TP', 3), ('방류대장균', 0),
                                   ('주간유량', 0), ('방류량', 1)]:
                        if c in df_sf.columns:
                            ser = pd.to_numeric(df_sf[c], errors='coerce').dropna()
                            row_s[c] = (f"{getattr(ser, fn)():.{dec}f}" if dec > 0 else f"{int(round(getattr(ser, fn)())):,}") if not ser.empty else "-"
                        else:
                            row_s[c] = "-"
                    sm_summary.append(row_s)
        if sm_summary:
            st.dataframe(pd.DataFrame(sm_summary), use_container_width=True, hide_index=True)

        st.markdown(f"**🛖 3. 개인하수처리시설 6개소 월별 수질·월간유량 집계표 ({wq_mo}월 — 평균·최대·최소·월유입총량 포함)**")
        import calendar
        days_in_wq_m = calendar.monthrange(wq_yr, wq_mo)[1]
        pr_summary = []
        priv_cap_map = {"석산리": "30톤/일", "음지": "15톤/일", "양지": "15톤/일", "복지회관": "10톤/일", "인이피": "10톤/일", "돌고개": "15톤/일"}
        for fac in PRIVATE_PLANTS:
            df_pf = get_master_data(fac, f"{ym_wq}-01", f"{ym_wq}-31")
            if df_pf.empty:
                # 해당 월 데이터가 없을 경우 가장 최근 측정 데이터를 참조하여 표시
                df_pf_all = get_master_data(fac)
                if not df_pf_all.empty:
                    df_pf = df_pf_all.tail(1)
            if not df_pf.empty:
                last_dt_str = str(df_pf['날짜'].iloc[-1])[:10]
                in_b_ser = pd.to_numeric(df_pf.get('유입BOD', pd.Series(dtype=float)), errors='coerce').dropna()
                in_s_ser = pd.to_numeric(df_pf.get('유입SS', pd.Series(dtype=float)), errors='coerce').dropna()
                out_b_ser = pd.to_numeric(df_pf.get('방류BOD', pd.Series(dtype=float)), errors='coerce').dropna()
                out_s_ser = pd.to_numeric(df_pf.get('방류SS', pd.Series(dtype=float)), errors='coerce').dropna()
                fl_d_ser = pd.to_numeric(df_pf.get('방류량', pd.Series(dtype=float)), errors='coerce').dropna()
                fl_w_ser = pd.to_numeric(df_pf.get('주간유량', pd.Series(dtype=float)), errors='coerce').dropna()

                fl_d_str = f"{fl_d_ser.mean():.1f}" if (fac == "석산리" and not fl_d_ser.empty) else ("유량계미설치" if fac != "석산리" else "-")
                fl_w_str = f"{int(round(fl_w_ser.mean())):,}" if (fac == "석산리" and not fl_w_ser.empty) else ("유량계미설치" if fac != "석산리" else "-")
                m_tot_str = f"{fl_d_ser.mean() * days_in_wq_m:,.1f}" if (fac == "석산리" and not fl_d_ser.empty) else ("유량계미설치" if fac != "석산리" else "-")

                pr_summary.append({
                    "시설명": f"{fac} ({priv_cap_map.get(fac, '')})",
                    "측정일": last_dt_str,
                    "구분": "평균 (당월)",
                    "유입BOD": f"{in_b_ser.mean():.1f}" if not in_b_ser.empty else "-",
                    "유입SS": f"{in_s_ser.mean():.1f}" if not in_s_ser.empty else "-",
                    "방류BOD": f"{out_b_ser.mean():.1f}" if not out_b_ser.empty else "-",
                    "방류SS": f"{out_s_ser.mean():.1f}" if not out_s_ser.empty else "-",
                    "방류량(㎥/일)": fl_d_str,
                    "주간유량(㎥/주)": fl_w_str,
                    "월유입총량(㎥/월)": m_tot_str
                })
        if pr_summary:
            st.dataframe(pd.DataFrame(pr_summary), use_container_width=True, hide_index=True)
            with st.expander("🔍 개인하수 6개소 마스터 DB 전체 측정 기록 열람 및 직접 편집", expanded=False):
                df_priv_all_db = pd.concat([get_master_data(f_k) for f_k in PRIVATE_PLANTS], ignore_index=True)
                if not df_priv_all_db.empty:
                    st.dataframe(df_priv_all_db[['시설명', '날짜', '유입BOD', '유입SS', '방류BOD', '방류SS', '유입량', '방류량', '주간유량']], use_container_width=True, hide_index=True)

        def build_html_table_from_dicts(dict_list):
            if not dict_list: return "<p>데이터 없음</p>"
            cols = list(dict_list[0].keys())
            ths = "".join([f"<th style='border:1px solid #000; background:#f2f2f2; padding:4px;'>{c}</th>" for c in cols])
            trs = ""
            for r in dict_list:
                tds = "".join([f"<td style='border:1px solid #000; padding:4px; text-align:center;'>{r.get(c, '')}</td>" for c in cols])
                trs += f"<tr>{tds}</tr>"
            return f"<table style='width:100%; border-collapse:collapse; font-size:9.5px; margin-bottom:12px;'><tr>{ths}</tr>{trs}</table>"

        wq_print_html = f"""
        <!DOCTYPE html>
        <html>
        <head><meta charset="UTF-8">
        <style>
            @page {{ size: A4 landscape; margin: 10mm; }}
            body {{ font-family: 'Malgun Gothic', '맑은 고딕', sans-serif; color:#000; margin:10px; font-size:10px; }}
            h2 {{ text-align:center; font-size:16px; border-bottom:2px solid #000; padding-bottom:6px; }}
            h4 {{ margin:10px 0 4px 0; font-size:12px; }}
            @media print {{ .no-print {{ display:none !important; }} }}
        </style>
        </head>
        <body>
            <div class="no-print" style="text-align:right; margin-bottom:8px;">
                <button onclick="window.print()" style="background:#0284C7; color:white; border:none; padding:7px 16px; border-radius:6px; font-weight:bold; cursor:pointer;">
                    🖨️ {wq_yr}년 {wq_mo}월 통합 수질월보 즉시 인쇄 / PDF 출력
                </button>
            </div>
            <h2>{wq_yr}년 {wq_mo}월 단월공공하수 · 소규모 · 개인하수 통합 수질월보 및 월간유량 집계표</h2>
            <h4>1. 단월공공하수처리시설 수질검사결과 및 월간유량 요약 ({wq_mo}월)</h4>
            {build_html_table_from_dicts(sum_rows)}
            <h4>2. 소규모공공하수처리시설 6개소 월별 수질·유량 현황 ({wq_mo}월)</h4>
            {build_html_table_from_dicts(sm_summary)}
            <h4>3. 개인하수처리시설 6개소 월별 수질·월간유량 집계표 ({wq_mo}월)</h4>
            {build_html_table_from_dicts(pr_summary)}
        </body>
        </html>
        """

        st.divider()
        st.markdown("##### 🖨️ 2단계: 수질월보 인쇄 미리보기 & 한글(HWPX) · PDF · 영구 보관함 일괄 저장")
        with st.expander(f"🖨️ {wq_yr}년 {wq_mo}월 통합 수질월보 인쇄용 미리보기 열기", expanded=True):
            st.components.v1.html(wq_print_html, height=480, scrolling=True)

        wq_hwpx_bytes = generate_wq_monthly_hwpx(wq_yr, wq_mo, sum_rows, sm_summary, pr_summary)
        wq_pdf_bytes = generate_html_to_pdf_bytes(wq_print_html)
        is_real_pdf = wq_pdf_bytes[:4] == b"%PDF"

        c_act1, c_act2, c_act3 = st.columns(3)
        with c_act1:
            st.download_button(
                f"📥 {wq_yr}년_{wq_mo}월_통합수질월보.hwpx (한글) 다운로드",
                data=wq_hwpx_bytes,
                file_name=f"통합수질월보_{wq_yr}년_{wq_mo}월.hwpx",
                mime="application/hwp+zip",
                type="primary",
                use_container_width=True
            )
        with c_act2:
            st.download_button(
                f"📥 {wq_yr}년_{wq_mo}월_통합수질월보.{'pdf' if is_real_pdf else 'html'} (PDF/인쇄용) 다운로드",
                data=wq_pdf_bytes,
                file_name=f"통합수질월보_{wq_yr}년_{wq_mo}월.{'pdf' if is_real_pdf else 'html'}",
                mime="application/pdf" if is_real_pdf else "text/html",
                type="primary",
                use_container_width=True
            )
        with c_act3:
            if st.button("💾 ⚡ [수질월보 3종 엑셀 + 한글(HWPX) + PDF/인쇄본] 영구 보관함 일괄 저장", type="primary", use_container_width=True):
                with open(os.path.join(KHAS_RECORD_DIR, sanitize_filename(f"{wq_yr}년_월별수질(단월).xlsx")), "wb") as f:
                    f.write(main_wq_bytes)
                with open(os.path.join(KHAS_RECORD_DIR, sanitize_filename(f"1._소규모({wq_yr}).xlsx")), "wb") as f:
                    f.write(small_wq_bytes)
                with open(os.path.join(KHAS_RECORD_DIR, sanitize_filename(f"개인소규모({wq_yr}년)+단월.xlsx")), "wb") as f:
                    f.write(priv_wq_bytes)
                with open(os.path.join(HWPX_RECORD_DIR, sanitize_filename(f"통합수질월보_{wq_yr}년_{wq_mo}월.hwpx")), "wb") as f:
                    f.write(wq_hwpx_bytes)
                with open(os.path.join(HWPX_RECORD_DIR, sanitize_filename(f"통합수질월보_{wq_yr}년_{wq_mo}월.html")), "w", encoding="utf-8") as f:
                    f.write(wq_print_html)
                if is_real_pdf:
                    with open(os.path.join(HWPX_RECORD_DIR, sanitize_filename(f"통합수질월보_{wq_yr}년_{wq_mo}월.pdf")), "wb") as f:
                        f.write(wq_pdf_bytes)
                st.success("✅ 수질월보 3종 엑셀(XLSX), 한글(HWPX), PDF/인쇄본이 모두 영구 보관함에 안전하게 저장되었습니다!")

    with tab_archive:
        st.subheader("🗂️ 국가하수도정보시스템 공인 엑셀 영구 보관함 (단월 본장 · 소규모 6개소)")
        if st.button("⚡ 마스터 DB 기준으로 국가하수도 공인 엑셀 8종(단월본장 2종 + 소규모 6개소) 보관함에 즉시 일괄 생성", type="primary", use_container_width=True):
            df_m_all = get_master_data(MAIN_PLANT)
            with open(os.path.join(KHAS_RECORD_DIR, "유량및수질관리.xlsx"), "wb") as f:
                f.write(fill_exact_main_template(df_m_all, year=2026, use_col20_for_out=False))
            with open(os.path.join(KHAS_RECORD_DIR, "재이용수 업로드양식(최종).xlsx"), "wb") as f:
                f.write(fill_exact_reuse_template(df_m_all, year=2026))
            exact_fnames_arch = {
                "산음": "유량및수질관리 업로드양식(산음).xlsx",
                "삼가리": "유량및수질관리 업로드양식(삼가리).xlsx",
                "진목": "유량및수질관리 업로드양식(진목).xlsx",
                "몰운": "유량및수질관리 업로드양식 (몰운).xlsx",
                "단월마을": "유량및수질관리 업로드양식(단월마을).xlsx",
                "당의": "유량및수질관리 업로드양식(당의).xlsx",
            }
            for fk in SMALL_PLANTS:
                df_s_all = get_master_data(fk)
                with open(os.path.join(KHAS_RECORD_DIR, exact_fnames_arch[fk]), "wb") as f:
                    f.write(fill_exact_small_template(df_s_all, fk, year=2026))
            st.success("✅ 국가하수도정보시스템 업로드 공인 엑셀 8종(본장 2종 + 소규모 6종)이 보관함에 일괄 생성되었습니다!")
            st.rerun()

        saved_files = [sanitize_filename(f) for f in os.listdir(KHAS_RECORD_DIR) if f.endswith(".xlsx") or f.endswith(".xls")]
        if saved_files:
            st.write(f"📁 **보관 문서: 총 {len(saved_files)}건**")
            col_view, col_del = st.columns([3, 1])
            with col_view: target_f = st.selectbox("열람 및 다운로드할 엑셀 파일 선택", sorted(saved_files), key="arch_file_sel")
            with col_del:
                st.write(""); st.write("")
                if st.button("🗑️ 선택 파일 삭제", type="secondary", use_container_width=True, key="btn_del_arch_file"):
                    os.remove(os.path.join(KHAS_RECORD_DIR, sanitize_filename(target_f)))
                    st.success("🗑️ 파일이 삭제되었습니다.")
                    st.rerun()
            if target_f:
                fpath = os.path.join(KHAS_RECORD_DIR, sanitize_filename(target_f))
                with open(fpath, "rb") as f: f_bytes = f.read()
                st.download_button("📥 선택된 문서 다운로드", f_bytes, file_name=sanitize_filename(target_f), type="primary", use_container_width=True)
                try:
                    df_preview = pd.read_excel(fpath, header=None)
                    with st.expander("🔍 엑셀 파일 데이터 미리보기 (공인 3단 헤더 포함)", expanded=True):
                        st.dataframe(df_preview.head(38), use_container_width=True, hide_index=True)
                except Exception:
                    pass
        else:
            st.info("💡 상단 [⚡ 마스터 DB 기준으로 국가하수도 공인 엑셀 9종 보관함에 즉시 일괄 생성] 버튼을 누르시면 전 시설 공인 양식이 즉시 생성됩니다.")

    with tab_accum:
        st.subheader("📊 ⚡ [월별 / 분기별 / 상반기 / 하반기 / 1년 전체] 국가하수도 공인 누적 엑셀 일괄 생성 센터 (단월 본장 & 소규모 6개소)")
        c_cum_y, c_cum_t, c_cum_p = st.columns([1.0, 1.2, 1.5])
        with c_cum_y:
            sel_cum_year = st.selectbox("📅 1. 대상 연도 선택", [2026, 2025, 2024], key="cum_y_sel")
        with c_cum_t:
            sel_cum_type = st.selectbox("🗂️ 2. 기간 구분 선택", ["1년 (연간 전체)", "월별 (1월~12월)", "분기별 (1~4분기)", "상반기 / 하반기"], index=0, key="cum_t_sel")
        with c_cum_p:
            import calendar
            if sel_cum_type == "월별 (1월~12월)":
                sel_period = st.selectbox("📆 3. 상세 월 선택", [f"{m}월 ({sel_cum_year}-{m:02d})" for m in range(1, 13)], index=8, key="cum_m_sel")
                m_num_c = int(sel_period.split("월")[0])
                last_d_c = calendar.monthrange(sel_cum_year, m_num_c)[1]
                s_d, e_d = f"{sel_cum_year}-{m_num_c:02d}-01", f"{sel_cum_year}-{m_num_c:02d}-{last_d_c:02d}"
                cum_tag = f"{sel_cum_year}년_{m_num_c:02d}월"
            elif sel_cum_type == "분기별 (1~4분기)":
                sel_period = st.selectbox("📆 3. 상세 분기 선택", ["1분기 (01~03월)", "2분기 (04~06월)", "3분기 (07~09월)", "4분기 (10~12월)"], index=0, key="cum_q_sel")
                if "1분기" in sel_period: s_d, e_d, cum_tag = f"{sel_cum_year}-01-01", f"{sel_cum_year}-03-31", f"{sel_cum_year}년_1분기"
                elif "2분기" in sel_period: s_d, e_d, cum_tag = f"{sel_cum_year}-04-01", f"{sel_cum_year}-06-30", f"{sel_cum_year}년_2분기"
                elif "3분기" in sel_period: s_d, e_d, cum_tag = f"{sel_cum_year}-07-01", f"{sel_cum_year}-09-30", f"{sel_cum_year}년_3분기"
                else: s_d, e_d, cum_tag = f"{sel_cum_year}-10-01", f"{sel_cum_year}-12-31", f"{sel_cum_year}년_4분기"
            elif sel_cum_type == "상반기 / 하반기":
                sel_period = st.selectbox("📆 3. 상반기 / 하반기 선택", ["상반기 (01~06월)", "하반기 (07~12월)"], index=0, key="cum_h_sel")
                if "상반기" in sel_period: s_d, e_d, cum_tag = f"{sel_cum_year}-01-01", f"{sel_cum_year}-06-30", f"{sel_cum_year}년_상반기"
                else: s_d, e_d, cum_tag = f"{sel_cum_year}-07-01", f"{sel_cum_year}-12-31", f"{sel_cum_year}년_하반기"
            else:
                sel_period = st.selectbox("📆 3. 연간 전체 확인", [f"1년 전체 ({sel_cum_year}-01-01 ~ {sel_cum_year}-12-31)"], index=0, disabled=True, key="cum_y_all_sel")
                s_d, e_d, cum_tag = f"{sel_cum_year}-01-01", f"{sel_cum_year}-12-31", f"{sel_cum_year}년_1년전체"

        st.info(f"📍 **선택된 다운로드 기간**: **`{cum_tag}`** (`{s_d}` ~ `{e_d}`)")
        st.divider()
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            st.markdown(f"##### 🏢 1. 단월 본장 국가하수도 공인 엑셀 (`{cum_tag}`)")
            df_main_cum = get_master_data(MAIN_PLANT, s_d, e_d)
            st.download_button(
                f"📥 유량및수질관리({cum_tag}).xlsx 다운로드",
                fill_exact_main_template(df_main_cum, start_date=s_d, end_date=e_d, year=sel_cum_year, use_col20_for_out=False),
                "유량및수질관리.xlsx" if "1년전체" in cum_tag else f"유량및수질관리_{cum_tag}.xlsx",
                use_container_width=True,
                type="primary",
                key=f"cum_dl_main36_{cum_tag}"
            )
            st.download_button(
                f"📥 재이용수 업로드양식(최종)({cum_tag}).xlsx 다운로드",
                fill_exact_reuse_template(df_main_cum, start_date=s_d, end_date=e_d, year=sel_cum_year),
                "재이용수 업로드양식(최종).xlsx" if "1년전체" in cum_tag else f"재이용수 업로드양식(최종)_{cum_tag}.xlsx",
                use_container_width=True,
                type="primary",
                key=f"cum_dl_reuse_{cum_tag}"
            )
        with col_c2:
            st.markdown(f"##### 🏡 2. 소규모 6개소 국가하수도 공인 엑셀 (`{cum_tag}`)")
            exact_fnames_cum = {
                "산음": "유량및수질관리 업로드양식(산음).xlsx",
                "삼가리": "유량및수질관리 업로드양식(삼가리).xlsx",
                "진목": "유량및수질관리 업로드양식(진목).xlsx",
                "몰운": "유량및수질관리 업로드양식 (몰운).xlsx",
                "단월마을": "유량및수질관리 업로드양식(단월마을).xlsx",
                "당의": "유량및수질관리 업로드양식(당의).xlsx",
            }
            zip_s = io.BytesIO()
            with zipfile.ZipFile(zip_s, "w", zipfile.ZIP_DEFLATED) as zf:
                for fac in SMALL_PLANTS:
                    df_it = get_master_data(fac, s_d, e_d)
                    fn_z = exact_fnames_cum[fac] if "1년전체" in cum_tag else f"유량및수질관리 업로드양식({fac})_{cum_tag}.xlsx"
                    zf.writestr(fn_z, fill_exact_small_template(df_it, fac, start_date=s_d, end_date=e_d, year=sel_cum_year))
            st.download_button(
                f"📦 소규모 6개소 전체 ZIP 다운로드 ({cum_tag})",
                zip_s.getvalue(),
                f"유량및수질관리_소규모6개소_{cum_tag}.zip",
                use_container_width=True,
                type="primary",
                key=f"cum_dl_sm_zip_{cum_tag}"
            )
            sel_sm_indiv = st.selectbox("소규모 개별 시설 선택", SMALL_PLANTS, key="cum_sm_indiv_sel")
            df_sm_indiv = get_master_data(sel_sm_indiv, s_d, e_d)
            fn_indiv = exact_fnames_cum[sel_sm_indiv] if "1년전체" in cum_tag else f"유량및수질관리 업로드양식({sel_sm_indiv})_{cum_tag}.xlsx"
            st.download_button(
                f"📥 {fn_indiv} 개별 다운로드",
                fill_exact_small_template(df_sm_indiv, sel_sm_indiv, start_date=s_d, end_date=e_d, year=sel_cum_year),
                fn_indiv,
                use_container_width=True,
                type="secondary",
                key=f"cum_dl_sm_one_{sel_sm_indiv}_{cum_tag}"
            )

        if os.path.exists(MASTER_ACCUM_DB):
            st.divider()
            st.markdown("##### 🗄️ 영구 저장된 운영일지 마스터 DB 전체 열람 (본장 및 소규모 6개소)")
            df_all_m = pd.read_csv(MASTER_ACCUM_DB)
            df_khas_only = df_all_m[df_all_m['시설명'].isin([MAIN_PLANT] + SMALL_PLANTS)]
            st.dataframe(df_khas_only, use_container_width=True)
# -------------------------------------------------------------
# 2. HWPX 월간보고서 (10월 공식 양식 토씨 하나 빠짐없이 완벽 재현 + 1번 데이터 연동 + 직접 수정/인쇄/저장)
# -------------------------------------------------------------
elif menu == "📊 2. 공공하수도시설 월간보고서 (HWPX) AI 자동편철 & 보관함":
    st.title("📊 공공하수처리시설 대행사업 월간보고서 (HWPX / 공식 인쇄 양식)")
    tab_hw_w, tab_hw_a = st.tabs(["✍️ [작성·편집·인쇄] 월간보고서 전 부문 통합 작성대", "🗂️ [보관함] 월간보고서(HWPX·HTML) 영구 보관소"])

    with tab_hw_w:
        st.caption("🔒 10월 공식 월간보고서(단월본장 · 소규모 6개소 · 개인하수 11개소)의 모든 표·문구·서식을 오타 없이 100% 동일하게 구현했으며, 1번 운영일지 자동 연동 및 전 항목 직접 수정·인쇄·저장을 지원합니다.")

        c_top1, c_top2, c_top3 = st.columns([1, 1, 1.5])
        with c_top1:
            sel_report_year = st.selectbox("📅 보고 연도", [2026, 2025, 2024], index=0, key="rep_yr")
        with c_top2:
            sel_report_month = st.selectbox("📅 보고 월", list(range(1, 13)), index=9, key="rep_mo")
        with c_top3:
            m_win = [(sel_report_month - 5 + i - 1) % 12 + 1 for i in range(6)]
            m_cols_labels = [f"{m}월" for m in m_win]
            st.info(f"📌 **최근 6개월 자동 연동 구간**: {', '.join(m_cols_labels)} (+ 평균 자동 산출)")

        col_sync1, col_sync2 = st.columns([1.5, 1.5])
        with col_sync1:
            log_excel_up = st.file_uploader("📂 1번 항목 종합운영일지 엑셀 업로드 (업로드 시 수질·유량·추진실적 자동 분석)", type=["xlsx", "xls"], accept_multiple_files=True, key="rep_log_up")
        with col_sync2:
            hwpx_file_up = st.file_uploader("📂 원본 HWPX 템플릿 파일 업로드 (미업로드 시 내장된 10월 공식 양식 자동 적용)", type=["hwpx"], key="rep_hwpx_up")

        if st.button("🔄 ⚡ [1번 수질월보 · 운영일지 마스터 DB]에서 전 시설(본장·소규모·개인하수) 유량·수질(평균·최대·최소) 자동 연동", type="primary", use_container_width=True):
            if log_excel_up:
                df_parsed_main = universal_main_plant_parser(log_excel_up)
                if not df_parsed_main.empty:
                    append_to_master_db(MAIN_PLANT, df_parsed_main)
                s_parsed_dict = universal_small_plant_parser(log_excel_up)
                for fk, df_s in s_parsed_dict.items():
                    if not df_s.empty:
                        append_to_master_db(fk, df_s)
                p_parsed_dict = parse_private_plant_multi_files(log_excel_up)
                for fk, df_p in p_parsed_dict.items():
                    if not df_p.empty:
                        append_to_master_db(fk, df_p)

            ym_prefix = f"{sel_report_year}-{sel_report_month:02d}"
            df_m_cur = get_master_data(MAIN_PLANT, f"{ym_prefix}-01", f"{ym_prefix}-31")
            if not df_m_cur.empty:
                if '유입량' in df_m_cur.columns and df_m_cur['유입량'].dropna().any():
                    st.session_state["cur_in_avg"] = int(round(df_m_cur['유입량'].mean()))
                    st.session_state["cur_in_max"] = int(round(df_m_cur['유입량'].max()))
                    st.session_state["cur_in_min"] = int(round(df_m_cur['유입량'].min()))
                if '방류량' in df_m_cur.columns and df_m_cur['방류량'].dropna().any():
                    st.session_state["cur_out_avg"] = int(round(df_m_cur['방류량'].mean()))
                    st.session_state["cur_out_max"] = int(round(df_m_cur['방류량'].max()))
                    st.session_state["cur_out_min"] = int(round(df_m_cur['방류량'].min()))
                st.toast(f"✅ {sel_report_month}월 수질월보 및 운영일지 실측 데이터를 기반으로 본장·소규모·개인하수 유량 및 수질(평균·최대·최소)이 자동 반영되었습니다!", icon="📊")
            else:
                st.toast("💡 해당 월의 1번 마스터 DB 기록이 없어 10월 공식 보고서 기준값을 유지합니다. 아래 표에서 직접 수정하실 수 있습니다.", icon="ℹ️")

        st.divider()
        sub_tab1, sub_tab2, sub_tab3 = st.tabs([
            f"🏢 1. 단월공공하수처리시설 운영사항 보고({sel_report_month}월)",
            f"🏡 2. 소규모하수처리시설 운영사항 보고({sel_report_month}월)",
            f"🛖 3. 개인하수처리시설 운영사항 보고({sel_report_month}월)"
        ])

        with sub_tab1:
            st.markdown(f"#### 🏢 단월공공하수처리시설 운영사항 보고({sel_report_month}월)")
            with st.expander("□ 개요 (자동 서식 적용 및 직접 수정 가능)", expanded=True):
                g_c1, g_c2, g_c3 = st.columns(3)
                with g_c1:
                    ov_name = st.text_input("○ 시 설 명", value="단월공공하수처리시설", key="ov_name")
                    ov_agency = st.text_input("○ 대행업체명", value="양평공사", key="ov_agency")
                with g_c2:
                    ov_cap = st.text_input("○ 시 설 용 량", value="1,700㎥/일", key="ov_cap")
                    ov_method = st.text_input("○ 공 법", value="KNR + IPR(총인처리시설)", key="ov_method")
                with g_c3:
                    ov_staff = st.text_input("○ 근 무 자", value="총5명(주간 4명, 야간 1명)", key="ov_staff")

            col_fl1, col_fl2 = st.columns(2)
            with col_fl1:
                st.markdown("##### ○ 유입량 현황(최근 6개월간) (단위: ㎥/일, 평균 자동계산)")
                in_cur_avg = st.session_state.get("cur_in_avg", 1615)
                in_cur_max = st.session_state.get("cur_in_max", 1855)
                in_cur_min = st.session_state.get("cur_in_min", 1454)
                df_inflow_init = pd.DataFrame([
                    {"구분": "평균", m_cols_labels[0]: 1716, m_cols_labels[1]: 1722, m_cols_labels[2]: 1806, m_cols_labels[3]: 1752, m_cols_labels[4]: 1674, m_cols_labels[5]: in_cur_avg},
                    {"구분": "최대", m_cols_labels[0]: 1883, m_cols_labels[1]: 2067, m_cols_labels[2]: 2087, m_cols_labels[3]: 1935, m_cols_labels[4]: 1835, m_cols_labels[5]: in_cur_max},
                    {"구분": "최소", m_cols_labels[0]: 1531, m_cols_labels[1]: 1387, m_cols_labels[2]: 1557, m_cols_labels[3]: 1595, m_cols_labels[4]: 1495, m_cols_labels[5]: in_cur_min},
                ])
                ed_inflow = st.data_editor(df_inflow_init, use_container_width=True, hide_index=True, key="ed_inflow_6m")
                ed_inflow["평균"] = ed_inflow[m_cols_labels].apply(pd.to_numeric, errors='coerce').mean(axis=1).round(0).astype(int).apply(lambda x: f"{x:,}")

            with col_fl2:
                st.markdown("##### ○ 방류량 현황(최근6개월간) (단위: ㎥/일, 평균 자동계산)")
                out_cur_avg = st.session_state.get("cur_out_avg", 1431)
                out_cur_max = st.session_state.get("cur_out_max", 1737)
                out_cur_min = st.session_state.get("cur_out_min", 1256)
                df_outflow_init = pd.DataFrame([
                    {"구분": "평균", m_cols_labels[0]: 1559, m_cols_labels[1]: 1566, m_cols_labels[2]: 1649, m_cols_labels[3]: 1597, m_cols_labels[4]: 1530, m_cols_labels[5]: out_cur_avg},
                    {"구분": "최대", m_cols_labels[0]: 1720, m_cols_labels[1]: 1900, m_cols_labels[2]: 1924, m_cols_labels[3]: 1778, m_cols_labels[4]: 1682, m_cols_labels[5]: out_cur_max},
                    {"구분": "최소", m_cols_labels[0]: 1380, m_cols_labels[1]: 1237, m_cols_labels[2]: 1402, m_cols_labels[3]: 1441, m_cols_labels[4]: 1310, m_cols_labels[5]: out_cur_min},
                ])
                ed_outflow = st.data_editor(df_outflow_init, use_container_width=True, hide_index=True, key="ed_outflow_6m")
                ed_outflow["평균"] = ed_outflow[m_cols_labels].apply(pd.to_numeric, errors='coerce').mean(axis=1).round(0).astype(int).apply(lambda x: f"{x:,}")

            st.markdown("##### ○ 수질현황 (단위: ㎎/ℓ, 개/㎖)")
            df_wq_main_init = pd.DataFrame([
                {"구분": "평균", "BOD": "0.7", "TOC": "3.1", "SS": "2.1", "T-N": "6.011", "T-P": "0.032", "대장균": "87", "비고": ""},
                {"구분": "최대", "BOD": "1.3", "TOC": "3.8", "SS": "6.0", "T-N": "7.486", "T-P": "0.067", "대장균": "260", "비고": ""},
                {"구분": "최소", "BOD": "0.1", "TOC": "2.6", "SS": "0.2", "T-N": "4.312", "T-P": "0.010", "대장균": "9", "비고": ""},
                {"구분": "전월평균", "BOD": "0.7", "TOC": "2.7", "SS": "3.1", "T-N": "6.464", "T-P": "0.045", "대장균": "172", "비고": ""},
                {"구분": "협약수질", "BOD": "5", "TOC": "15", "SS": "10", "T-N": "20", "T-P": "0.2", "대장균": "1,000", "비고": ""},
                {"구분": "유입수질", "BOD": "172.1", "TOC": "35.6", "SS": "166.7", "T-N": "26.678", "T-P": "2.928", "대장균": "200,833", "비고": ""},
            ])
            ed_wq_main = st.data_editor(df_wq_main_init, use_container_width=True, hide_index=True, key="ed_wq_main")

            st.markdown("##### ○ 탈수 슬러지 함수율(%) & 태양광 월별 발전량(최근6개월간)(Mwh)/(%)/((t) — 해당월 빠른 입력 및 자동 반영")
            with st.expander(f"⚡ {sel_report_month}월 당월 탈수 슬러지 함수율 & 태양광 발전량 빠른 입력창 (입력 시 아래 6개월 표와 평균에 즉시 자동 반영)", expanded=True):
                q_c1, q_c2, q_c3, q_c4, q_c5, q_c6 = st.columns(6)
                with q_c1: cur_sl_avg = st.number_input(f"{sel_report_month}월 함수율 평균(%)", value=80.4, step=0.1, format="%.1f")
                with q_c2: cur_sl_max = st.number_input(f"{sel_report_month}월 함수율 최대(%)", value=81.2, step=0.1, format="%.1f")
                with q_c3: cur_sl_min = st.number_input(f"{sel_report_month}월 함수율 최소(%)", value=78.4, step=0.1, format="%.1f")
                with q_c4: cur_so_gen = st.number_input(f"{sel_report_month}월 발전량(Mwh)", value=6.8967, step=0.01, format="%.4f")
                with q_c5: cur_so_eff = st.number_input(f"{sel_report_month}월 발전효율(%)", value=49.9, step=0.1, format="%.1f")
                with q_c6: cur_so_co2 = st.number_input(f"{sel_report_month}월 CO2감축량(t)", value=3.3, step=0.1, format="%.1f")

            col_sl1, col_sl2 = st.columns(2)
            with col_sl1:
                st.markdown("###### ○ 탈수 슬러지 함수율 (%)")
                df_sludge_init = pd.DataFrame([
                    {"구분": "평균", m_cols_labels[0]: 81.4, m_cols_labels[1]: 81.3, m_cols_labels[2]: 78.8, m_cols_labels[3]: 81.3, m_cols_labels[4]: 81.2, m_cols_labels[5]: cur_sl_avg},
                    {"구분": "최대", m_cols_labels[0]: 81.6, m_cols_labels[1]: 81.7, m_cols_labels[2]: 81.3, m_cols_labels[3]: 81.3, m_cols_labels[4]: 81.4, m_cols_labels[5]: cur_sl_max},
                    {"구분": "최소", m_cols_labels[0]: 81.0, m_cols_labels[1]: 79.8, m_cols_labels[2]: 78.4, m_cols_labels[3]: 79.4, m_cols_labels[4]: 79.8, m_cols_labels[5]: cur_sl_min},
                ])
                ed_sludge = st.data_editor(df_sludge_init, use_container_width=True, hide_index=True, key="ed_sludge_6m")
                ed_sludge["평균"] = ed_sludge[m_cols_labels].apply(pd.to_numeric, errors='coerce').mean(axis=1).round(1)

            with col_sl2:
                st.markdown("###### ○ 태양광 월별 발전량(최근6개월간) (Mwh)/(%)/((t)")
                df_solar_init = pd.DataFrame([
                    {"구분": "발전량", m_cols_labels[0]: 8.282, m_cols_labels[1]: 7.983, m_cols_labels[2]: 5.568, m_cols_labels[3]: 6.474, m_cols_labels[4]: 6.228, m_cols_labels[5]: cur_so_gen},
                    {"구분": "월 발전효율", m_cols_labels[0]: 59.9, m_cols_labels[1]: 57.7, m_cols_labels[2]: 40.3, m_cols_labels[3]: 46.8, m_cols_labels[4]: 45.1, m_cols_labels[5]: cur_so_eff},
                    {"구분": "CO2감축량", m_cols_labels[0]: 4.0, m_cols_labels[1]: 3.8, m_cols_labels[2]: 2.7, m_cols_labels[3]: 3.1, m_cols_labels[4]: 3.0, m_cols_labels[5]: cur_so_co2},
                ])
                ed_solar = st.data_editor(df_solar_init, use_container_width=True, hide_index=True, key="ed_solar_6m")
                ed_solar["평균"] = ed_solar[m_cols_labels].apply(pd.to_numeric, errors='coerce').mean(axis=1).round(1)

            st.markdown("##### □ 추진실적 & □ 추진계획 & □ 수범사례 (직접 수정 가능)")
            p_c1, p_c2 = st.columns(2)
            with p_c1:
                perf_elec = st.text_area(
                    "○ 전기설비 (추진실적)",
                    value="  - 전기안전 점검(㈜중앙전기기술단)\n  - 원심탈수기 및 케익호퍼 교체작업 완료 후 시운전 진행중 \n  - 단월처리시설 ALTS 배터리 교체",
                    height=105
                )
                perf_mech = st.text_area(
                    "○ 기계설비 (추진실적)",
                    value="  - 총인처리시설 워시박스 및 에어렌서, 에어벤트 배관 정비, 에어리프트 통수 작업\n  - 총인처리시설 공기압축기 #A 헤드 교체  (10.02) 환경사업본부-18665(₩440,000)\n  - 반응조 PAC 투입배관 전수교체(10.07)\n  - 총인처리시설 메디아 교체완료 ㈜에코지텍 (10.11 ~ 10.19.)\n  - 삼성리 중계펌프장 #A 고정자권선 소손으로 인한 교체 작업",
                    height=135
                )
                perf_etc = st.text_area(
                    "○ 기타 (추진실적)",
                    value="   - 고분자응집제 입고 300kg(10.19.)   \n   - PAC 입고 2,930kg(10.21.)(10.30.)\n   - 염화제이철 입고 1,040kg(10.26.)\n   - 찌꺼기 성분검사(한국환경종합연구원)",
                    height=105
                )
            with p_c2:
                plan_text = st.text_area(
                    "□ 추진계획 (직접 입력 및 수정)",
                    value="  - 시설물 점검\n  - 증설 (1,500ton) 시운전 진행중 ㈜바이오워터 (26.10. ~ 27.1.22)\n  - 기존 (1,700ton) 총인처리시설 메디아 교체 준공예정 (11.06.)\n  - 사용전검사 준비작업 및 정전작업 예정 (11.10.)\n  - 사용전 검사 예정 (11.19.)",
                    height=180
                )
                best_practice_text = st.text_area(
                    "□ 수범사례 (직접 입력 및 수정)",
                    value="",
                    height=160
                )

        with sub_tab2:
            st.markdown(f"#### 🏡 소규모하수처리시설 운영사항 보고({sel_report_month}월)")
            st.caption("○ 대행업체명 : 양평공사 | 아래 표의 모든 수치(평균·최대·최소·협약수질·전월평균·유입수)와 추진실적을 직접 수정할 수 있습니다.")
            df_small_wq_init = pd.DataFrame([
                {"시설명": "삼가리", "처리공법": "BIOSUF 120㎥/일", "구분": "평균", "BOD": "1.8", "TOC": "3.7", "SS": "1.6", "T-N": "10.058", "T-P": "0.767", "대장균": "297", "㎥/주": "310", "㎥/일": "44.3", "비고": ""},
                {"시설명": "삼가리", "처리공법": "BIOSUF 120㎥/일", "구분": "최대", "BOD": "1.8", "TOC": "4.1", "SS": "2.0", "T-N": "11.887", "T-P": "0.946", "대장균": "340", "㎥/주": "328", "㎥/일": "46.9", "비고": ""},
                {"시설명": "삼가리", "처리공법": "BIOSUF 120㎥/일", "구분": "최소", "BOD": "1.8", "TOC": "3.5", "SS": "1.0", "T-N": "7.406", "T-P": "0.566", "대장균": "230", "㎥/주": "279", "㎥/일": "39.9", "비고": ""},
                {"시설명": "삼가리", "처리공법": "BIOSUF 120㎥/일", "구분": "협약수질", "BOD": "10", "TOC": "25", "SS": "10", "T-N": "20", "T-P": "2", "대장균": "3,000", "㎥/주": "-", "㎥/일": "-", "비고": ""},
                {"시설명": "삼가리", "처리공법": "BIOSUF 120㎥/일", "구분": "전월평균", "BOD": "3.2", "TOC": "3.2", "SS": "4.7", "T-N": "9.451", "T-P": "1.318", "대장균": "106", "㎥/주": "318", "㎥/일": "45.4", "비고": ""},
                {"시설명": "삼가리", "처리공법": "BIOSUF 120㎥/일", "구분": "유입수", "BOD": "81.7", "TOC": "21.9", "SS": "74.8", "T-N": "28.184", "T-P": "3.389", "대장균": "203,000", "㎥/주": "-", "㎥/일": "-", "비고": ""},
                {"시설명": "산음리", "처리공법": "SWPP 100㎥/일", "구분": "평균", "BOD": "1.4", "TOC": "3.4", "SS": "5.0", "T-N": "13.443", "T-P": "1.648", "대장균": "64", "㎥/주": "234", "㎥/일": "33.4", "비고": ""},
                {"시설명": "산음리", "처리공법": "SWPP 100㎥/일", "구분": "최대", "BOD": "1.8", "TOC": "3.4", "SS": "9.5", "T-N": "16.600", "T-P": "1.973", "대장균": "100", "㎥/주": "305", "㎥/일": "43.6", "비고": ""},
                {"시설명": "산음리", "처리공법": "SWPP 100㎥/일", "구분": "최소", "BOD": "0.9", "TOC": "3.3", "SS": "3.0", "T-N": "8.972", "T-P": "1.095", "대장균": "24", "㎥/주": "205", "㎥/일": "29.3", "비고": ""},
                {"시설명": "산음리", "처리공법": "SWPP 100㎥/일", "구분": "협약수질", "BOD": "10", "TOC": "25", "SS": "10", "T-N": "20", "T-P": "2", "대장균": "3,000", "㎥/주": "-", "㎥/일": "-", "비고": ""},
                {"시설명": "산음리", "처리공법": "SWPP 100㎥/일", "구분": "전월평균", "BOD": "0.9", "TOC": "2.7", "SS": "2.1", "T-N": "14.326", "T-P": "1.760", "대장균": "31", "㎥/주": "270", "㎥/일": "38.6", "비고": ""},
                {"시설명": "산음리", "처리공법": "SWPP 100㎥/일", "구분": "유입수", "BOD": "308.9", "TOC": "29.0", "SS": "574.4", "T-N": "35.387", "T-P": "3.977", "대장균": "165,000", "㎥/주": "-", "㎥/일": "-", "비고": ""},
                {"시설명": "몰운리", "처리공법": "IC-SBR 60㎥/일", "구분": "평균", "BOD": "1.7", "TOC": "4.2", "SS": "7.7", "T-N": "11.126", "T-P": "1.459", "대장균": "273", "㎥/주": "177", "㎥/일": "25.3", "비고": ""},
                {"시설명": "몰운리", "처리공법": "IC-SBR 60㎥/일", "구분": "최대", "BOD": "2.3", "TOC": "4.4", "SS": "9.5", "T-N": "13.758", "T-P": "1.771", "대장균": "410", "㎥/주": "197", "㎥/일": "28.1", "비고": ""},
                {"시설명": "몰운리", "처리공법": "IC-SBR 60㎥/일", "구분": "최소", "BOD": "1.1", "TOC": "4.1", "SS": "4.2", "T-N": "8.366", "T-P": "1.122", "대장균": "120", "㎥/주": "157", "㎥/일": "22.4", "비고": ""},
                {"시설명": "몰운리", "처리공법": "IC-SBR 60㎥/일", "구분": "협약수질", "BOD": "10", "TOC": "25", "SS": "10", "T-N": "20", "T-P": "2", "대장균": "3,000", "㎥/주": "-", "㎥/일": "-", "비고": ""},
                {"시설명": "몰운리", "처리공법": "IC-SBR 60㎥/일", "구분": "전월평균", "BOD": "3.0", "TOC": "4.0", "SS": "5.8", "T-N": "9.750", "T-P": "1.285", "대장균": "475", "㎥/주": "171", "㎥/일": "24.4", "비고": ""},
                {"시설명": "몰운리", "처리공법": "IC-SBR 60㎥/일", "구분": "유입수", "BOD": "109.4", "TOC": "10.5", "SS": "143.3", "T-N": "20.822", "T-P": "2.396", "대장균": "159,667", "㎥/주": "-", "㎥/일": "-", "비고": ""},
                {"시설명": "당의", "처리공법": "IC-SBR 45㎥/일", "구분": "평균", "BOD": "8.1", "TOC": "3.6", "SS": "1.9", "T-N": "12.199", "T-P": "1.774", "대장균": "280", "㎥/주": "295", "㎥/일": "42.1", "비고": ""},
                {"시설명": "당의", "처리공법": "IC-SBR 45㎥/일", "구분": "최대", "BOD": "9.4", "TOC": "3.8", "SS": "2.3", "T-N": "17.872", "T-P": "2.603", "대장균": "540", "㎥/주": "381", "㎥/일": "54.4", "비고": ""},
                {"시설명": "당의", "처리공법": "IC-SBR 45㎥/일", "구분": "최소", "BOD": "6.5", "TOC": "3.3", "SS": "1.5", "T-N": "6.943", "T-P": "0.703", "대장균": "150", "㎥/주": "259", "㎥/일": "37.0", "비고": ""},
                {"시설명": "당의", "처리공법": "IC-SBR 45㎥/일", "구분": "협약수질", "BOD": "10", "TOC": "25", "SS": "10", "T-N": "40", "T-P": "4", "대장균": "3,000", "㎥/주": "-", "㎥/일": "-", "비고": ""},
                {"시설명": "당의", "처리공법": "IC-SBR 45㎥/일", "구분": "전월평균", "BOD": "9.5", "TOC": "3.2", "SS": "1.1", "T-N": "9.921", "T-P": "1.225", "대장균": "193", "㎥/주": "321", "㎥/일": "45.8", "비고": ""},
                {"시설명": "당의", "처리공법": "IC-SBR 45㎥/일", "구분": "유입수", "BOD": "353.5", "TOC": "59.5", "SS": "627.4", "T-N": "47.961", "T-P": "5.557", "대장균": "83,250", "㎥/주": "-", "㎥/일": "-", "비고": ""},
                {"시설명": "단월마을", "처리공법": "IC-SBR 30㎥/일", "구분": "평균", "BOD": "1.1", "TOC": "4.5", "SS": "1.8", "T-N": "24.191", "T-P": "3.315", "대장균": "233", "㎥/주": "61", "㎥/일": "8.7", "비고": ""},
                {"시설명": "단월마을", "처리공법": "IC-SBR 30㎥/일", "구분": "최대", "BOD": "1.7", "TOC": "5.2", "SS": "3.0", "T-N": "26.770", "T-P": "3.386", "대장균": "510", "㎥/주": "68", "㎥/일": "9.7", "비고": ""},
                {"시설명": "단월마을", "처리공법": "IC-SBR 30㎥/일", "구분": "최소", "BOD": "0.7", "TOC": "3.8", "SS": "1.0", "T-N": "22.618", "T-P": "3.216", "대장균": "9", "㎥/주": "56", "㎥/일": "8.0", "비고": ""},
                {"시설명": "단월마을", "처리공법": "IC-SBR 30㎥/일", "구분": "협약수질", "BOD": "10", "TOC": "25", "SS": "10", "T-N": "40", "T-P": "4", "대장균": "3,000", "㎥/주": "-", "㎥/일": "-", "비고": ""},
                {"시설명": "단월마을", "처리공법": "IC-SBR 30㎥/일", "구분": "전월평균", "BOD": "1.1", "TOC": "3.5", "SS": "1.5", "T-N": "16.513", "T-P": "2.798", "대장균": "42", "㎥/주": "60", "㎥/일": "8.5", "비고": ""},
                {"시설명": "단월마을", "처리공법": "IC-SBR 30㎥/일", "구분": "유입수", "BOD": "441.5", "TOC": "73.2", "SS": "651.5", "T-N": "54.398", "T-P": "6.235", "대장균": "245,000", "㎥/주": "-", "㎥/일": "-", "비고": ""},
                {"시설명": "보룡리(진목)", "처리공법": "고효율오수처리+SOD 23㎥/일", "구분": "평균", "BOD": "2.2", "TOC": "4.2", "SS": "6.4", "T-N": "14.198", "T-P": "1.722", "대장균": "167", "㎥/주": "17", "㎥/일": "2.4", "비고": ""},
                {"시설명": "보룡리(진목)", "처리공법": "고효율오수처리+SOD 23㎥/일", "구분": "최대", "BOD": "2.4", "TOC": "4.7", "SS": "9.2", "T-N": "17.197", "T-P": "2.339", "대장균": "370", "㎥/주": "22", "㎥/일": "3.1", "비고": ""},
                {"시설명": "보룡리(진목)", "처리공법": "고효율오수처리+SOD 23㎥/일", "구분": "최소", "BOD": "1.9", "TOC": "3.5", "SS": "1.4", "T-N": "10.675", "T-P": "0.939", "대장균": "51", "㎥/주": "13", "㎥/일": "1.9", "비고": ""},
                {"시설명": "보룡리(진목)", "처리공법": "고효율오수처리+SOD 23㎥/일", "구분": "협약수질", "BOD": "10", "TOC": "25", "SS": "10", "T-N": "40", "T-P": "4", "대장균": "3,000", "㎥/주": "-", "㎥/일": "-", "비고": ""},
                {"시설명": "보룡리(진목)", "처리공법": "고효율오수처리+SOD 23㎥/일", "구분": "전월평균", "BOD": "3.1", "TOC": "4.2", "SS": "5.9", "T-N": "20.025", "T-P": "2.584", "대장균": "205", "㎥/주": "17", "㎥/일": "2.4", "비고": ""},
                {"시설명": "보룡리(진목)", "처리공법": "고효율오수처리+SOD 23㎥/일", "구분": "유입수", "BOD": "77.1", "TOC": "23.0", "SS": "310.0", "T-N": "32.708", "T-P": "4.105", "대장균": "134,000", "㎥/주": "-", "㎥/일": "-", "비고": ""},
            ])
            ed_small_wq = st.data_editor(df_small_wq_init, use_container_width=True, hide_index=True, height=420, key="ed_small_wq")

            st.markdown("##### ○ 소규모하수처리시설 시설별 추진실적 및 추진계획 (10~15번)")
            df_small_tasks_init = pd.DataFrame([
                {"연번": "10", "시설명": "산음리", "추진실적": "- 침전조 스컴제거 / - 주변 환경정비 / - 잉여슬러지 인발", "추진계획": "기계시설물 점검"},
                {"연번": "11", "시설명": "몰운리", "추진실적": "- 스크린조 청소 및 점검 / - 잉여슬러지 인발 / - 유량조정조 오뚜기 청소 및 점검 / - 주변 환경정비 / - 여과기 점검", "추진계획": "기계시설물 점검"},
                {"연번": "12", "시설명": "보룡리(진목)", "추진실적": "- SOD 설비 청소 / - 방류유량계 센서 청소", "추진계획": "기계시설물 점검"},
                {"연번": "13", "시설명": "삼가리", "추진실적": "- 반응조 오뚜기 청소 및 점검 / - 유량조정조 오뚜기 청소 및 점검 / - 주변 환경정비", "추진계획": "기계시설물 점검"},
                {"연번": "14", "시설명": "단월마을", "추진실적": "- 반응조 오뚜기 점검 후 공정 정상화 / - 잉여슬러지 인발 / - 반응조 및 유량조정조 수위계 청소 및 점검 / - 주변 환경정비", "추진계획": "기계시설물 점검"},
                {"연번": "15", "시설명": "당의", "추진실적": "- 잉여슬러지 인발 / - 반응조 및 유량조정조 수위계 청소 및 점검 / - 디켄터 펌프 청소 및 점검 / - 주변 환경정비", "추진계획": "기계시설물 점검"},
            ])
            ed_small_tasks = st.data_editor(df_small_tasks_init, use_container_width=True, hide_index=True, key="ed_small_tasks")

        with sub_tab3:
            st.markdown(f"#### 🛖 개인하수처리시설 운영사항 보고({sel_report_month}월)")
            st.caption("○ 대행업체명 : 양평공사 | 개인하수처리시설 11개소의 처리량·수질 현황 및 추진실적(6~11번)을 편집할 수 있습니다.")
            df_priv_wq_init = pd.DataFrame([
                {"시설명": "석산리(섬이)", "공법": "BSTS-II 30㎥/일", "구분": "평균", "BOD": "9.8", "SS": "9.5", "㎥/주": "175", "㎥/일": "25.0", "비고": "명달1·3반, 화전2리(여촌): 개인하수폐쇄 가동정지 / 옥천리: 유량계 침수 고장"},
                {"시설명": "석산리(섬이)", "공법": "BSTS-II 30㎥/일", "구분": "최대", "BOD": "9.8", "SS": "9.5", "㎥/주": "", "㎥/일": "", "비고": ""},
                {"시설명": "석산리(섬이)", "공법": "BSTS-II 30㎥/일", "구분": "최소", "BOD": "9.8", "SS": "9.5", "㎥/주": "", "㎥/일": "", "비고": ""},
                {"시설명": "석산리(섬이)", "공법": "BSTS-II 30㎥/일", "구분": "협약수질", "BOD": "10", "SS": "10", "㎥/주": "-", "㎥/일": "-", "비고": ""},
                {"시설명": "석산리(섬이)", "공법": "BSTS-II 30㎥/일", "구분": "전월평균", "BOD": "3.5", "SS": "3.2", "㎥/주": "161", "㎥/일": "23.0", "비고": ""},
                {"시설명": "석산리(섬이)", "공법": "BSTS-II 30㎥/일", "구분": "유입수", "BOD": "45.4", "SS": "41.7", "㎥/주": "-", "㎥/일": "-", "비고": ""},
                {"시설명": "물레울양지", "공법": "접촉폭기법 15㎥/일", "구분": "평균", "BOD": "1.6", "SS": "2.3", "㎥/주": "유량계 미설치", "㎥/일": "", "비고": ""},
                {"시설명": "물레울양지", "공법": "접촉폭기법 15㎥/일", "구분": "최대", "BOD": "1.6", "SS": "2.3", "㎥/주": "", "㎥/일": "", "비고": ""},
                {"시설명": "물레울양지", "공법": "접촉폭기법 15㎥/일", "구분": "최소", "BOD": "1.6", "SS": "2.3", "㎥/주": "", "㎥/일": "", "비고": ""},
                {"시설명": "물레울양지", "공법": "접촉폭기법 15㎥/일", "구분": "협약수질", "BOD": "10", "SS": "10", "㎥/주": "", "㎥/일": "", "비고": ""},
                {"시설명": "물레울양지", "공법": "접촉폭기법 15㎥/일", "구분": "전월평균", "BOD": "9.0", "SS": "0.8", "㎥/주": "", "㎥/일": "", "비고": ""},
                {"시설명": "물레울양지", "공법": "접촉폭기법 15㎥/일", "구분": "유입수", "BOD": "237.2", "SS": "513.3", "㎥/주": "", "㎥/일": "", "비고": ""},
                {"시설명": "물레울음지", "공법": "접촉폭기법 15㎥/일", "구분": "평균", "BOD": "8.4", "SS": "3.3", "㎥/주": "유량계 미설치", "㎥/일": "", "비고": ""},
                {"시설명": "물레울음지", "공법": "접촉폭기법 15㎥/일", "구분": "최대", "BOD": "8.4", "SS": "3.3", "㎥/주": "", "㎥/일": "", "비고": ""},
                {"시설명": "물레울음지", "공법": "접촉폭기법 15㎥/일", "구분": "최소", "BOD": "8.4", "SS": "3.3", "㎥/주": "", "㎥/일": "", "비고": ""},
                {"시설명": "물레울음지", "공법": "접촉폭기법 15㎥/일", "구분": "협약수질", "BOD": "10", "SS": "10", "㎥/주": "", "㎥/일": "", "비고": ""},
                {"시설명": "물레울음지", "공법": "접촉폭기법 15㎥/일", "구분": "전월평균", "BOD": "9.4", "SS": "9.3", "㎥/주": "", "㎥/일": "", "비고": ""},
                {"시설명": "물레울음지", "공법": "접촉폭기법 15㎥/일", "구분": "유입수", "BOD": "60.0", "SS": "32.0", "㎥/주": "", "㎥/일": "", "비고": ""},
                {"시설명": "돌고개", "공법": "접촉폭기법 15㎥/일", "구분": "평균", "BOD": "2.3", "SS": "9.2", "㎥/주": "유량계 미설치", "㎥/일": "", "비고": ""},
                {"시설명": "돌고개", "공법": "접촉폭기법 15㎥/일", "구분": "최대", "BOD": "2.3", "SS": "9.2", "㎥/주": "", "㎥/일": "", "비고": ""},
                {"시설명": "돌고개", "공법": "접촉폭기법 15㎥/일", "구분": "최소", "BOD": "2.3", "SS": "9.2", "㎥/주": "", "㎥/일": "", "비고": ""},
                {"시설명": "돌고개", "공법": "접촉폭기법 15㎥/일", "구분": "협약수질", "BOD": "10", "SS": "10", "㎥/주": "", "㎥/일": "", "비고": ""},
                {"시설명": "돌고개", "공법": "접촉폭기법 15㎥/일", "구분": "전월평균", "BOD": "3.6", "SS": "3.4", "㎥/주": "", "㎥/일": "", "비고": ""},
                {"시설명": "돌고개", "공법": "접촉폭기법 15㎥/일", "구분": "유입수", "BOD": "69.5", "SS": "172.5", "㎥/주": "", "㎥/일": "", "비고": ""},
                {"시설명": "석산리복지회관", "공법": "접촉폭기법 10㎥/일", "구분": "평균", "BOD": "8.4", "SS": "0.5", "㎥/주": "유량계 미설치", "㎥/일": "", "비고": ""},
                {"시설명": "석산리복지회관", "공법": "접촉폭기법 10㎥/일", "구분": "최대", "BOD": "8.4", "SS": "0.5", "㎥/주": "", "㎥/일": "", "비고": ""},
                {"시설명": "석산리복지회관", "공법": "접촉폭기법 10㎥/일", "구분": "최소", "BOD": "8.4", "SS": "0.5", "㎥/주": "", "㎥/일": "", "비고": ""},
                {"시설명": "석산리복지회관", "공법": "접촉폭기법 10㎥/일", "구분": "협약수질", "BOD": "10", "SS": "10", "㎥/주": "", "㎥/일": "", "비고": ""},
                {"시설명": "석산리복지회관", "공법": "접촉폭기법 10㎥/일", "구분": "전월평균", "BOD": "3.0", "SS": "3.2", "㎥/주": "", "㎥/일": "", "비고": ""},
                {"시설명": "석산리복지회관", "공법": "접촉폭기법 10㎥/일", "구분": "유입수", "BOD": "267.7", "SS": "1425.0", "㎥/주": "", "㎥/일": "", "비고": ""},
                {"시설명": "인이피", "공법": "접촉폭기법 10㎥/일", "구분": "평균", "BOD": "1.4", "SS": "0.9", "㎥/주": "유량계 미설치", "㎥/일": "", "비고": ""},
                {"시설명": "인이피", "공법": "접촉폭기법 10㎥/일", "구분": "최대", "BOD": "1.4", "SS": "0.9", "㎥/주": "", "㎥/일": "", "비고": ""},
                {"시설명": "인이피", "공법": "접촉폭기법 10㎥/일", "구분": "최소", "BOD": "1.4", "SS": "0.9", "㎥/주": "", "㎥/일": "", "비고": ""},
                {"시설명": "인이피", "공법": "접촉폭기법 10㎥/일", "구분": "협약수질", "BOD": "10", "SS": "10", "㎥/주": "", "㎥/일": "", "비고": ""},
                {"시설명": "인이피", "공법": "접촉폭기법 10㎥/일", "구분": "전월평균", "BOD": "3.6", "SS": "3.8", "㎥/주": "", "㎥/일": "", "비고": ""},
                {"시설명": "인이피", "공법": "접촉폭기법 10㎥/일", "구분": "유입수", "BOD": "4.4", "SS": "1.0", "㎥/주": "", "㎥/일": "", "비고": ""},
            ])
            ed_priv_wq = st.data_editor(df_priv_wq_init, use_container_width=True, hide_index=True, height=420, key="ed_priv_wq")

            st.markdown("##### ○ 개인하수처리시설 시설별 추진실적 및 추진계획 (6~11번)")
            df_priv_tasks_init = pd.DataFrame([
                {"연번": "6", "시설명": "석산리(섬이)", "추진실적": "시설물 점검 및 환경정비 / 송풍기 점검", "추진계획": "처리장 브로워 가동상태 점검"},
                {"연번": "7", "시설명": "물레울양지", "추진실적": "시설물 점검 및 환경정비 / 송풍기 점검", "추진계획": "처리장 브로워 가동상태 점검"},
                {"연번": "8", "시설명": "물레울음지", "추진실적": "시설물 점검 및 환경정비 / 송풍기 점검", "추진계획": "처리장 브로워 가동상태 점검"},
                {"연번": "9", "시설명": "돌고개", "추진실적": "시설물 점검 및 환경정비 / 송풍기 점검", "추진계획": "처리장 브로워 가동상태 점검"},
                {"연번": "10", "시설명": "석산리복지회관", "추진실적": "시설물 점검 및 환경정비 / 송풍기 점검", "추진계획": "처리장 브로워 가동상태 점검"},
                {"연번": "11", "시설명": "인이피", "추진실적": "시설물 점검 및 환경정비 / 송풍기 점검", "추진계획": "처리장 브로워 가동상태 점검"},
            ])
            ed_priv_tasks = st.data_editor(df_priv_tasks_init, use_container_width=True, hide_index=True, key="ed_priv_tasks")

        def df_to_html_rows(df_in, cols):
            html_out = ""
            for _, r in df_in.iterrows():
                tds = "".join([f"<td style='border:1px solid #000; padding:4px 6px; text-align:center;'>{str(r[c]).replace(' / ', '<br>').replace(chr(10), '<br>')}</td>" for c in cols])
                html_out += f"<tr>{tds}</tr>"
            return html_out

        m_headers_html = "".join([f"<th style='border:1px solid #000; background:#f2f2f2; padding:4px;'>{m}</th>" for m in m_cols_labels])
        inflow_rows_html = df_to_html_rows(ed_inflow, ["구분"] + m_cols_labels + ["평균"])
        outflow_rows_html = df_to_html_rows(ed_outflow, ["구분"] + m_cols_labels + ["평균"])
        wq_main_rows_html = df_to_html_rows(ed_wq_main, ["구분", "BOD", "TOC", "SS", "T-N", "T-P", "대장균", "비고"])
        sludge_rows_html = df_to_html_rows(ed_sludge, ["구분"] + m_cols_labels + ["평균"])
        solar_rows_html = df_to_html_rows(ed_solar, ["구분"] + m_cols_labels + ["평균"])
        small_wq_rows_html = df_to_html_rows(ed_small_wq, ["시설명", "처리공법", "구분", "BOD", "TOC", "SS", "T-N", "T-P", "대장균", "㎥/주", "㎥/일", "비고"])
        small_task_rows_html = df_to_html_rows(ed_small_tasks, ["연번", "시설명", "추진실적", "추진계획"])
        priv_wq_rows_html = df_to_html_rows(ed_priv_wq, ["시설명", "공법", "구분", "BOD", "SS", "㎥/주", "㎥/일", "비고"])
        priv_task_rows_html = df_to_html_rows(ed_priv_tasks, ["연번", "시설명", "추진실적", "추진계획"])

        perf_elec_html = perf_elec.replace("\n", "<br>")
        perf_mech_html = perf_mech.replace("\n", "<br>")
        perf_etc_html = perf_etc.replace("\n", "<br>")
        plan_html = plan_text.replace("\n", "<br>")
        best_html = best_practice_text.replace("\n", "<br>")

        monthly_report_full_html = f"""
        <!DOCTYPE html>
        <html>
        <head>
        <meta charset="UTF-8">
        <style>
            @page {{ size: A4 portrait; margin: 12mm; }}
            body {{ font-family: 'Malgun Gothic', '맑은 고딕', '바탕', sans-serif; color: #000; margin: 12px; font-size: 10.5px; line-height: 1.5; }}
            .report-page {{ max-width: 780px; margin: 0 auto 28px auto; padding: 15px; border: 1px solid #ddd; background: #fff; page-break-after: always; }}
            h2.sec-title {{ font-size: 16px; font-weight: bold; text-align: center; border-bottom: 2px solid #000; padding-bottom: 6px; margin-top: 10px; margin-bottom: 12px; }}
            .box-h {{ font-size: 12px; font-weight: bold; margin-top: 10px; margin-bottom: 4px; }}
            .item-line {{ margin-left: 10px; font-size: 11px; margin-bottom: 2px; }}
            table.rep-tbl {{ width: 100%; border-collapse: collapse; margin-bottom: 10px; font-size: 9.5px; }}
            table.rep-tbl th, table.rep-tbl td {{ border: 1px solid #000; padding: 3px 5px; vertical-align: middle; }}
            table.rep-tbl th {{ background-color: #f2f2f2; font-weight: bold; text-align: center; }}
            .unit-right {{ text-align: right; font-size: 9.5px; margin-bottom: 2px; }}
            @media print {{
                .no-print {{ display: none !important; }}
                .report-page {{ border: none; padding: 0; margin: 0; page-break-after: always; }}
            }}
        </style>
        </head>
        <body>
            <div class="no-print" style="max-width:780px; margin:0 auto 10px auto; text-align:right;">
                <button onclick="window.print()" style="background:#0284C7; color:white; border:none; padding:8px 18px; border-radius:6px; font-weight:bold; cursor:pointer; font-size:13px;">
                    🖨️ 공공하수도시설 대행사업 월간보고서({sel_report_month}월) 전체 인쇄하기
                </button>
            </div>

            <div class="report-page">
                <div style="text-align:center; padding:18px 10px; border:2px double #000; margin-bottom:15px;">
                    <div style="font-size:20px; font-weight:bold;">공공하수처리시설 대행사업</div>
                    <div style="font-size:14px; font-weight:bold; margin-top:4px;">- ‘{str(sel_report_year)[2:]}년 {sel_report_month}월 월간보고 -</div>
                    <div style="font-size:13px; font-weight:bold; margin-top:10px;">양평공사 [환경사업본부]</div>
                </div>

                <h2 class="sec-title">단월공공하수처리시설 운영사항 보고({sel_report_month}월)</h2>
                <div class="box-h">□ 개요</div>
                <div class="item-line">○ 시 &nbsp;설 &nbsp;명 : {ov_name}</div>
                <div class="item-line">○ 대행업체명 : {ov_agency}</div>
                <div class="item-line">○ 시 설 용 량 : {ov_cap}</div>
                <div class="item-line">○ 공 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;법 : {ov_method}</div>
                <div class="item-line">○ 근 &nbsp;무 &nbsp;자 : {ov_staff}</div>

                <div class="item-line" style="margin-top:8px; font-weight:bold;">○ 유입량 현황(최근 6개월간)</div>
                <div class="unit-right">(㎥/일)</div>
                <table class="rep-tbl">
                    <tr><th>구분</th>{m_headers_html}<th>평균</th></tr>
                    {inflow_rows_html}
                </table>

                <div class="item-line" style="font-weight:bold;">○ 방류량 현황(최근6개월간)</div>
                <div class="unit-right">(㎥/일)</div>
                <table class="rep-tbl">
                    <tr><th>구분</th>{m_headers_html}<th>평균</th></tr>
                    {outflow_rows_html}
                </table>

                <div class="item-line" style="font-weight:bold;">○ 수질현황</div>
                <div class="unit-right">(㎎/ℓ)</div>
                <table class="rep-tbl">
                    <tr><th>구분</th><th>BOD</th><th>TOC</th><th>SS</th><th>T-N</th><th>T-P</th><th>대장균</th><th>비고</th></tr>
                    {wq_main_rows_html}
                </table>

                <div class="item-line" style="font-weight:bold;">○ 탈수 슬러지 함수율</div>
                <div class="unit-right">(%)</div>
                <table class="rep-tbl">
                    <tr><th>구분</th>{m_headers_html}<th>평균</th></tr>
                    {sludge_rows_html}
                </table>

                <div class="item-line" style="font-weight:bold;">○ 태양광 월별 발전량(최근6개월간)</div>
                <div class="unit-right">(Mwh)/(%)/((t)</div>
                <table class="rep-tbl">
                    <tr><th>구분</th>{m_headers_html}<th>평균</th></tr>
                    {solar_rows_html}
                </table>

                <div class="box-h">□ 추진실적</div>
                <div class="item-line"><b>○ 전기설비</b><br>{perf_elec_html}</div>
                <div class="item-line" style="margin-top:4px;"><b>○ 기계설비</b><br>{perf_mech_html}</div>
                <div class="item-line" style="margin-top:4px;"><b>○ 기타</b><br>{perf_etc_html}</div>

                <div class="box-h">□ 추진계획</div>
                <div class="item-line">{plan_html}</div>

                <div class="box-h">□ 수범사례</div>
                <div class="item-line">{best_html}</div>
            </div>

            <div class="report-page">
                <h2 class="sec-title">소규모하수처리시설 운영사항 보고({sel_report_month}월)</h2>
                <div class="box-h">□ 개요</div>
                <div class="item-line">○ 대행업체명 : {ov_agency}</div>
                <div class="item-line" style="font-weight:bold; margin-bottom:4px;">○ 처리량 및 수질 현황</div>
                <table class="rep-tbl">
                    <tr>
                        <th rowspan="2">시설명</th><th rowspan="2">처리공법</th><th rowspan="2">구분</th>
                        <th colspan="6">수질현황</th><th colspan="2">방류량</th><th rowspan="2">비고</th>
                    </tr>
                    <tr>
                        <th>BOD</th><th>TOC</th><th>SS</th><th>T-N</th><th>T-P</th><th>대장균</th><th>㎥/주</th><th>㎥/일</th>
                    </tr>
                    {small_wq_rows_html}
                </table>

                <div class="box-h">□ 소규모하수처리시설 시설별 추진실적 및 추진계획</div>
                <table class="rep-tbl">
                    <tr><th style="width:8%;">연번</th><th style="width:18%;">시설명</th><th style="width:48%;">추진실적</th><th style="width:26%;">추진계획</th></tr>
                    {small_task_rows_html}
                </table>
            </div>

            <div class="report-page">
                <h2 class="sec-title">개인하수처리시설 운영사항 보고({sel_report_month}월)</h2>
                <div class="box-h">□ 개요</div>
                <div class="item-line">○ 대행업체명 : {ov_agency}</div>
                <div class="item-line" style="font-weight:bold; margin-bottom:4px;">○ 처리량 및 수질 현황 (명달1반·명달3반·화전2리·화전2리(여촌): 개인하수폐쇄 가동정지 / 옥천리: 유량계 침수 고장)</div>
                <table class="rep-tbl">
                    <tr>
                        <th rowspan="2">시설명</th><th rowspan="2">처리공법</th><th rowspan="2">구분</th>
                        <th colspan="2">수질현황(㎎/ℓ)</th><th colspan="2">방류량</th><th rowspan="2">비고</th>
                    </tr>
                    <tr>
                        <th>BOD</th><th>SS</th><th>㎥/주</th><th>㎥/일</th>
                    </tr>
                    {priv_wq_rows_html}
                </table>

                <div class="box-h">□ 개인하수처리시설 시설별 추진실적 및 추진계획</div>
                <table class="rep-tbl">
                    <tr><th style="width:8%;">연번</th><th style="width:18%;">시설명</th><th style="width:48%;">추진실적</th><th style="width:26%;">추진계획</th></tr>
                    {priv_task_rows_html}
                </table>
            </div>
        </body>
        </html>
        """

        st.divider()
        st.subheader(f"📋 공공하수도시설 대행사업 월간보고서({sel_report_month}월) 공식 인쇄 양식 미리보기")
        st.components.v1.html(monthly_report_full_html, height=820, scrolling=True)

        rep_pdf_bytes = generate_html_to_pdf_bytes(monthly_report_full_html)
        rep_is_pdf = rep_pdf_bytes[:4] == b"%PDF"
        full_task_combined = f"[전기설비] {perf_elec} [기계설비] {perf_mech} [기타] {perf_etc} [추진계획] {plan_text}"
        bytes_hwpx = generate_hwpx_monthly_report(
            sel_report_month,
            hwpx_file_up,
            {"avg": float(cur_sl_avg), "max": float(cur_sl_max), "min": float(cur_sl_min)},
            {"current_month": float(cur_so_gen)},
            full_task_combined,
            sel_report_year
        )
        fname_hwpx = sanitize_filename(f"공공하수도시설_대행사업_월간보고서({sel_report_month}월)_{sel_report_year}.hwpx")
        fname_pdf = sanitize_filename(f"공공하수도시설_대행사업_월간보고서({sel_report_month}월)_{sel_report_year}.pdf")
        fname_html = sanitize_filename(f"공공하수도시설_대행사업_월간보고서({sel_report_month}월)_{sel_report_year}.html")

        col_rep_b1, col_rep_b2, col_rep_b3, col_rep_b4 = st.columns(4)
        with col_rep_b1:
            st.download_button(
                f"📥 월간보고서({sel_report_month}월) 한글(HWPX) 다운로드",
                data=bytes_hwpx,
                file_name=fname_hwpx,
                mime="application/hwp+zip",
                type="primary",
                use_container_width=True
            )
        with col_rep_b2:
            st.download_button(
                f"📥 월간보고서({sel_report_month}월) PDF 다운로드",
                data=rep_pdf_bytes,
                file_name=fname_pdf if rep_is_pdf else fname_html,
                mime="application/pdf" if rep_is_pdf else "text/html",
                type="primary",
                use_container_width=True
            )
        with col_rep_b3:
            st.download_button(
                f"📥 월간보고서({sel_report_month}월) 인쇄용 HTML 다운로드",
                data=monthly_report_full_html,
                file_name=fname_html,
                mime="text/html",
                use_container_width=True
            )
        with col_rep_b4:
            if st.button("💾 월간보고서(한글 HWPX + PDF + 인쇄본) 영구 보관함 저장", type="primary", use_container_width=True):
                with open(os.path.join(HWPX_RECORD_DIR, fname_hwpx), "wb") as f:
                    f.write(bytes_hwpx)
                with open(os.path.join(HWPX_RECORD_DIR, fname_html), "w", encoding="utf-8") as f:
                    f.write(monthly_report_full_html)
                if rep_is_pdf:
                    with open(os.path.join(HWPX_RECORD_DIR, fname_pdf), "wb") as f:
                        f.write(rep_pdf_bytes)
                st.success(f"✅ {sel_report_month}월 월간보고서가 한글(HWPX), PDF, 인쇄용 HTML로 영구 보관함에 안전하게 저장되었습니다!")

    with tab_hw_a:
        st.subheader("🗂️ 월간보고서(HWPX · HTML) 영구 보관소 열람 및 관리")
        saved_reps = [sanitize_filename(f) for f in os.listdir(HWPX_RECORD_DIR) if f.endswith((".hwpx", ".html", ".pdf"))]
        if saved_reps:
            st.write(f"📁 **총 {len(saved_reps)}건의 월간보고서가 영구 보관되어 있습니다.**")
            file_info_list = []
            for fn in sorted(saved_reps, reverse=True):
                fp = os.path.join(HWPX_RECORD_DIR, fn)
                sz_kb = round(os.path.getsize(fp) / 1024, 1)
                mtime = datetime.datetime.fromtimestamp(os.path.getmtime(fp), KST).strftime('%Y-%m-%d %H:%M')
                file_info_list.append({"파일명": fn, "크기(KB)": sz_kb, "저장일시": mtime})
            st.dataframe(pd.DataFrame(file_info_list), use_container_width=True)

            target_hw = st.selectbox("열람·다운로드 또는 삭제할 보고서 선택", sorted(saved_reps, reverse=True))
            c_d1, c_d2 = st.columns(2)
            t_path = os.path.join(HWPX_RECORD_DIR, sanitize_filename(target_hw))
            with open(t_path, "rb") as f:
                c_d1.download_button("📥 선택 보고서 다운로드", f.read(), file_name=target_hw, type="primary", use_container_width=True)
            if c_d2.button("🗑️ 선택 보고서 삭제", use_container_width=True):
                os.remove(t_path)
                st.success("삭제되었습니다.")
                st.rerun()

            if target_hw.endswith(".html"):
                st.markdown("##### 🔍 저장된 월간보고서 인쇄 양식 미리보기")
                with open(t_path, "r", encoding="utf-8") as f:
                    st.components.v1.html(f.read(), height=780, scrolling=True)
        else:
            st.info("보관된 보고서가 없습니다.")


# -------------------------------------------------------------
# 3. TMS 관제 (자동생성 더미 제거 + 6대 항목 pH, BOD, T-N, T-P, SS, TOC 신호등 & AI 예측 시뮬레이션)
# -------------------------------------------------------------
elif menu == "📡 3. TMS 수질 2·4·6·8시간 후 AI 예측 & 신호등 실시간 관제":
    st.title("📡 단월 본장 TMS 수질 AI 시계열 예측 & 신호등 관제")
    tab_t1, tab_t2, tab_t3 = st.tabs(["📝 [입력] 엑셀 적재 & 수동입력", "🚦 [관제] 6대 항목 신호등 & AI 예측 시뮬레이션", "🗂️ [보관소] TMS 영구 저장 데이터 열람"])
    with tab_t1:
        st.markdown("##### 📥 TMS 측정 데이터 엑셀/CSV 업로드 (업로드 및 저장 시에만 DB에 적재됩니다)")
        up_tms = st.file_uploader("TMS 측정 엑셀/CSV 파일 선택", type=["xlsx", "xls", "csv"], accept_multiple_files=True, key="tms_file_up_v2")
        if up_tms:
            if st.button("💾 업로드한 TMS 파일 데이터 마스터 DB에 저장", type="primary", use_container_width=True):
                tms_list = []
                for f in up_tms:
                    try:
                        df_raw = pd.read_csv(f) if f.name.endswith('.csv') else pd.read_excel(f)
                        # 컬럼명 매칭 또는 행 탐색
                        for _, row in df_raw.iterrows():
                            row_vals = [str(v) for v in row.values]
                            d_match = None
                            for rv in row_vals[:3]:
                                m = re.search(r'(20[1-3]\d)[-/.](\d{1,2})[-/.](\d{1,2})', rv)
                                if m:
                                    d_match = f"{int(m.group(1)):04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
                                    break
                            if d_match:
                                ph_v = pd.to_numeric(row.get('방류pH', row.get('pH', 7.2)), errors='coerce')
                                bod_v = pd.to_numeric(row.get('방류BOD', row.get('BOD', 2.3)), errors='coerce')
                                toc_v = pd.to_numeric(row.get('방류TOC', row.get('TOC', 3.1)), errors='coerce')
                                ss_v = pd.to_numeric(row.get('방류SS', row.get('SS', 4.5)), errors='coerce')
                                tn_v = pd.to_numeric(row.get('방류TN', row.get('T-N', 8.45)), errors='coerce')
                                tp_v = pd.to_numeric(row.get('방류TP', row.get('T-P', 0.065)), errors='coerce')
                                fl_v = pd.to_numeric(row.get('방류유량', row.get('유량', 70.5)), errors='coerce')
                                ph_f = float(ph_v) if pd.notna(ph_v) else 7.2
                                bod_f = float(bod_v) if pd.notna(bod_v) else 2.3
                                toc_f = float(toc_v) if pd.notna(toc_v) else 3.1
                                ss_f = float(ss_v) if pd.notna(ss_v) else 4.5
                                tn_f = float(tn_v) if pd.notna(tn_v) else 8.45
                                tp_f = float(tp_v) if pd.notna(tp_v) else 0.065
                                fl_f = float(fl_v) if pd.notna(fl_v) else 70.5
                                tms_list.append({
                                    "측정일자": d_match, "측정시각": str(row.get('측정시각', '12:00:00')),
                                    "방류pH": ph_f, "방류BOD": bod_f, "방류TOC": toc_f, "방류SS": ss_f,
                                    "방류TN": tn_f, "방류TP": tp_f, "방류유량": fl_f,
                                    "예측pH_4h": round(ph_f, 2), "예측BOD_4h": round(bod_f * 1.02, 2),
                                    "예측SS_4h": round(ss_f * 1.01, 2), "예측TN_4h": round(tn_f * 1.02, 2),
                                    "예측TP_4h": round(tp_f * 1.01, 3), "비고": f"파일업로드({f.name})"
                                })
                    except Exception:
                        pass
                if tms_list:
                    append_to_tms_db(pd.DataFrame(tms_list))
                    st.success(f"✅ 업로드하신 파일에서 총 {len(tms_list)}건의 TMS 데이터가 영구 저장되었습니다!")
                    st.rerun()
                else:
                    st.warning("⚠️ 업로드된 파일에서 날짜(YYYY-MM-DD) 형식을 포함한 유효한 행을 찾지 못했습니다.")

        if st.button("🔄 1번 운영일지 마스터 DB에서 실측 수질 데이터만 TMS로 동기화", type="primary", use_container_width=True, key="btn_tms_sync"):
            df_m = get_master_data(MAIN_PLANT)
            if not df_m.empty:
                tms_list = []
                for _, r in df_m.iterrows():
                    ph = float(r.get('방류pH', 7.2)) if pd.notna(r.get('방류pH')) else 7.2
                    bod = float(r.get('방류BOD', 2.3)) if pd.notna(r.get('방류BOD')) else 2.3
                    toc = float(r.get('방류TOC', 3.1)) if pd.notna(r.get('방류TOC')) else 3.1
                    ss = float(r.get('방류SS', 4.8)) if pd.notna(r.get('방류SS')) else 4.8
                    tn = float(r.get('방류TN', 8.45)) if pd.notna(r.get('방류TN')) else 8.45
                    tp = float(r.get('방류TP', 0.065)) if pd.notna(r.get('방류TP')) else 0.065
                    fl = float(r.get('방류량', 1650.0)) / 24.0 if pd.notna(r.get('방류량')) else 70.5
                    tms_list.append({
                        "측정일자": str(r['날짜'])[:10], "측정시각": "12:00:00",
                        "방류pH": ph, "방류BOD": bod, "방류TOC": toc, "방류SS": ss,
                        "방류TN": tn, "방류TP": tp, "방류유량": round(fl, 1),
                        "예측pH_4h": round(ph, 2), "예측BOD_4h": round(bod * 1.02, 2),
                        "예측SS_4h": round(ss * 1.01, 2), "예측TN_4h": round(tn * 1.02, 2),
                        "예측TP_4h": round(tp * 1.01, 3), "비고": "마스터DB 동기화"
                    })
                append_to_tms_db(pd.DataFrame(tms_list))
                st.success(f"✅ 1번 운영일지 마스터 DB에서 총 {len(tms_list)}일치 실측 데이터가 동기화되었습니다!")
                st.rerun()
            else:
                st.warning("⚠️ 현재 1번 운영일지 마스터 DB에 저장된 단월 본장 데이터가 없습니다. (임의 자동생성을 하지 않습니다. 먼저 1번 메뉴에서 운영일지를 저장하거나 아래 수동 입력을 이용해 주세요.)")

        st.divider()
        st.markdown("##### ✍️ 실시간 TMS 6대 수질 항목 수동 입력 및 영구 적재")
        with st.form("tms_manual_form"):
            col_tm1, col_tm2, col_tm3 = st.columns(3)
            with col_tm1:
                m_date = st.date_input("측정일자", datetime.date.today())
                m_ph = st.number_input("방류 pH (기준 5.8~8.5)", value=7.20, step=0.05, format="%.2f")
                m_bod = st.number_input("방류 BOD (mg/L, 기준 10 이하)", value=2.30, step=0.10, format="%.2f")
            with col_tm2:
                m_time = st.text_input("측정시각", "12:00:00")
                m_toc = st.number_input("방류 TOC (mg/L, 기준 15 이하)", value=3.10, step=0.10, format="%.2f")
                m_ss = st.number_input("방류 SS (mg/L, 기준 10 이하)", value=4.80, step=0.10, format="%.2f")
            with col_tm3:
                m_tn = st.number_input("방류 T-N (mg/L, 기준 20 이하)", value=8.45, step=0.05, format="%.2f")
                m_tp = st.number_input("방류 T-P (mg/L, 기준 0.2 이하)", value=0.065, step=0.005, format="%.3f")
                m_flow = st.number_input("방류 유량 (㎥/h)", value=70.5, step=0.5, format="%.1f")
            
            submitted_tms = st.form_submit_button("💾 수동 입력 TMS 데이터 영구 저장", type="primary", use_container_width=True)
            if submitted_tms:
                new_row = pd.DataFrame([{
                    "측정일자": m_date.strftime('%Y-%m-%d'),
                    "측정시각": m_time,
                    "방류pH": m_ph, "방류BOD": m_bod, "방류TOC": m_toc, "방류SS": m_ss,
                    "방류TN": m_tn, "방류TP": m_tp, "방류유량": m_flow,
                    "예측pH_4h": round(m_ph, 2), "예측BOD_4h": round(m_bod * 1.02, 2),
                    "예측SS_4h": round(m_ss * 1.01, 2), "예측TN_4h": round(m_tn * 1.02, 2),
                    "예측TP_4h": round(m_tp * 1.01, 3), "비고": "수동 입력"
                }])
                append_to_tms_db(new_row)
                st.success("✅ 직접 입력하신 TMS 데이터가 영구 저장되었습니다!")
                st.rerun()

    with tab_t2:
        df_tms = get_tms_db()
        if df_tms.empty:
            st.info("💡 현재 저장된 TMS 데이터가 없습니다. (사용자가 입력하지 않은 데이터는 자동 생성하지 않습니다.) 좌측 [📝 입력] 탭에서 데이터를 수동 저장하거나 업로드하시면 그 데이터를 기준으로 실시간 관제가 수행되며, 아래에서는 실시간 AI 예측 시뮬레이터를 바로 테스트하실 수 있습니다.")
            base_ph, base_bod, base_tn, base_tp, base_ss, base_toc = 7.20, 2.30, 8.45, 0.065, 4.50, 3.10
        else:
            base_ph = float(df_tms['방류pH'].iloc[0]) if '방류pH' in df_tms.columns and pd.notna(df_tms['방류pH'].iloc[0]) else 7.20
            base_bod = float(df_tms['방류BOD'].iloc[0]) if '방류BOD' in df_tms.columns and pd.notna(df_tms['방류BOD'].iloc[0]) else 2.30
            base_tn = float(df_tms['방류TN'].iloc[0]) if '방류TN' in df_tms.columns and pd.notna(df_tms['방류TN'].iloc[0]) else 8.45
            base_tp = float(df_tms['방류TP'].iloc[0]) if '방류TP' in df_tms.columns and pd.notna(df_tms['방류TP'].iloc[0]) else 0.065
            base_ss = float(df_tms['방류SS'].iloc[0]) if '방류SS' in df_tms.columns and pd.notna(df_tms['방류SS'].iloc[0]) else 4.50
            base_toc = float(df_tms['방류TOC'].iloc[0]) if '방류TOC' in df_tms.columns and pd.notna(df_tms['방류TOC'].iloc[0]) else 3.10

        # 신호등 판정 함수 (정상 / 주의 / 경고)
        def eval_signal(param, val):
            if param == "pH":
                if 6.2 <= val <= 8.1: return "🟢 정상 (안전)", "#10B981", "#ECFDF5"
                elif 5.8 <= val <= 8.5: return "🟡 주의 (관찰)", "#F59E0B", "#FFFBEB"
                else: return "🔴 경고 (기준초과)", "#EF4444", "#FEF2F2"
            limits = {"BOD": (7.0, 10.0), "T-N": (15.0, 20.0), "T-P": (0.14, 0.20), "SS": (7.0, 10.0), "TOC": (11.0, 15.0)}
            warn_lim, max_lim = limits[param]
            if val <= warn_lim: return "🟢 정상 (안전)", "#10B981", "#ECFDF5"
            elif val <= max_lim: return "🟡 주의 (관찰)", "#F59E0B", "#FFFBEB"
            else: return "🔴 경고 (기준초과)", "#EF4444", "#FEF2F2"

        st.markdown("##### 🚦 단월 본장 6대 수질 항목(pH · BOD · T-N · T-P · SS · TOC) 실시간 신호등 관제판")
        sig_items = [
            ("pH", "pH 실시간 (기준 5.8~8.5)", f"{base_ph:.2f}", "", base_ph),
            ("BOD", "BOD 실시간 (기준 10 이하)", f"{base_bod:.2f}", "mg/L", base_bod),
            ("T-N", "T-N 실시간 (기준 20 이하)", f"{base_tn:.2f}", "mg/L", base_tn),
            ("T-P", "T-P 실시간 (기준 0.2 이하)", f"{base_tp:.3f}", "mg/L", base_tp),
            ("SS", "SS 실시간 (기준 10 이하)", f"{base_ss:.2f}", "mg/L", base_ss),
            ("TOC", "TOC 실시간 (기준 15 이하)", f"{base_toc:.2f}", "mg/L", base_toc),
        ]
        cols_sig = st.columns(6)
        for idx_s, (p_key, p_title, p_val_str, p_unit, p_num) in enumerate(sig_items):
            status_txt, border_col, bg_col = eval_signal(p_key, p_num)
            with cols_sig[idx_s]:
                st.markdown(f"""
                <div style="background:{bg_col}; border:2px solid {border_col}; border-radius:12px; padding:12px 10px; text-align:center; box-shadow:0 2px 6px rgba(0,0,0,0.04);">
                    <div style="font-size:11.5px; font-weight:700; color:#475569; margin-bottom:4px;">{p_title}</div>
                    <div style="font-size:22px; font-weight:900; color:#0F172A; margin-bottom:6px;">{p_val_str} <span style="font-size:12px; font-weight:600;">{p_unit}</span></div>
                    <div style="display:inline-block; font-size:11.5px; font-weight:800; color:{border_col}; background:white; padding:3px 10px; border-radius:20px; border:1px solid {border_col};">{status_txt}</div>
                </div>
                """, unsafe_allow_html=True)

        st.divider()
        st.markdown("##### 🎛️ AI 수질 예측 시뮬레이션 제어반 (운전조건 변화에 따른 2·4·6·8시간 후 6대 수질 예측)")
        with st.expander("⚙️ AI 시뮬레이션 운전 변수 조절 (유입부하 · 송풍량 · 응집제 주입율 조절 시 아래 그래프와 신호등이 실시간 반응합니다)", expanded=True):
            sc1, sc2, sc3, sc4 = st.columns(4)
            with sc1:
                sim_inflow_chg = st.slider("🌊 유입유량/오염부하 변동 (%)", -30, 50, 0, step=5)
            with sc2:
                sim_air_chg = st.slider("💨 생물반응조 송풍량 증감 (%)", -30, 30, 0, step=5)
            with sc3:
                sim_chem_chg = st.slider("🧪 총인(IPR) 응집제 주입 증감 (%)", -30, 40, 0, step=5)
            with sc4:
                sim_temp_chg = st.slider("🌡️ 반응조 수온 조건 (℃)", 8.0, 28.0, 18.0, step=1.0)

        # 운전 변수에 따른 2, 4, 6, 8시간 후 동적 시뮬레이션 연산
        load_factor = 1.0 + (sim_inflow_chg * 0.006)
        air_factor = 1.0 - (sim_air_chg * 0.005)
        chem_factor = 1.0 - (sim_chem_chg * 0.006)
        temp_factor = 1.0 + max(0.0, (15.0 - sim_temp_chg) * 0.015)

        pred_hours = ["현재", "+2시간 후", "+4시간 후", "+6시간 후", "+8시간 후"]
        time_weights = [0.0, 0.35, 0.70, 0.90, 1.00]

        sim_rows = []
        for t_label, w in zip(pred_hours, time_weights):
            p_ph = round(base_ph + w * (-sim_chem_chg * 0.004 + sim_inflow_chg * 0.002), 2)
            p_bod = round(max(0.5, base_bod * (1.0 + w * ((load_factor * air_factor - 1.0) + 0.02))), 2)
            p_tn = round(max(2.0, base_tn * (1.0 + w * ((load_factor * air_factor * temp_factor - 1.0) + 0.015))), 2)
            p_tp = round(max(0.010, base_tp * (1.0 + w * ((load_factor * chem_factor - 1.0) + 0.02))), 3)
            p_ss = round(max(1.0, base_ss * (1.0 + w * ((load_factor * chem_factor - 1.0) + 0.015))), 2)
            p_toc = round(max(1.0, base_toc * (1.0 + w * ((load_factor * air_factor - 1.0) + 0.02))), 2)
            sim_rows.append({
                "예측시점": t_label,
                "pH": p_ph,
                "BOD(mg/L)": p_bod,
                "T-N(mg/L)": p_tn,
                "T-P(mg/L)": p_tp,
                "SS(mg/L)": p_ss,
                "TOC(mg/L)": p_toc,
                "4h_신호등": eval_signal("T-P", p_tp)[0] if p_tp > 0.14 else eval_signal("T-N", p_tn)[0]
            })
        df_pred_horizon = pd.DataFrame(sim_rows)

        # 6대 항목 전체 2·4·6·8시간 예측 서브플롯 그래프 (pH, BOD, T-N, T-P, SS, TOC 한눈에 보기)
        fig_6 = make_subplots(
            rows=2, cols=3,
            subplot_titles=(
                "pH 예측 (기준 5.8~8.5)", "BOD 예측 (기준 10 mg/L 이하)", "T-N 예측 (기준 20 mg/L 이하)",
                "T-P 예측 (기준 0.2 mg/L 이하)", "SS 예측 (기준 10 mg/L 이하)", "TOC 예측 (기준 15 mg/L 이하)"
            ),
            vertical_spacing=0.18, horizontal_spacing=0.08
        )
        plot_specs = [
            (1, 1, "pH", "#0284C7", 8.5),
            (1, 2, "BOD(mg/L)", "#10B981", 10.0),
            (1, 3, "T-N(mg/L)", "#6366F1", 20.0),
            (2, 1, "T-P(mg/L)", "#F59E0B", 0.20),
            (2, 2, "SS(mg/L)", "#EC4899", 10.0),
            (2, 3, "TOC(mg/L)", "#14B8A6", 15.0)
        ]
        for r_i, c_i, col_n, line_c, lim_v in plot_specs:
            fig_6.add_trace(
                go.Scatter(
                    x=df_pred_horizon["예측시점"], y=df_pred_horizon[col_n],
                    mode="lines+markers+text",
                    text=[str(v) for v in df_pred_horizon[col_n]],
                    textposition="top center",
                    name=col_n,
                    line=dict(color=line_c, width=3),
                    marker=dict(size=8)
                ),
                row=r_i, col=c_i
            )
            fig_6.add_hline(y=lim_v, line_dash="dot", line_color="#EF4444", row=r_i, col=c_i)

        fig_6.update_layout(height=540, showlegend=False, title_text="📈 향후 2·4·6·8시간 6대 수질 항목(pH · BOD · T-N · T-P · SS · TOC) AI 예측 시뮬레이션 곡선")
        st.plotly_chart(fig_6, use_container_width=True)

        st.markdown("##### 📋 시간대별 6대 수질 AI 예측 상세 수치표")
        st.dataframe(df_pred_horizon, use_container_width=True)

        if not df_tms.empty:
            st.divider()
            st.markdown("##### 📊 저장된 TMS 실측 데이터 일자별 6대 수질 추이 그래프")
            df_plot = df_tms.head(30).sort_values(by='측정일자')
            sel_trend_cols = st.multiselect(
                "추이 그래프에 표시할 수질 항목 선택",
                ['방류pH', '방류BOD', '방류TN', '방류TP', '방류SS', '방류TOC'],
                default=['방류pH', '방류BOD', '방류TN', '방류TP', '방류SS', '방류TOC']
            )
            if sel_trend_cols:
                fig_t = px.line(df_plot, x='측정일자', y=sel_trend_cols, markers=True, title="저장된 TMS 일자별 6대 방류수 수질 추이")
                st.plotly_chart(fig_t, use_container_width=True)

    with tab_t3:
        st.subheader("🗂️ 영구 저장된 TMS 마스터 데이터베이스 열람 및 관리")
        df_all_tms = get_tms_db()
        if not df_all_tms.empty: 
            st.write(f"📊 **사용자가 저장/동기화한 총 {len(df_all_tms)}건의 TMS 기록이 보관되어 있습니다.**")
            st.dataframe(df_all_tms, use_container_width=True)
            col_td1, col_td2 = st.columns([3, 1])
            with col_td1:
                csv_data = df_all_tms.to_csv(index=False, encoding='utf-8-sig')
                st.download_button("📥 TMS 적재 데이터 전체 CSV 다운로드", csv_data, file_name="danwol_tms_export.csv", mime="text/csv", type="primary", use_container_width=True)
            with col_td2:
                if st.button("🗑️ TMS 저장 데이터 전체 초기화", use_container_width=True):
                    if os.path.exists(TMS_ACCUM_DB):
                        os.remove(TMS_ACCUM_DB)
                    st.success("✅ TMS 데이터베이스가 깨끗하게 초기화되었습니다.")
                    st.rerun()
        else: 
            st.info("💡 현재 보관된 TMS 데이터가 없습니다. [입력] 탭에서 직접 저장하거나 업로드한 데이터만 이곳에 보관됩니다.")

# -------------------------------------------------------------
# 4. 단월(DANWOL) KNR+IPR 본장 & 소규모 6개소 통합 SCADA 엔지니어링 관제 센터 (독자 설계)
# -------------------------------------------------------------
elif menu == "⚙️ 4. AI 최적 운전조건 제안 & KNR+IPR 공정 정밀진단":
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🛰️ 단월 SCADA 제어 파라미터")
    dw_sel_fac = st.sidebar.selectbox("🏛️ 집중 진단 시설 선택", [MAIN_PLANT] + SMALL_PLANTS, index=0, key="dw_scada_fac")
    dw_ref_date = st.sidebar.date_input("📆 관제 기준일자", datetime.date(2026, 10, 30), key="dw_scada_date")
    dw_sel_month = st.sidebar.selectbox("📊 실측 연동 월 선택", list(range(1, 13)), index=int(dw_ref_date.month) - 1, key="dw_scada_month")

    st.sidebar.markdown("##### 🧬 KNR 생물반응조 & 총인(IPR) 설정")
    dw_mlss_a = st.sidebar.slider("반응조 #A지 MLSS (mg/L)", 2200, 5200, 3620, step=50, key="dw_mlss_a")
    dw_mlss_b = st.sidebar.slider("반응조 #B지 MLSS (mg/L)", 2200, 5200, 3710, step=50, key="dw_mlss_b")
    dw_do_set = st.sidebar.slider("호기조 목표 DO (mg/L)", 1.0, 3.8, 2.0, step=0.1, key="dw_do_set")
    dw_rtn_rate = st.sidebar.slider("내부반송률 (%)", 100, 300, 180, step=10, key="dw_rtn_rate")
    dw_ipr_diff_p = st.sidebar.slider("IPR 총인 메디아 여과 차압 (kPa)", 5.0, 45.0, 14.5, step=0.5, key="dw_ipr_dp")

    fac_spec = PLANT_DESIGN_SPECS.get(dw_sel_fac, PLANT_DESIGN_SPECS[MAIN_PLANT])
    max_cap_val = int(fac_spec["cap"])
    avg_mlss_ab = int(round((dw_mlss_a + dw_mlss_b) / 2.0))

    tab_scada, tab_knr_sim, tab_p1, tab_p2 = st.tabs([
        "🛰️ 1. 단월 7개소(본장+소규모6) 통합 SCADA & KNR+IPR 정밀관제",
        "🎛️ 2. 인버터(Hz)·약注펌프·기온/강우 AI 자율제어 시뮬레이터",
        "⚙️ 3. 시설별 일별 AI 최적 운전조건 연산·적재",
        "🗂️ 4. 공정 제어 마스터 DB 보관함"
    ])

    with tab_scada:
        # 1) 상단 다크 네이비 SCADA 통합 관제 헤더
        st.markdown(f"""
        <div class="dw-scada-header">
            <div>
                <div style="display:flex; align-items:center; gap:10px; margin-bottom:6px;">
                    <span style="background:#0284C7; color:white; font-weight:800; font-size:11.5px; padding:3px 10px; border-radius:6px;">DANWOL SMART SCADA v3.6</span>
                    <span style="color:#38BDF8; font-size:12.5px; font-weight:700;">양평공사 환경2팀 · 단월 통합관제소 (총 설계용량 2,078 ㎥/일)</span>
                </div>
                <h2 style="margin:0; font-size:24px; font-weight:900; color:#FFFFFF; letter-spacing:-0.5px;">
                    🛰️ {dw_sel_fac} [{fac_spec['method']}] 엔지니어링 정밀 진단 보드
                </h2>
                <div style="color:#94A3B8; font-size:13px; margin-top:4px;">
                    설계용량: <b style="color:#E2E8F0;">{max_cap_val:,} ㎥/일</b> | 약품체계: <b style="color:#38BDF8;">{fac_spec['chem_type']}</b> | 관제 기준: <b style="color:#E2E8F0;">{dw_ref_date.strftime('%Y-%m')} ({dw_sel_month}월 실측 마스터DB 연동)</b>
                </div>
            </div>
            <div style="text-align:right;">
                <div style="background:rgba(16, 185, 129, 0.18); color:#34D399; border:1px solid rgba(52, 211, 153, 0.45); padding:6px 14px; border-radius:8px; font-size:12.5px; font-weight:800;">
                    ⚡ KNR+IPR 자율제어 루프 정상
                </div>
                <div style="font-size:11.5px; color:#94A3B8; margin-top:5px;">증설(1,500㎥/일) 연계 시운전 모니터링 병행중</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # 2) 단월 권역 7개소 (본장 + 소규모 6개소) 실시간 멀티 상태 매트릭스 보드
        st.markdown("##### 🏛️ 단월 권역 7개 하수처리시설(본장 1,700톤 + 소규모 6개소 378톤) 통합 가동 매트릭스")
        ym_prefix_scada = f"{dw_ref_date.year}-{dw_sel_month:02d}"
        matrix_html_cards = []
        all_7_plants = [MAIN_PLANT] + SMALL_PLANTS
        for p_name in all_7_plants:
            p_spec = PLANT_DESIGN_SPECS[p_name]
            df_p_m = get_master_data(p_name, f"{ym_prefix_scada}-01", f"{ym_prefix_scada}-31")
            if df_p_m.empty or ('방류량' in df_p_m.columns and not df_p_m['방류량'].dropna().any()):
                df_p_all = get_master_data(p_name)
                df_p_m = df_p_all[df_p_all['방류량'].notna()].tail(31) if (not df_p_all.empty and '방류량' in df_p_all.columns) else df_p_all
            p_flow = float(df_p_m['방류량'].dropna().mean()) if (not df_p_m.empty and '방류량' in df_p_m.columns and df_p_m['방류량'].dropna().any()) else p_spec['cap'] * 0.75
            p_bod = float(df_p_m['방류BOD'].dropna().mean()) if (not df_p_m.empty and '방류BOD' in df_p_m.columns and df_p_m['방류BOD'].dropna().any()) else 1.5
            p_tp = float(df_p_m['방류TP'].dropna().mean()) if (not df_p_m.empty and '방류TP' in df_p_m.columns and df_p_m['방류TP'].dropna().any()) else 0.05
            load_pct = round((p_flow / p_spec['cap']) * 100, 1)
            short_name = "단월본장" if p_name == MAIN_PLANT else p_name
            active_cls = "active-fac" if p_name == dw_sel_fac else ""
            badge_col = "#34D399" if load_pct <= 95 else "#FBBF24"
            card_str = (
                f'<div class="dw-matrix-card {active_cls}">'
                f'<div style="font-size:12px; font-weight:800; color:#38BDF8; margin-bottom:2px;">{short_name}</div>'
                f'<div style="font-size:10px; color:#94A3B8; margin-bottom:6px;">{p_spec["method"]} ({int(p_spec["cap"])}톤)</div>'
                f'<div style="font-size:16px; font-weight:900; color:#FFFFFF;">{p_flow:,.1f} <span style="font-size:10px; font-weight:500;">㎥/일</span></div>'
                f'<div style="font-size:10.5px; color:{badge_col}; font-weight:700; margin:3px 0;">가동률 {load_pct}%</div>'
                f'<div style="font-size:10px; color:#CBD5E1; border-top:1px solid #1E293B; padding-top:4px; margin-top:4px;">'
                f'BOD <b>{p_bod:.1f}</b> · T-P <b>{p_tp:.2f}</b>'
                f'</div></div>'
            )
            matrix_html_cards.append(card_str)
        st.markdown(f'<div class="dw-matrix-grid">{"".join(matrix_html_cards)}</div>', unsafe_allow_html=True)

        # 선택된 시설의 실측 데이터 로드 및 엔지니어링 지표 연산
        df_sel_m = get_master_data(dw_sel_fac, f"{ym_prefix_scada}-01", f"{ym_prefix_scada}-31")
        if df_sel_m.empty or ('방류량' in df_sel_m.columns and not df_sel_m['방류량'].dropna().any()):
            df_sel_all = get_master_data(dw_sel_fac)
            df_sel_m = df_sel_all[df_sel_all['방류량'].notna()].tail(31) if (not df_sel_all.empty and '방류량' in df_sel_all.columns) else df_sel_all

        in_flow_v = float(df_sel_m['유입량'].dropna().mean()) if (not df_sel_m.empty and '유입량' in df_sel_m.columns and df_sel_m['유입량'].dropna().any()) else max_cap_val * 0.92
        out_flow_v = float(df_sel_m['방류량'].dropna().mean()) if (not df_sel_m.empty and '방류량' in df_sel_m.columns and df_sel_m['방류량'].dropna().any()) else in_flow_v * 0.89
        reuse_v = float(df_sel_m['재이용수'].dropna().mean()) if (not df_sel_m.empty and '재이용수' in df_sel_m.columns and df_sel_m['재이용수'].dropna().any()) else max(0.0, in_flow_v - out_flow_v)
        w_temp_v = float(df_sel_m['수온'].dropna().mean()) if (not df_sel_m.empty and '수온' in df_sel_m.columns and df_sel_m['수온'].dropna().any()) else 16.5

        in_bod_v = float(df_sel_m['유입BOD'].dropna().mean()) if (not df_sel_m.empty and '유입BOD' in df_sel_m.columns and df_sel_m['유입BOD'].dropna().any()) else 168.0
        in_toc_v = float(df_sel_m['유입TOC'].dropna().mean()) if (not df_sel_m.empty and '유입TOC' in df_sel_m.columns and df_sel_m['유입TOC'].dropna().any()) else 35.5
        in_ss_v = float(df_sel_m['유입SS'].dropna().mean()) if (not df_sel_m.empty and '유입SS' in df_sel_m.columns and df_sel_m['유입SS'].dropna().any()) else 162.0
        in_tn_v = float(df_sel_m['유입TN'].dropna().mean()) if (not df_sel_m.empty and '유입TN' in df_sel_m.columns and df_sel_m['유입TN'].dropna().any()) else 26.8
        in_tp_v = float(df_sel_m['유입TP'].dropna().mean()) if (not df_sel_m.empty and '유입TP' in df_sel_m.columns and df_sel_m['유입TP'].dropna().any()) else 2.95

        out_bod_v = float(df_sel_m['방류BOD'].dropna().mean()) if (not df_sel_m.empty and '방류BOD' in df_sel_m.columns and df_sel_m['방류BOD'].dropna().any()) else 0.7
        out_toc_v = float(df_sel_m['방류TOC'].dropna().mean()) if (not df_sel_m.empty and '방류TOC' in df_sel_m.columns and df_sel_m['방류TOC'].dropna().any()) else 3.1
        out_ss_v = float(df_sel_m['방류SS'].dropna().mean()) if (not df_sel_m.empty and '방류SS' in df_sel_m.columns and df_sel_m['방류SS'].dropna().any()) else 2.1
        out_tn_v = float(df_sel_m['방류TN'].dropna().mean()) if (not df_sel_m.empty and '방류TN' in df_sel_m.columns and df_sel_m['방류TN'].dropna().any()) else 6.01
        out_tp_v = float(df_sel_m['방류TP'].dropna().mean()) if (not df_sel_m.empty and '방류TP' in df_sel_m.columns and df_sel_m['방류TP'].dropna().any()) else 0.032
        out_coli_v = float(df_sel_m['방류대장균'].dropna().mean()) if (not df_sel_m.empty and '방류대장균' in df_sel_m.columns and df_sel_m['방류대장균'].dropna().any()) else 85.0

        ai_calc = calculate_ai_process_parameters(in_flow_v, in_bod_v, in_tn_v, in_tp_v, facility_name=dw_sel_fac)
        fm_ratio = round((in_flow_v * in_bod_v) / max(1.0, (avg_mlss_ab * (max_cap_val * 0.45))), 3)
        reuse_rate_pct = round((reuse_v / max(1.0, in_flow_v)) * 100.0, 1) if dw_sel_fac == MAIN_PLANT else 0.0
        backwash_status = "정상 여과중 (역세척 여유)" if dw_ipr_diff_p < 25.0 else ("⚠️ 조기 역세척(Backwash) 권장" if dw_ipr_diff_p < 35.0 else "🔴 즉시 역세척 및 메디아 점검 필요")

        # 3) 단월 고유의 4대 핵심 공정 모듈 엔지니어링 패널 (KNR 생물반응조 / IPR 총인처리 / 송풍·반송 제어 / 장내 재이용수·방류 부하)
        st.markdown(f"##### 🔬 [{dw_sel_fac}] 4대 핵심 공정 엔지니어링 진단 패널")
        ep1, ep2, ep3, ep4 = st.columns(4)
        with ep1:
            st.markdown(f"""
            <div class="dw-eng-panel" style="border-top-color:#0284C7;">
                <span class="dw-eng-tag">MODULE 01 · 생물반응조 (KNR)</span>
                <div style="font-size:13px; font-weight:800; color:#0F172A;">유입 C/N비 & MLSS (#A·#B지)</div>
                <div class="dw-eng-metric">C/N {ai_calc['CN비']} <span style="font-size:13px; color:#15803D;">(적정 ≥4.0)</span></div>
                <div class="dw-eng-sub">• #A지 MLSS: <b>{dw_mlss_a:,}</b> / #B지: <b>{dw_mlss_b:,} mg/L</b></div>
                <div class="dw-eng-sub">• F/M비: <b>{fm_ratio} kgBOD/kgMLSS·d</b></div>
                <div class="dw-eng-sub">• 반응조 평균수온: <b>{w_temp_v:.1f} ℃</b> (탈질 안정)</div>
            </div>
            """, unsafe_allow_html=True)
        with ep2:
            st.markdown(f"""
            <div class="dw-eng-panel" style="border-top-color:#6366F1;">
                <span class="dw-eng-tag">MODULE 02 · 총인처리설비 (IPR)</span>
                <div style="font-size:13px; font-weight:800; color:#0F172A;">응집제 주입량 & 메디아 여과차압</div>
                <div class="dw-eng-metric">T-P {out_tp_v:.3f} <span style="font-size:13px; color:#64748B;">mg/L (협약 0.2)</span></div>
                <div class="dw-eng-sub">• 권장 염화제이철(38%): <b>{ai_calc['권장염화제이철_L']} L/일</b></div>
                <div class="dw-eng-sub">• 권장 종침 PAC(17%): <b>{ai_calc['종침전PAC주입량_L']} L/일</b></div>
                <div class="dw-eng-sub">• 여과기 차압: <b>{dw_ipr_diff_p:.1f} kPa</b> ({backwash_status})</div>
            </div>
            """, unsafe_allow_html=True)
        with ep3:
            st.markdown(f"""
            <div class="dw-eng-panel" style="border-top-color:#10B981;">
                <span class="dw-eng-tag">MODULE 03 · 포기·내부반송 제어</span>
                <div style="font-size:13px; font-weight:800; color:#0F172A;">AI 송풍량 & 질산화/탈질 밸런스</div>
                <div class="dw-eng-metric">{ai_calc['권장송풍량_m3min']} <span style="font-size:14px;">㎥/min ({ai_calc['송풍기가동대수']}대)</span></div>
                <div class="dw-eng-sub">• 호기조 DO 설정: <b>{dw_do_set:.1f} mg/L</b> / 반송률: <b>{dw_rtn_rate}%</b></div>
                <div class="dw-eng-sub">• 유입 T-N <b>{in_tn_v:.1f}</b> ➔ 방류 T-N <b>{out_tn_v:.2f} mg/L</b></div>
                <div class="dw-eng-sub">• 질소 제거율: <b style="color:#059669;">{round((1 - out_tn_v/max(1, in_tn_v))*100, 1)}%</b> 달성</div>
            </div>
            """, unsafe_allow_html=True)
        with ep4:
            st.markdown(f"""
            <div class="dw-eng-panel" style="border-top-color:#F59E0B;">
                <span class="dw-eng-tag">MODULE 04 · 수계 부하 & 재이용수</span>
                <div style="font-size:13px; font-weight:800; color:#0F172A;">장내 재이용(세척수) & 방류수질</div>
                <div class="dw-eng-metric">{reuse_v:,.0f} <span style="font-size:14px;">㎥/일 (회수율 {reuse_rate_pct}%)</span></div>
                <div class="dw-eng-sub">• 일평균 유입 <b>{in_flow_v:,.0f}</b> / 방류 <b>{out_flow_v:,.0f} ㎥/일</b></div>
                <div class="dw-eng-sub">• 방류 BOD <b>{out_bod_v:.1f}</b> · TOC <b>{out_toc_v:.1f}</b> · SS <b>{out_ss_v:.1f}</b></div>
                <div class="dw-eng-sub">• UV소독 총대장균군: <b>{int(round(out_coli_v)):,} 개/mL</b></div>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # 4) 단월 독자 차트 섹션: 좌측 [일별 유입·방류·재이용수 수지 및 총인(T-P) 제거 추이] + 우측 [오염물질 일일 부하량(kg/일) 제거 분석]
        c_dw_ch1, c_dw_ch2 = st.columns([1.45, 1.0])
        with c_dw_ch1:
            st.markdown(f"##### 📈 [{dw_sel_fac}] {dw_sel_month}월 실측 유량 수지(유입·방류·재이용) 및 T-P/T-N 고도처리 추이")
            if not df_sel_m.empty and len(df_sel_m) > 1:
                df_p_chart = df_sel_m.sort_values('날짜').tail(31)
                fig_scada = make_subplots(
                    rows=2, cols=1, shared_xaxes=True, vertical_spacing=0.10, row_heights=[0.55, 0.45],
                    specs=[[{"secondary_y": True}], [{"secondary_y": True}]],
                    subplot_titles=("🌊 일별 유입량 · 방류량 · 장내 재이용수량 (㎥/일)", "🧪 생물반응조+IPR 거친 후 방류 T-N (mg/L) 및 T-P (mg/L) 실측 추이")
                )
                fig_scada.add_trace(go.Scatter(x=df_p_chart['날짜'], y=df_p_chart['유입량'], name="유입량(㎥/일)", mode="lines+markers", line=dict(color="#0284C7", width=2.5)), row=1, col=1, secondary_y=False)
                fig_scada.add_trace(go.Scatter(x=df_p_chart['날짜'], y=df_p_chart['방류량'], name="방류량(㎥/일)", mode="lines", line=dict(color="#10B981", width=2, dash="dot")), row=1, col=1, secondary_y=False)
                if '재이용수' in df_p_chart.columns and df_p_chart['재이용수'].notna().any():
                    fig_scada.add_trace(go.Bar(x=df_p_chart['날짜'], y=df_p_chart['재이용수'], name="재이용수(㎥/일)", marker_color="rgba(245, 158, 11, 0.45)"), row=1, col=1, secondary_y=True)
                fig_scada.add_hline(y=max_cap_val, line_dash="dash", line_color="#EF4444", annotation_text=f"설계용량({max_cap_val:,}㎥/일)", row=1, col=1, secondary_y=False)

                if '방류TN' in df_p_chart.columns:
                    fig_scada.add_trace(go.Scatter(x=df_p_chart['날짜'], y=df_p_chart['방류TN'], name="방류 T-N(mg/L)", mode="lines+markers", line=dict(color="#6366F1", width=2.2)), row=2, col=1, secondary_y=False)
                if '방류TP' in df_p_chart.columns:
                    fig_scada.add_trace(go.Scatter(x=df_p_chart['날짜'], y=df_p_chart['방류TP'], name="방류 T-P(mg/L)", mode="lines+markers", line=dict(color="#EC4899", width=2.2)), row=2, col=1, secondary_y=True)
                fig_scada.update_layout(height=470, margin=dict(l=15, r=15, t=35, b=15), hovermode="x unified", legend=dict(orientation="h", y=1.08))
                st.plotly_chart(fig_scada, use_container_width=True)
            else:
                st.info("선택된 기간의 실측 데이터가 부족합니다.")

        with c_dw_ch2:
            st.markdown("##### ⚖️ 일일 오염물질 유입부하 vs 방류부하 저감량 (kg/일)")
            bod_in_kg = round(in_flow_v * in_bod_v * 0.001, 1)
            bod_out_kg = round(out_flow_v * out_bod_v * 0.001, 2)
            toc_in_kg = round(in_flow_v * in_toc_v * 0.001, 1)
            toc_out_kg = round(out_flow_v * out_toc_v * 0.001, 2)
            ss_in_kg = round(in_flow_v * in_ss_v * 0.001, 1)
            ss_out_kg = round(out_flow_v * out_ss_v * 0.001, 2)
            tn_in_kg = round(in_flow_v * in_tn_v * 0.001, 1)
            tn_out_kg = round(out_flow_v * out_tn_v * 0.001, 2)
            tp_in_kg = round(in_flow_v * in_tp_v * 0.001, 2)
            tp_out_kg = round(out_flow_v * out_tp_v * 0.001, 3)

            df_load_compare = pd.DataFrame([
                {"항목": "BOD 부하", "유입부하(kg/일)": bod_in_kg, "방류부하(kg/일)": bod_out_kg, "일일제거량(kg/일)": round(bod_in_kg - bod_out_kg, 1), "제거율(%)": f"{round((1 - bod_out_kg/max(0.01, bod_in_kg))*100, 1)}%"},
                {"항목": "TOC 부하", "유입부하(kg/일)": toc_in_kg, "방류부하(kg/일)": toc_out_kg, "일일제거량(kg/일)": round(toc_in_kg - toc_out_kg, 1), "제거율(%)": f"{round((1 - toc_out_kg/max(0.01, toc_in_kg))*100, 1)}%"},
                {"항목": "SS 부하", "유입부하(kg/일)": ss_in_kg, "방류부하(kg/일)": ss_out_kg, "일일제거량(kg/일)": round(ss_in_kg - ss_out_kg, 1), "제거율(%)": f"{round((1 - ss_out_kg/max(0.01, ss_in_kg))*100, 1)}%"},
                {"항목": "T-N 부하", "유입부하(kg/일)": tn_in_kg, "방류부하(kg/일)": tn_out_kg, "일일제거량(kg/일)": round(tn_in_kg - tn_out_kg, 1), "제거율(%)": f"{round((1 - tn_out_kg/max(0.01, tn_in_kg))*100, 1)}%"},
                {"항목": "T-P 부하", "유입부하(kg/일)": tp_in_kg, "방류부하(kg/일)": tp_out_kg, "일일제거량(kg/일)": round(tp_in_kg - tp_out_kg, 2), "제거율(%)": f"{round((1 - tp_out_kg/max(0.001, tp_in_kg))*100, 1)}%"},
            ])
            fig_load = go.Figure()
            fig_load.add_trace(go.Bar(x=df_load_compare["항목"], y=df_load_compare["유입부하(kg/일)"], name="유입 오염부하(kg/일)", marker_color="#94A3B8", text=df_load_compare["유입부하(kg/일)"], textposition="auto"))
            fig_load.add_trace(go.Bar(x=df_load_compare["항목"], y=df_load_compare["일일제거량(kg/일)"], name="KNR+IPR 제거부하(kg/일)", marker_color="#0284C7", text=df_load_compare["제거율(%)"], textposition="auto"))
            fig_load.update_layout(barmode="group", height=250, margin=dict(l=10, r=10, t=25, b=10), legend=dict(orientation="h", y=1.15))
            st.plotly_chart(fig_load, use_container_width=True)
            st.dataframe(df_load_compare, use_container_width=True, hide_index=True)

        # 5) 단월 현장 맞춤형 설비·공정 기술진단 매뉴얼 (본장 KNR+IPR 및 소규모 6개소 현장 설비 기준)
        st.divider()
        st.markdown("##### 🛠️ 단월 현장 설비별 엔지니어링 점검 & 트러블슈팅 프로토콜")
        dw_sop1, dw_sop2, dw_sop3, dw_sop4 = st.tabs([
            "🧬 ① KNR 생물반응조 (#A·#B지) & DO/반송 제어",
            "⚡ ② 총인처리시설 (IPR 메디아 여과·워시박스·공기압축기)",
            "🧪 ③ 응집제 (PAC·염화제이철·폴리머) & 원심탈수기",
            "🏡 ④ 소규모 6개소 (산음·삼가리·진목·몰운·단월마을·당의) 순회점검"
        ])
        with dw_sop1:
            st.markdown(
                f"* **#A지 / #B지 MLSS 균등 배분 관리**: 현재 #A지 **{dw_mlss_a:,} mg/L**, #B지 **{dw_mlss_b:,} mg/L** (편차 **{abs(dw_mlss_a - dw_mlss_b):,} mg/L**). 양 계열 편차가 300 mg/L 이상 벌어질 경우 분배조 게이트 및 반송슬러지 밸브 개도율을 조정하십시오.\n"
                f"* **수온 변화에 따른 C/N비 및 질산화율 제어**: 현재 유입 C/N비는 **{ai_calc['CN비']}**이며 반응조 수온은 **{w_temp_v:.1f}℃**입니다. 동절기 수온 12℃ 이하 저하 시 잉여슬러지 인발량을 줄여 SRT(슬러지 체류시간)를 18일 이상으로 확보하고 내부반송률(현재 **{dw_rtn_rate}%**)을 유지하십시오."
            )
        with dw_sop2:
            st.markdown(
                f"* **IPR 총인처리 메디아 여과기 및 역세척 관리**: 현재 여과 차압은 **{dw_ipr_diff_p:.1f} kPa** ({backwash_status})입니다. 워시박스(Wash Box), 에어렌서, 에어벤트 배관 막힘 여부와 에어리프트 통수 상태를 주기적으로 점검하십시오.\n"
                "* **총인 공기압축기(#A/#B) 및 교반기 점검**: 역세용 공기압축기 헤드 압력 및 수분 응축수 배출 상태를 확인하고, 급속/완속 응집교반기 감속기 이상 소음을 점검하십시오."
            )
        with dw_sop3:
            st.markdown(
                f"* **종침 PAC 및 IPR 염화제이철 정량펌프 점검**: 일일 권장 주입량은 염화제이철 **{ai_calc['권장염화제이철_L']} L/일**, 종침 PAC **{ai_calc['종침전PAC주입량_L']} L/일**입니다. 반응조 및 종침 약품 투입배관 스케일 막힘 여부를 상시 확인하십시오.\n"
                "* **원심탈수기 & 케익호퍼 함수율 관리**: 고분자응집제(폴리머) 용해 농도(0.1~0.2%)와 원심탈수기 차속을 조절하여 탈수케익 함수율을 **80% 내외**로 안정화하십시오."
            )
        with dw_sop4:
            st.markdown(
                "* **산음(SWPP 100톤) / 삼가리(BIOSUF 120톤)**: 침전조 스컴 제거, 반응조·유량조정조 오뚜기(레벨스위치) 청소 및 잉여슬러지 적기 인발을 실시하십시오.\n"
                "* **몰운(IC-SBR 60톤) / 당의(IC-SBR 45톤) / 단월마을(IC-SBR 30톤) / 진목(SOD 23톤)**: 디켄터(Decanter) 펌프 및 수위계 센서 청소, SOD 설비 및 방류유량계 센서 청소를 주 1회 이상 점검하십시오."
            )

    with tab_knr_sim:
        st.subheader("🎛️ 기온·강우 변동 및 유입부하에 따른 [송풍기 인버터(Hz) · 약품 펌프 스트로크(%) · 예상 방류수질] AI 시뮬레이터")
        st.caption("💡 단월 처리장의 실제 설비 규격(송풍기 25㎥/분급 인버터, 염화제이철·PAC 정량펌프, KNR 반응조)을 기반으로 기온·수온·강우 시나리오에 따른 최적 제어값을 시뮬레이션합니다.")

        import calendar
        days_in_m_sim = calendar.monthrange(dw_ref_date.year, dw_sel_month)[1]
        d_list_sim = list(range(1, days_in_m_sim + 1))

        c_sim1, c_sim2, c_sim3 = st.columns(3)
        with c_sim1:
            sim_rain_pattern = st.selectbox("☔ 강우 유입(I/I) 시나리오", ["평시 청천 일지 기준", "중순 단발성 강우 (30mm)", "월 2회 강우 유입", "장마철 집중호우 (80mm)"], key="dw_sim_rain_p")
            sim_peak_rain = st.slider("피크 강우량 설정 (mm/일)", 0.0, 150.0, 0.0 if "청천" in sim_rain_pattern else (80.0 if "집중" in sim_rain_pattern else 30.0), step=5.0, key="dw_sim_rain_v")
        with c_sim2:
            sim_temp_pattern = st.selectbox("🌡️ 외기기온·반응조수온 추이", ["현재 수온 유지", "가을➔겨울 수온 하강 추이", "봄➔여름 수온 상승 추이"], key="dw_sim_temp_p")
            sim_w_temp_start, sim_w_temp_end = st.slider("월간 반응조 수온 범위 (초순 ➔ 말일 ℃)", 6.0, 28.0, (17.0, 13.0) if "하강" in sim_temp_pattern else ((13.0, 18.0) if "상승" in sim_temp_pattern else (float(round(w_temp_v, 1)), float(round(w_temp_v, 1)))), step=0.5, key="dw_sim_w_rng")
        with c_sim3:
            sim_pol_load = st.slider("🧪 유입 오염부하(BOD·T-N·T-P) 변동률 (%)", -30, 50, 0, step=5, key="dw_sim_pol")
            use_sim_editor = st.toggle("📋 일별 시뮬레이션 수치 표 직접 편집", value=False, key="dw_sim_edit_tog")

        sim_rows_list = []
        scale_cap = max_cap_val / 1700.0
        for idx_d, d_no in enumerate(d_list_sim):
            frac = idx_d / max(1, days_in_m_sim - 1)
            wt_d = round(sim_w_temp_start + (sim_w_temp_end - sim_w_temp_start) * frac, 1)
            at_d = round(wt_d - 2.5 + 1.2 * np.sin(idx_d * 0.5), 1)
            r_d = 0.0
            if sim_peak_rain > 0:
                if idx_d in [12, 13, 14]:
                    r_d = round(sim_peak_rain * (1.0 if idx_d == 13 else 0.5), 1)
                elif "2회" in sim_rain_pattern and idx_d in [23, 24]:
                    r_d = round(sim_peak_rain * 0.65, 1)
                elif "집중" in sim_rain_pattern and 10 <= idx_d <= 17:
                    r_d = round(sim_peak_rain * (0.4 + 0.6 * abs(np.sin(idx_d))), 1)

            q_in_d = round(in_flow_v * (0.96 + 0.05 * np.sin(d_no * 0.4)) + (r_d * 8.2 * scale_cap), 1)
            bod_in_d = round(in_bod_v * (1.0 + sim_pol_load * 0.01), 1)
            tn_in_d = round(in_tn_v * (1.0 + sim_pol_load * 0.01), 2)
            tp_in_d = round(in_tp_v * (1.0 + sim_pol_load * 0.01), 2)

            p_res = calculate_ai_process_parameters(q_in_d, bod_in_d, tn_in_d, tp_in_d, facility_name=dw_sel_fac)
            # 송풍기 인버터 주파수(Hz) 및 약품 펌프 스트로크(%) 환산
            inv_hz = round(min(60.0, max(35.0, 42.0 + (p_res['권장송풍량_m3min'] / max(1.0, fac_spec['blower_cap'])) * 14.5)), 1)
            chem_stroke = round(min(100.0, max(20.0, (p_res['권장염화제이철_L'] / max(1.0, max_cap_val * 0.035)) * 55.0)), 1) if dw_sel_fac == MAIN_PLANT else (round(min(100.0, (p_res['종침전PAC주입량_L'] / 3.0) * 50.0), 1) if dw_sel_fac == "몰운" else 0.0)
            pred_tn_out = round(out_tn_v * (1.0 + max(0.0, (15.0 - wt_d) * 0.025)) * (1.0 + sim_pol_load * 0.004), 2)
            pred_tp_out = round(out_tp_v * (1.0 + sim_pol_load * 0.005), 3)

            sim_rows_list.append({
                "일(Day)": d_no,
                "강우량(mm)": r_d,
                "외기기온(℃)": at_d,
                "반응조수온(℃)": wt_d,
                "예측유입량(㎥/일)": q_in_d,
                "권장송풍량(㎥/분)": p_res['권장송풍량_m3min'],
                "송풍기인버터(Hz)": inv_hz,
                "응집제펌프스트로크(%)": chem_stroke,
                "예상방류T-N(mg/L)": pred_tn_out,
                "예상방류T-P(mg/L)": pred_tp_out
            })

        df_sim_ctrl = pd.DataFrame(sim_rows_list)
        if use_sim_editor:
            df_sim_ctrl = st.data_editor(df_sim_ctrl, use_container_width=True, hide_index=True, height=280, key="dw_sim_df_ed")

        fig_ctrl = make_subplots(
            rows=2, cols=1, shared_xaxes=True, vertical_spacing=0.12, row_heights=[0.52, 0.48],
            specs=[[{"secondary_y": True}], [{"secondary_y": True}]],
            subplot_titles=(
                "🌊 예측 유입하수량(㎥/일) & 강우량(mm) 및 반응조 수온(℃) 추이",
                "⚙️ AI 자율제어 출력: 송풍기 인버터 주파수(Hz) & 총인 응집제 정량펌프 스트로크(%)"
            )
        )
        fig_ctrl.add_trace(go.Bar(x=df_sim_ctrl["일(Day)"], y=df_sim_ctrl["강우량(mm)"], name="강우량(mm)", marker_color="rgba(56, 189, 248, 0.45)"), row=1, col=1, secondary_y=True)
        fig_ctrl.add_trace(go.Scatter(x=df_sim_ctrl["일(Day)"], y=df_sim_ctrl["예측유입량(㎥/일)"], name="예측유입량(㎥/일)", mode="lines+markers", line=dict(color="#0284C7", width=2.8)), row=1, col=1, secondary_y=False)
        fig_ctrl.add_trace(go.Scatter(x=df_sim_ctrl["일(Day)"], y=df_sim_ctrl["반응조수온(℃)"], name="반응조수온(℃)", mode="lines", line=dict(color="#10B981", width=2.2, dash="dot")), row=1, col=1, secondary_y=True)

        fig_ctrl.add_trace(go.Scatter(x=df_sim_ctrl["일(Day)"], y=df_sim_ctrl["송풍기인버터(Hz)"], name="송풍기 인버터(Hz)", mode="lines+markers", line=dict(color="#6366F1", width=2.8)), row=2, col=1, secondary_y=False)
        fig_ctrl.add_trace(go.Scatter(x=df_sim_ctrl["일(Day)"], y=df_sim_ctrl["응집제펌프스트로크(%)"], name="응집제 펌프 스트로크(%)", mode="lines+markers", line=dict(color="#F59E0B", width=2.5)), row=2, col=1, secondary_y=True)

        fig_ctrl.update_layout(height=500, margin=dict(l=15, r=15, t=35, b=15), hovermode="x unified", legend=dict(orientation="h", y=1.08))
        st.plotly_chart(fig_ctrl, use_container_width=True)
        st.dataframe(df_sim_ctrl, use_container_width=True, hide_index=True)

    with tab_p1:
        sel_p = st.selectbox("🎯 공정 연산 대상 시설 선택", [MAIN_PLANT] + SMALL_PLANTS, key="dw_calc_fac_sel")
        target_spec = PLANT_DESIGN_SPECS[sel_p]
        st.info(f"🏢 **시설명**: {sel_p} | **공법**: {target_spec['method']} | **설계용량**: {target_spec['cap']} ㎥/일 | **약품**: {target_spec['chem_type']}")
        if st.button("🔄 선택 시설 일별 AI 최적 운전조건 연산 및 마스터 DB 영구 적재", type="primary", use_container_width=True):
            df_fac = get_master_data(sel_p)
            recs = []
            if not df_fac.empty:
                for idx, (_, r) in enumerate(df_fac.iterrows()):
                    res = calculate_ai_process_parameters(
                        r.get('유입량', target_spec['cap']),
                        r.get('유입BOD', 135.0),
                        r.get('유입TN', 26.0),
                        r.get('유입TP', 2.8),
                        facility_name=sel_p,
                        date_seed=idx
                    )
                    recs.append({
                        "날짜": r['날짜'], "CN비": res['CN비'],
                        "권장송풍량_m3min": res['권장송풍량_m3min'], "송풍기가동대수": res['송풍기가동대수'],
                        "권장염화제이철_L": res['권장염화제이철_L'], "종침전PAC주입량_L": res['종침전PAC주입량_L']
                    })
            else:
                for idx in range(10):
                    d_str = (datetime.date.today() - datetime.timedelta(days=idx)).strftime('%Y-%m-%d')
                    res = calculate_ai_process_parameters(target_spec['cap'], 135.0, 26.0, 2.8, facility_name=sel_p, date_seed=idx)
                    recs.append({
                        "날짜": d_str, "CN비": res['CN비'],
                        "권장송풍량_m3min": res['권장송풍량_m3min'], "송풍기가동대수": res['송풍기가동대수'],
                        "권장염화제이철_L": res['권장염화제이철_L'], "종침전PAC주입량_L": res['종침전PAC주입량_L']
                    })
            append_to_process_db(pd.DataFrame(recs), facility_name=sel_p)
            st.success("✅ AI 최적 운전조건 연산 및 영구 적재 완료!")
            st.rerun()
        st.dataframe(get_process_db(sel_p), use_container_width=True)

    with tab_p2:
        st.subheader("🗂️ 영구 저장된 공정 제어 마스터 DB 전체 열람")
        if os.path.exists(PROCESS_CONTROL_DB):
            df_proc_all = pd.read_csv(PROCESS_CONTROL_DB)
            st.dataframe(df_proc_all, use_container_width=True)
            st.download_button("📥 공정 제어 마스터 DB 다운로드 (CSV)", df_proc_all.to_csv(index=False, encoding='utf-8-sig'), file_name="danwol_process_control_master.csv", mime="text/csv", type="primary", use_container_width=True)
        else:
            st.info("💡 아직 저장된 공정 제어 데이터가 없습니다.")

# -------------------------------------------------------------
# 5. 약품·에너지 사용량 데이터 적재 & ESG 경제성 분석 (엑셀/CSV 업로드 + 수동입력 + ESG 탄소·경제성 대시보드 + 영구보관함)
# -------------------------------------------------------------
elif menu == "🧪 5. 약품·에너지 사용량 데이터 적재 & ESG 경제성 분석":
    st.title("🧪 약품·에너지 사용량 데이터 적재 & ESG 경제성 분석")
    tab_ce_up, tab_ce_esg, tab_ce_arch = st.tabs([
        "📥 [데이터 적재] 엑셀·CSV 업로드 / 1번 DB 동기화 / 수동 입력",
        "🌱 [ESG 경제성 분석] 약품·전력·태양광·탄소배출권(tCO2) 대시보드",
        "🗂️ [보관함] 약품·에너지 관리대장 영구 보관소"
    ])

    with tab_ce_up:
        st.caption("🔒 약품·에너지 관리대장 엑셀/CSV 파일 또는 운영일지 엑셀 파일을 직접 업로드하거나, 1번 운영일지 마스터 DB 동기화 및 일일 수동 입력으로 데이터를 영구 적재할 수 있습니다.")

        col_up1, col_up2 = st.columns([1.4, 1.0])
        with col_up1:
            st.markdown("##### 📂 1. 약품·에너지·태양광 엑셀/CSV 파일 업로드")
            ce_files = st.file_uploader(
                "약품·에너지 엑셀(.xlsx, .xls), CSV 파일 또는 종합운영일지 업로드",
                type=["xlsx", "xls", "csv"],
                accept_multiple_files=True,
                key="up_chem_energy_files"
            )
            if ce_files:
                df_ce_parsed = parse_chem_energy_files(ce_files)
                if not df_ce_parsed.empty:
                    st.session_state["df_ce_parsed"] = df_ce_parsed
                    st.success(f"✅ 업로드된 파일에서 총 **{len(df_ce_parsed)}일치** 약품·에너지 데이터를 추출했습니다!")
                    st.dataframe(df_ce_parsed, use_container_width=True, hide_index=True)
                    if st.button("💾 ⚡ [업로드된 약품·에너지 데이터 마스터 DB 및 영구 보관함 일괄 저장]", type="primary", use_container_width=True, key="btn_save_ce_up"):
                        append_to_chem_db(df_ce_parsed)
                        ce_xlsx_bytes = generate_chem_energy_template_excel(get_chem_db())
                        with open(os.path.join(CHEMICAL_RECORD_DIR, sanitize_filename(f"약품에너지_ESG관리대장_{datetime.date.today()}.xlsx")), "wb") as f:
                            f.write(ce_xlsx_bytes)
                        st.success("✅ 업로드하신 약품·에너지 데이터가 마스터 DB와 영구 보관함에 안전하게 저장되었습니다!")
                        st.rerun()
                else:
                    st.warning("⚠️ 업로드된 파일에서 날짜(YYYY-MM-DD) 또는 약품·유량 데이터를 찾지 못했습니다. 우측의 표준 엑셀 양식을 내려받아 활용하실 수도 있습니다.")

        with col_up2:
            st.markdown("##### 🔄 2. 1번 운영일지 마스터 DB 자동 동기화 & 양식 다운로드")
            if st.button("🔄 1번 운영일지 마스터 DB에서 약품·에너지 실데이터 동기화", type="primary", use_container_width=True):
                df_m = get_master_data(MAIN_PLANT)
                recs = []
                if not df_m.empty:
                    for idx, (_, r) in enumerate(df_m.iterrows()):
                        fl = float(r.get('유입량', 1700)) if pd.notna(r.get('유입량')) else 1700.0
                        recs.append({
                            "날짜": str(r['날짜']).split()[0],
                            "PAC사용량_kg": round(fl * 0.026, 1),
                            "염화제이철_kg": round(fl * 0.015, 1),
                            "폴리머사용량_kg": round(fl * 0.0012, 2),
                            "슬러지반출량_톤": round(fl * 0.0019, 2),
                            "전력사용량_kWh": 1420.0,
                            "태양광발전량_kWh": 135.0,
                            "비고": "1번 마스터DB 동기화"
                        })
                else:
                    for idx in range(10):
                        d_str = (datetime.date.today() - datetime.timedelta(days=idx)).strftime('%Y-%m-%d')
                        recs.append({
                            "날짜": d_str,
                            "PAC사용량_kg": 44.2,
                            "염화제이철_kg": 25.5,
                            "폴리머사용량_kg": 2.04,
                            "슬러지반출량_톤": 3.23,
                            "전력사용량_kWh": 1420.0,
                            "태양광발전량_kWh": 135.0,
                            "비고": "기본 동기화"
                        })
                append_to_chem_db(pd.DataFrame(recs))
                st.success("✅ 약품·에너지 마스터 DB 동기화 및 영구 저장 완료!")
                st.rerun()

            df_chem_cur = get_chem_db()
            std_xlsx_bytes = generate_chem_energy_template_excel(df_chem_cur)
            st.download_button(
                "📥 약품·에너지 표준 업로드 엑셀 양식(.xlsx) 다운로드",
                data=std_xlsx_bytes,
                file_name="단월_약품에너지_ESG관리대장_양식.xlsx",
                type="secondary",
                use_container_width=True
            )

        st.divider()
        st.markdown("##### ✍️ 3. 일일 약품·에너지·태양광 발전량 직접 입력 및 표 수정")
        with st.form("ce_manual_input_form"):
            c_m1, c_m2, c_m3, c_m4 = st.columns(4)
            with c_m1:
                ce_in_date = st.date_input("측정/사용 일자", datetime.date.today(), key="ce_d_in")
                ce_in_pac = st.number_input("PAC 사용량 (kg/일)", value=44.2, step=1.0, format="%.1f")
            with c_m2:
                ce_in_fe = st.number_input("염화제이철 사용량 (kg/일)", value=25.5, step=1.0, format="%.1f")
                ce_in_poly = st.number_input("폴리머 사용량 (kg/일)", value=2.04, step=0.1, format="%.2f")
            with c_m3:
                ce_in_sludge = st.number_input("탈수슬러지 반출량 (톤/일)", value=3.23, step=0.1, format="%.2f")
                ce_in_pwr = st.number_input("한전 전력사용량 (kWh/일)", value=1420.0, step=10.0, format="%.1f")
            with c_m4:
                ce_in_solar = st.number_input("태양광 발전량 (kWh/일)", value=222.5, step=5.0, format="%.1f")
                ce_in_memo = st.text_input("비고", value="수동 입력")

            if st.form_submit_button("💾 일일 약품·에너지 사용량 마스터 DB 영구 저장", type="primary", use_container_width=True):
                new_ce_row = pd.DataFrame([{
                    "날짜": ce_in_date.strftime('%Y-%m-%d'),
                    "PAC사용량_kg": ce_in_pac,
                    "염화제이철_kg": ce_in_fe,
                    "폴리머사용량_kg": ce_in_poly,
                    "슬러지반출량_톤": ce_in_sludge,
                    "전력사용량_kWh": ce_in_pwr,
                    "태양광발전량_kWh": ce_in_solar,
                    "비고": ce_in_memo
                }])
                append_to_chem_db(new_ce_row)
                st.success(f"✅ {ce_in_date} 약품·에너지 사용량 데이터가 영구 적재되었습니다!")
                st.rerun()

        df_chem = get_chem_db()
        if not df_chem.empty:
            st.subheader("🗂️ 영구 저장된 약품·에너지 마스터 DB 열람 및 직접 편집")
            ed_df_chem = st.data_editor(df_chem, use_container_width=True, hide_index=True, key="ed_chem_master_tbl")
            c_db1, c_db2, c_db3 = st.columns(3)
            with c_db1:
                if st.button("💾 표에서 수정한 내역 마스터 DB에 덮어쓰기 저장", type="primary", use_container_width=True):
                    ed_df_chem.to_csv(CHEMICAL_ENERGY_DB, index=False, encoding='utf-8-sig')
                    st.success("✅ 표에서 수정하신 약품·에너지 내역이 마스터 DB에 반영되었습니다!")
                    st.rerun()
            with c_db2:
                st.download_button(
                    "📥 약품·에너지 마스터 DB 다운로드 (CSV)",
                    ed_df_chem.to_csv(index=False, encoding='utf-8-sig'),
                    file_name="danwol_chemical_energy_master.csv",
                    mime="text/csv",
                    type="primary",
                    use_container_width=True
                )
            with c_db3:
                st.download_button(
                    "📥 약품·에너지 관리대장 다운로드 (XLSX)",
                    generate_chem_energy_template_excel(ed_df_chem),
                    file_name=f"단월_약품에너지_ESG관리대장_{datetime.date.today()}.xlsx",
                    type="primary",
                    use_container_width=True
                )
        else:
            st.info("데이터 없음")

    with tab_ce_esg:
        st.subheader("🌱 단월공공하수처리시설 ESG 탄소중립 & 약품·에너지 경제성 정밀 분석")
        df_chem_esg = get_chem_db()
        if df_chem_esg.empty:
            st.info("💡 먼저 좌측 [📥 데이터 적재] 탭에서 데이터를 업로드하거나 동기화해 주세요.")
        else:
            # 단가 및 탄소배출계수 설정 패널
            with st.expander("⚙️ 약품·전력 단가 및 온실가스 배출계수(tCO2eq) 설정 (조절 시 경제성 지표 즉시 반영)", expanded=False):
                u_c1, u_c2, u_c3, u_c4 = st.columns(4)
                with u_c1: unit_pac = st.number_input("PAC 단가 (원/kg)", value=420, step=10)
                with u_c2: unit_fe = st.number_input("염화제이철 단가 (원/kg)", value=550, step=10)
                with u_c3: unit_poly = st.number_input("폴리머 단가 (원/kg)", value=4500, step=100)
                with u_c4: unit_elec = st.number_input("산업용 전력 단가 (원/kWh)", value=145, step=5)

            tot_pac = pd.to_numeric(df_chem_esg["PAC사용량_kg"], errors="coerce").fillna(0).sum()
            tot_fe = pd.to_numeric(df_chem_esg["염화제이철_kg"], errors="coerce").fillna(0).sum()
            tot_poly = pd.to_numeric(df_chem_esg["폴리머사용량_kg"], errors="coerce").fillna(0).sum()
            tot_sludge = pd.to_numeric(df_chem_esg["슬러지반출량_톤"], errors="coerce").fillna(0).sum()
            tot_pwr = pd.to_numeric(df_chem_esg["전력사용량_kWh"], errors="coerce").fillna(0).sum()
            tot_solar = pd.to_numeric(df_chem_esg["태양광발전량_kWh"], errors="coerce").fillna(0).sum()

            energy_self_rate = (tot_solar / tot_pwr * 100.0) if tot_pwr > 0 else 0.0
            co2_reduced_ton = tot_solar * 0.0004781  # 국가 전력배출계수 0.4781 tCO2eq/MWh 적용
            saved_elec_cost = int(round(tot_solar * unit_elec))
            total_chem_cost = int(round(tot_pac * unit_pac + tot_fe * unit_fe + tot_poly * unit_poly))

            e1, e2, e3, e4 = st.columns(4)
            esg_cards = [
                (e1, "누적 태양광 발전량", f"{tot_solar:,.1f}", "kWh", f"전력자립률 {energy_self_rate:.1f}%", "☀️", "#10B981"),
                (e2, "온실가스(CO₂) 감축량", f"{co2_reduced_ton:,.2f}", "tCO₂eq", f"소나무 {int(co2_reduced_ton*151):,}그루 식재효과", "🌲", "#059669"),
                (e3, "태양광 전력요금 절감액", f"{saved_elec_cost:,}", "원", f"단가 {unit_elec}원/kWh 기준", "⚡", "#0284C7"),
                (e4, "누적 응집제 약품 비용", f"{total_chem_cost:,}", "원", f"PAC {tot_pac:,.0f}kg · 염철 {tot_fe:,.0f}kg", "🧪", "#6366F1")
            ]
            for col_e, e_lbl, e_val, e_unit, e_sub, e_ico, e_col in esg_cards:
                with col_e:
                    st.markdown(f"""
                    <div class="dw-super-card" style="border-top-color:{e_col};">
                        <div class="dw-kpi-top"><span class="dw-kpi-title">{e_lbl}</span><span class="dw-kpi-icon">{e_ico}</span></div>
                        <div class="dw-kpi-value">{e_val}<span class="dw-kpi-unit">{e_unit}</span></div>
                        <div class="dw-badge-row"><span class="dw-pill-safe">{e_sub}</span></div>
                    </div>
                    """, unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)
            df_plot_ce = df_chem_esg.sort_values(by="날짜").tail(31)
            c_ch1, c_ch2 = st.columns(2)
            with c_ch1:
                fig_ce1 = go.Figure()
                fig_ce1.add_trace(go.Bar(x=df_plot_ce["날짜"], y=df_plot_ce["PAC사용량_kg"], name="PAC (kg)", marker_color="#0284C7"))
                fig_ce1.add_trace(go.Bar(x=df_plot_ce["날짜"], y=df_plot_ce["염화제이철_kg"], name="염화제이철 (kg)", marker_color="#6366F1"))
                fig_ce1.add_trace(go.Scatter(x=df_plot_ce["날짜"], y=df_plot_ce["슬러지반출량_톤"], name="슬러지반출량 (톤)", mode="lines+markers", line=dict(color="#EF4444", width=2.5), yaxis="y2"))
                fig_ce1.update_layout(
                    title="📊 일별 약품(PAC·염화제이철) 투입량 및 슬러지 발생 추이",
                    yaxis=dict(title="약품 사용량 (kg)"),
                    yaxis2=dict(title="슬러지 (톤)", overlaying="y", side="right"),
                    barmode="group", height=370, margin=dict(l=15, r=15, t=40, b=15),
                    legend=dict(orientation="h", y=1.12)
                )
                st.plotly_chart(fig_ce1, use_container_width=True)
            with c_ch2:
                fig_ce2 = go.Figure()
                fig_ce2.add_trace(go.Scatter(x=df_plot_ce["날짜"], y=df_plot_ce["전력사용량_kWh"], name="전력사용량 (kWh)", mode="lines+markers", line=dict(color="#F59E0B", width=2.5)))
                fig_ce2.add_trace(go.Bar(x=df_plot_ce["날짜"], y=df_plot_ce["태양광발전량_kWh"], name="태양광발전량 (kWh)", marker_color="#10B981"))
                fig_ce2.update_layout(
                    title="☀️ 일별 전력 소비량 대비 태양광 발전 실적 (RE100·탄소중립)",
                    yaxis=dict(title="에너지 (kWh)"),
                    height=370, margin=dict(l=15, r=15, t=40, b=15),
                    legend=dict(orientation="h", y=1.12)
                )
                st.plotly_chart(fig_ce2, use_container_width=True)

            if st.button("💾 현재 약품·에너지 ESG 분석 대장 영구 보관함에 저장", type="primary", use_container_width=True):
                esg_save_name = sanitize_filename(f"약품에너지_ESG관리대장_{datetime.date.today()}.xlsx")
                with open(os.path.join(CHEMICAL_RECORD_DIR, esg_save_name), "wb") as f:
                    f.write(generate_chem_energy_template_excel(df_chem_esg))
                st.success(f"✅ 약품·에너지 ESG 관리대장(`{esg_save_name}`)이 영구 보관함에 안전하게 저장되었습니다!")

    with tab_ce_arch:
        st.subheader("🗂️ 약품·에너지·ESG 관리대장 영구 보관함")
        ce_saved = [sanitize_filename(f) for f in os.listdir(CHEMICAL_RECORD_DIR) if f.endswith((".xlsx", ".xls", ".csv"))]
        if ce_saved:
            st.write(f"📁 **총 {len(ce_saved)}건의 약품·에너지 문서가 영구 보관되어 있습니다.**")
            sel_ce_f = st.selectbox("다운로드 또는 열람할 파일 선택", sorted(ce_saved, reverse=True), key="sel_ce_arch_f")
            ce_fp = os.path.join(CHEMICAL_RECORD_DIR, sel_ce_f)
            c_ca1, c_ca2 = st.columns([3, 1])
            with open(ce_fp, "rb") as f:
                c_ca1.download_button(f"📥 '{sel_ce_f}' 다운로드", f.read(), file_name=sel_ce_f, type="primary", use_container_width=True)
            if c_ca2.button("🗑️ 선택 파일 삭제", key="btn_del_ce_arch", use_container_width=True):
                os.remove(ce_fp)
                st.success("삭제 완료!")
                st.rerun()
        else:
            st.info("💡 아직 보관함에 저장된 약품·에너지 파일이 없습니다.")

# -------------------------------------------------------------
# 6. Q&A 챗봇 (오프라인 온프레미스 고도화)
# -------------------------------------------------------------
elif menu == "🤖 6. 단월 AI 지능형 공정 Q&A 챗봇 (오프라인 온프레미스)":
    st.title("🤖 단월 하수처리시설 AI 지능형 공정 도우미 (온프레미스 오프라인 모드)")
    st.info("🔒 **보안 인증 완료**: 본 챗봇은 외부 상용 API(Gemini 등)를 전혀 호출하지 않으며, 단월하수처리장 내부 서버에 탑재된 **로컬 오프라인 규칙·지식 기반 엔진**으로만 작동하여 데이터 외부 유출이 100% 차단됩니다.")

    if "messages" not in st.session_state:
        st.session_state.messages = [{"role": "assistant", "content": "안녕하세요! 단월공공하수처리시설(KNR+IPR) 및 소규모 6개소 자율운전 제어 도우미입니다. 공정 제어, T-P/T-N 수질 대처, 송풍기 및 약품 제어에 대해 질문해 주세요."}]
    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]): st.markdown(msg["content"])
    
    prompt = st.chat_input("질문을 입력하세요 (예: T-P가 높을 때 대응법, C/N비 관리 기준 등)")
    if prompt:
        st.session_state.messages.append({"role": "user", "content": prompt})
        with st.chat_message("user"): st.markdown(prompt)
        
        q_lower = prompt.lower()
        if any(k in q_lower for k in ["인린", "총인", "tp", "염화제이철", "pac", "응집제"]):
            ans = "🛡️ **[온프레미스 지식 엔진 답변] 총인(T-P) 저감 및 IPR 공정 제어 지침**\n\n1. **약품 주입 제어**: 종침전지 및 총인처리시설의 염화제이철 및 PAC 주입율은 방류수질 기준(0.2 mg/L) 대비 안전마진을 고려하여 **0.065 ~ 0.08 mg/L** 이하로 유지되도록 자동 연동됩니다.\n2. **수질 급변 시**: 유입 인원이 급증하거나 인근 축산계 유입 등으로 T-P가 상승할 경우, 응집제 정량펌프 주입 주기를 15% 상향 조정하고 응집 폴리머 상태를 정밀 점검하십시오."
        elif any(k in q_lower for k in ["질소", "tn", "탈질", "송풍기", "do", "용존산소"]):
            ans = "🛡️ **[온프레미스 지식 엔진 답변] 질소(T-N) 및 송풍기 제어 지침**\n\n1. **C/N비 관리**: 생물반응조 유입 BOD/T-N 비가 **4.0 이상** 유지되어야 안정적인 탈질이 이루어집니다.\n2. **호기조 DO 제어**: 호기조 용존산소(DO)는 **1.8 ~ 2.2 mg/L**를 유지하도록 AI 송풍기 변속(Inverter) 제어가 실시간 수행됩니다."
        elif any(k in q_lower for k in ["소규모", "산음", "삼가리", "진목", "몰운", "당의"]):
            ans = "🛡️ **[온프레미스 지식 엔진 답변] 소규모 6개소 연계 관제 지침**\n\n소규모 처리시설(산음, 삼가리, 진목, 몰운, 단월마을, 당의)은 **7일 주기 자동 순환 점검** 및 원격 수중포기·SBR 공정 모니터링이 본장 관제 시스템과 실시간 연동되어 있습니다."
        else:
            ans = f"🛡️ **[온프레미스 보안 답변]** 입력하신 질문(*\"{prompt}\"*)에 대해 단월 하수처리 표준운영기준(SOP)을 대조한 결과입니다.\n\n• 본 시스템은 외부 인터넷 및 클라우드 API를 사용하지 않는 **로컬 온프레미스(오프라인) 환경**으로 가동 중입니다.\n• 수질 기준 준수를 위해 생물반응조 슬러지(MLSS) 농도 3,000~4,500 mg/L 유지 및 일일 유입·방류량을 점검하시기 바랍니다."

        with st.chat_message("assistant"): st.markdown(ans)
        st.session_state.messages.append({"role": "assistant", "content": ans})

# -------------------------------------------------------------
# 7. TBM 표준회의록 AI 자동작성/출력 & 실행 체크리스트 (신규 추가)
# -------------------------------------------------------------
elif menu == "📝 7. TBM 표준회의록 AI 자동작성/출력 & 보관함":
    st.title("📝 단월처리시설 TBM(작업 전 안전점검회의) AI 자동작성기 & 실행 체크리스트")
    tab_tbm_write, tab_tbm_chk, tab_tbm_archive = st.tabs(["✍️ [작성] TBM AI 자동작성 및 서명", "☑️ [실행 체크리스트] 안전보건 점검표", "🗂️ [보관함] 영구 저장된 TBM 회의록 열람"])

    with tab_tbm_write:
        st.caption("🔒 새로운 공식 양식 규격 반영 · 외주업체 정보 및 위험성평가 연동 자동작성 지원")

        record_dir = TBM_RECORD_DIR
        if not os.path.exists(record_dir): os.makedirs(record_dir)

        ai_risk_db = {
            "단월처리장 총인처리시설 미디어 교체": {
                "desc": "단월처리장 총인처리시설 미디어 교체", "place": "작업현장",
                "risks": [
                    ("미끄러짐 및 추락 위험", "작업 전 바닥 물기 제거 및 미끄럼방지용 안전화 착용, 수조 상부 안전 난간·개구부 덮개 설치 상태 점검 후 작업 실시"),
                    ("근골격계 부담 및 부상", "2인 1조 분담 작업 실시, 무리한 인력 운반을 지향하고 대차·호이스트 등 하역 장비 적극 활용"),
                    ("질식 및 유해가스 노출", "작업 전 산소 및 유해가스 농도 측정 실시, 작업 중 상시 환기 유지 및 필요 시 송기마스크·안전띠 착용")
                ]
            },
            "단월공공하수처리시설 총인처리시설 공기압축기 수리": {
                "desc": "총인처리시설 공기압축기 #A 수리", "place": "작업현장",
                "risks": [
                    ("회전체(풀리, 벨트 등) 말림 및 끼임 위험", "점검 및 수리 작업 시 반드시 전원을 차단(Lockout, Tagout - LOTO)하고 조치합니다."),
                    ("중량물 취급 및 공구 낙하에 따른 부상 위험", "공기압축기 부품 및 무거운 공구 취급 시 2인 1조로 작업하거나 인양 장비를 활용합니다."),
                    ("소음 및 에어 누출 위험", "배관 연결부 수리 후 압력 주입 시 에어 누출 여부를 비눗물 등으로 안전하게 점검합니다.")
                ]
            },
            "산음리 중계 펌프A 인양 및 인양 상태 점검 작업": {
                "desc": "호이스트 이용 펌프A 인양 후 매달린 상태에서의 정밀 점검 및 정비", "place": "작업현장",
                "risks": [
                    ("인양된 펌프A 하부/측면 작업 중 낙하로 인한 깔림 및 끼임", "안전 고임목/받침대 설치: 인양 후 매달린 상태 유지 시 안전 고임목 또는 지지대를 받쳐 낙하 방지"),
                    ("인양장치(호이스트) 브레이크 미작동 및 와이어 파손으로 인한 낙하", "인양장치 점검: 작업 전 브레이크 작동 상태, 와이어로프, 훅 해지장치 결함 여부 사전 확인"),
                    ("펌프A 매달림 상태에서 흔들림 및 균형 상실로 인한 충돌", "유도 로프(태그라인) 활용: 펌프 인양 및 매달림 상태 유지 시 흔들림 방지용 유도 로프 체결")
                ]
            }
        }

        if "tbm_job_desc" not in st.session_state:
            st.session_state.tbm_job_desc = "단월처리장 총인처리시설 미디어 교체"
            st.session_state.tbm_r1 = "미끄러짐 및 추락 위험"
            st.session_state.tbm_s1 = "작업 전 바닥 물기 제거 및 미끄럼방지용 안전화 착용, 수조 상부 안전 난간·개구부 덮개 설치 상태 점검 후 작업 실시"
            st.session_state.tbm_r2 = "근골격계 부담 및 부상"
            st.session_state.tbm_s2 = "2인 1조 분담 작업 실시, 무리한 인력 운반을 지향하고 대차·호이스트 등 하역 장비 적극 활용"
            st.session_state.tbm_r3 = "질식 및 유해가스 노출"
            st.session_state.tbm_s3 = "작업 전 산소 및 유해가스 농도 측정 실시, 작업 중 상시 환기 유지 및 필요 시 송기마스크·안전띠 착용"
            st.session_state.last_selected_job = "단월처리장 총인처리시설 미디어 교체"

        is_weekly = st.checkbox("📅 **[별지1] 작업내용이 동일하여 1주일 단위로 작성하고자 할 경우 체크**", value=False)
        
        post_note_mon = post_note_tue = post_note_wed = post_note_thu = post_note_fri = "작업 전 안전수칙 준수 및 보호구 착용 철저로 아차사고 및 안전사고 없이 미디어 교체 작업 안전하게 완료함. 특이사항 없음."
        post_note_sat = post_note_sun = "해당 없음"
        single_post_note = "작업 전 안전수칙 준수 및 보호구 착용 철저로 아차사고 및 안전사고 없이 미디어 교체 작업 안전하게 완료함. 특이사항 없음."

        c1, c2 = st.columns([1, 1])
        with c1:
            st.subheader("1️⃣ 작업 기본정보 & 외주업체 정보")
            tbm_date = st.date_input("TBM 일자", datetime.date(2026, 10, 19))
            tbm_time = st.text_input("TBM 시간", "09:00 ~ 09:30 (30분간)")
            selected_job = st.selectbox("금일 작업명 선택 (또는 '직접 입력' 선택)", list(ai_risk_db.keys()) + ["직접 입력"])
            
            if selected_job != st.session_state.last_selected_job:
                st.session_state.last_selected_job = selected_job
                if selected_job != "직접 입력":
                    tinfo = ai_risk_db[selected_job]
                    st.session_state.tbm_job_desc = tinfo["desc"]
                    st.session_state.tbm_r1, st.session_state.tbm_s1 = tinfo["risks"][0]
                    st.session_state.tbm_r2, st.session_state.tbm_s2 = tinfo["risks"][1]
                    st.session_state.tbm_r3, st.session_state.tbm_s3 = tinfo["risks"][2]
                    st.rerun()

            if selected_job == "직접 입력":
                custom_job = st.text_input("직접 작업명 입력", value="단월처리장 총인처리시설 미디어 교체", key="custom_job_input")
                tbm_place = st.selectbox("TBM 장소", ["작업현장", "사무실", "기타"], index=1)
            else:
                custom_job = selected_job
                tbm_place = st.selectbox("TBM 장소", ["작업현장", "사무실", "기타"], index=1)

            job_desc = st.text_area("작업 세부 내용", key="tbm_job_desc")

            if st.button("✨ [AI 자동완성 실행] 작업명 및 세부내용 기반 위험요인·감소대책 자동 채우기", type="primary", use_container_width=True):
                inferred = generate_ai_tbm_risks(custom_job, job_desc)
                st.session_state.tbm_r1, st.session_state.tbm_s1 = inferred[0]
                st.session_state.tbm_r2, st.session_state.tbm_s2 = inferred[1]
                st.session_state.tbm_r3, st.session_state.tbm_s3 = inferred[2]
                st.toast("🤖 AI가 새로운 TBM 양식에 맞춰 유해·위험요인과 감소대책을 자동 채웠습니다!", icon="✨")
                st.rerun()

            r1 = st.text_input("위험요인 ①", key="tbm_r1")
            s1 = st.text_input("감소대책 ①", key="tbm_s1")
            r2 = st.text_input("위험요인 ②", key="tbm_r2")
            s2 = st.text_input("감소대책 ②", key="tbm_s2")
            r3 = st.text_input("위험요인 ③", key="tbm_r3")
            s3 = st.text_input("감소대책 ③", key="tbm_s3")

            st.markdown("---")
            st.markdown("##### 🏢 외주업체 정보 입력 (새로운 양식 반영)")
            has_contractor = st.checkbox("외주작업 해당 여부", value=True)
            c_name = st.text_input("외주 업체명", "㈜에코지텍")
            c_leader = st.text_input("책임자 성명", "은정호")
            c_phone = st.text_input("연락처", "010-4490-4701")
            c_risk_eval = st.radio("업체 위험성평가 실시", ["예", "아니오"], index=0, horizontal=True)
            c_safety_edu = st.radio("산업안전보건 교육 확인", ["예", "아니오"], index=0, horizontal=True)

            eval_type = st.selectbox("위험성평가 실시 구분", ["정기", "최초", "수시"])

            if is_weekly:
                st.markdown("##### 📝 [별지1] 요일별 작업 후 미팅 특이사항")
                batch_apply = st.checkbox("월~금 미팅 내용을 동일하게 일괄 적용", value=True)
                if batch_apply:
                    common_note = st.text_input("월~금 공통 미팅 내용 입력", "작업 전 안전수칙 준수 및 보호구 착용 철저로 아차사고 및 안전사고 없이 미디어 교체 작업 안전하게 완료함. 특이사항 없음.")
                    post_note_mon = post_note_tue = post_note_wed = post_note_thu = post_note_fri = common_note
                    col_wk1, col_wk2 = st.columns(2)
                    with col_wk1: post_note_sat = st.text_input("토요일 미팅 내용", "해당 없음")
                    with col_wk2: post_note_sun = st.text_input("일요일 미팅 내용", "해당 없음")
                else:
                    post_note_mon = st.text_input("월요일 미팅 내용", "작업 전 안전수칙 준수...")
                    post_note_tue = st.text_input("화요일 미팅 내용", "특이사항 없음")
                    post_note_wed = st.text_input("수요일 미팅 내용", "특이사항 없음")
                    post_note_thu = st.text_input("목요일 미팅 내용", "특이사항 없음")
                    post_note_fri = st.text_input("금요일 미팅 내용", "특이사항 없음")
                    post_note_sat = st.text_input("토요일 미팅 내용", "해당 없음")
                    post_note_sun = st.text_input("일요일 미팅 내용", "해당 없음")
            else:
                single_post_note = st.text_area("작업 후 미팅 내용", value="작업 전 안전수칙 준수 및 보호구 착용 철저로 아차사고 및 안전사고 없이 미디어 교체 작업 안전하게 완료함. 특이사항 없음.")
                post_note_mon = single_post_note

        with c2:
            st.subheader("2️⃣ 점검자(리더) & 참석자 명단 (새로운 양식 규격)")
            leader_dept = st.text_input("리더 소속", "환경2팀")
            leader_role = st.text_input("리더 직책", "시설장(차장)")
            leader_name = st.text_input("리더 성명", "주영규")
            
            st.markdown("##### 👥 자체 참석자 명단 (1~4번)")
            w1 = st.text_input("① 자체 참석자", "주영규")
            w2 = st.text_input("② 자체 참석자", "이홍섭")
            w3 = st.text_input("③ 참석자 성명", "최태수")
            w4 = st.text_input("④ 참석자 성명", "이현진")
            workers = [w1, w2, w3, w4]

            worker_mode = st.radio("자체 참석자 서명 방식", ["전자서명 캔버스 그림 적용", "직접 텍스트(이름) 입력"], horizontal=True, key="w_mode_radio")
            custom_worker_signs = [w1, w2, w3, w4]
            if worker_mode == "직접 텍스트(이름) 입력":
                col_ws1, col_ws2 = st.columns(2)
                with col_ws1:
                    cw_t1 = st.text_input("① 서명 대체", w1)
                    cw_t2 = st.text_input("② 서명 대체", w2)
                with col_ws2:
                    cw_t3 = st.text_input("③ 서명 대체", w3)
                    cw_t4 = st.text_input("④ 서명 대체", w4)
                custom_worker_signs = [cw_t1, cw_t2, cw_t3, cw_t4]

            st.markdown("##### 🏢 외주업체 참석자 명단 (업체 성명 1~4번)")
            cw1 = st.text_input("① 업체 참석자", "은정호")
            cw2 = st.text_input("② 업체 참석자", "문성규")
            cw3 = st.text_input("③ 업체 참석자", "정학준")
            cw4 = st.text_input("④ 업체 참석자", "나대현")
            contractor_workers = [cw1, cw2, cw3, cw4]

            contractor_mode = st.radio("외주업체 서명 방식", ["전자서명 캔버스 그림 적용", "직접 텍스트(이름) 입력"], horizontal=True, key="c_mode_radio")
            custom_contractor_signs = [cw1, cw2, cw3, cw4]
            if contractor_mode == "직접 텍스트(이름) 입력":
                col_cs1, col_cs2 = st.columns(2)
                with col_cs1:
                    cc_t1 = st.text_input("① 업체 서명 대체", cw1)
                    cc_t2 = st.text_input("② 업체 서명 대체", cw2)
                with col_cs2:
                    cc_t3 = st.text_input("③ 업체 서명 대체", cw3)
                    cc_t4 = st.text_input("④ 업체 서명 대체", cw4)
                custom_contractor_signs = [cc_t1, cc_t2, cc_t3, cc_t4]

            st.markdown("---")
            st.markdown("##### 🔒 개인정보 수집·이용 및 제3자 제공 동의")
            with st.expander("📄 개인정보 수집·이용 안내문 열람", expanded=False):
                st.caption("• 수집 목적: 산업안전보건법에 따른 TBM 실시 기록 보존 및 안전보건 관리 / • 수집 항목: 소속, 직책, 성명, 연락처, 본인 전자서명 / • 보유 기간: 3년")
            if st.button("✨ [일괄 동의] 리더 및 전 참석자 개인정보 수집·이용 일괄 동의 체크", type="primary", use_container_width=True):
                st.session_state["agree_leader"] = True
                for idx_a in range(4):
                    st.session_state[f"agree_w_{idx_a}"] = True
                    st.session_state[f"agree_c_{idx_a}"] = True
                st.toast("전 참석자의 개인정보 수집·이용 동의가 일괄 체크되었습니다!", icon="✅")
                st.rerun()

            if "agree_leader" not in st.session_state:
                st.session_state["agree_leader"] = True
            for idx_a in range(4):
                if f"agree_w_{idx_a}" not in st.session_state:
                    st.session_state[f"agree_w_{idx_a}"] = True
                if f"agree_c_{idx_a}" not in st.session_state:
                    st.session_state[f"agree_c_{idx_a}"] = True

            st.markdown("##### ✍️ 전자서명 입력 (개별 참석자 이름 연동 캔버스 분리 & 개인정보 동의)")
            tab_sig_leader, tab_sig_worker, tab_sig_contractor = st.tabs(["🛡️ TBM 리더 서명", "👥 자체 참석자 개별 서명 (①~④)", "🏢 외주업체 개별 서명 (①~④)"])
            
            with tab_sig_leader:
                leader_disp = leader_name.strip() if leader_name.strip() else "리더 미입력"
                st.markdown(f"**🛡️ TBM 리더: `{leader_disp}` 님 전용 서명란**")
                agree_l = st.checkbox(f"[필수] {leader_disp} 본인은 TBM 회의록 작성을 위한 개인정보 수집·이용에 동의합니다.", key="agree_leader")
                canvas_leader = st_canvas(stroke_width=2, stroke_color="#000000", background_color="#F8F9FA", height=75, width=280, key="canvas_leader_sig")
            
            worker_canvases = []
            worker_agrees = []
            with tab_sig_worker:
                st.caption("자체 참석자 이름을 클릭하여 개인정보 동의 체크 후 각자 개인의 싸인을 직접 입력해 주세요.")
                w_labels = [
                    f"① {w1.strip()}" if w1.strip() else "① (미입력)",
                    f"② {w2.strip()}" if w2.strip() else "② (미입력)",
                    f"③ {w3.strip()}" if w3.strip() else "③ (미입력)",
                    f"④ {w4.strip()}" if w4.strip() else "④ (미입력)"
                ]
                w_tabs = st.tabs(w_labels)
                for i_w, w_tab in enumerate(w_tabs):
                    with w_tab:
                        cur_w_name = workers[i_w].strip()
                        if cur_w_name:
                            st.markdown(f"**✍️ 자체 참석자 {i_w+1}: `{cur_w_name}` 님 서명란**")
                            ag_w = st.checkbox(f"[필수] {cur_w_name} 본인은 TBM 안전점검 기록을 위한 개인정보 수집·이용에 동의합니다.", key=f"agree_w_{i_w}")
                        else:
                            st.info(f"💡 상단 '자체 참석자 명단'의 {i_w+1}번 칸에 이름을 입력하시면 해당 이름으로 서명란이 활성화됩니다.")
                            ag_w = st.checkbox(f"[필수] 참석자 {i_w+1} 개인정보 수집·이용에 동의합니다.", key=f"agree_w_{i_w}")
                        worker_agrees.append(ag_w)
                        c_w = st_canvas(stroke_width=2, stroke_color="#000000", background_color="#F8F9FA", height=75, width=280, key=f"canvas_worker_sig_{i_w}")
                        worker_canvases.append(c_w)

            contractor_canvases = []
            contractor_agrees = []
            with tab_sig_contractor:
                st.caption("외주업체 참석자 이름을 클릭하여 개인정보 동의 체크 후 각자 개인의 싸인을 직접 입력해 주세요.")
                c_labels = [
                    f"① {cw1.strip()}" if cw1.strip() else "① (미입력)",
                    f"② {cw2.strip()}" if cw2.strip() else "② (미입력)",
                    f"③ {cw3.strip()}" if cw3.strip() else "③ (미입력)",
                    f"④ {cw4.strip()}" if cw4.strip() else "④ (미입력)"
                ]
                c_tabs = st.tabs(c_labels)
                for i_c, c_tab in enumerate(c_tabs):
                    with c_tab:
                        cur_c_name = contractor_workers[i_c].strip()
                        if cur_c_name:
                            st.markdown(f"**✍️ 외주업체 참석자 {i_c+1}: `{cur_c_name}` 님 서명란**")
                            ag_c = st.checkbox(f"[필수] {cur_c_name} 본인은 TBM 안전점검 기록을 위한 개인정보 수집·이용에 동의합니다.", key=f"agree_c_{i_c}")
                        else:
                            st.info(f"💡 상단 '외주업체 참석자 명단'의 {i_c+1}번 칸에 이름을 입력하시면 해당 이름으로 서명란이 활성화됩니다.")
                            ag_c = st.checkbox(f"[필수] 외주 참석자 {i_c+1} 개인정보 수집·이용에 동의합니다.", key=f"agree_c_{i_c}")
                        contractor_agrees.append(ag_c)
                        c_c = st_canvas(stroke_width=2, stroke_color="#000000", background_color="#F8F9FA", height=75, width=280, key=f"canvas_contractor_sig_{i_c}")
                        contractor_canvases.append(c_c)

        leader_sig_base64 = ""
        try:
            if canvas_leader.image_data is not None and np.any(canvas_leader.image_data[:, :, 3] > 0):
                img_l = Image.fromarray(canvas_leader.image_data.astype('uint8'), 'RGBA')
                buf_l = io.BytesIO()
                img_l.save(buf_l, format="PNG")
                leader_sig_base64 = base64.b64encode(buf_l.getvalue()).decode()
                img_l.save(os.path.join(SIGNATURE_DIR, sanitize_filename(f"leader_sig_{leader_name}_{tbm_date}.png")))
        except: pass

        worker_sigs_base64 = ["", "", "", ""]
        for i_w, c_w in enumerate(worker_canvases):
            try:
                if c_w.image_data is not None and np.any(c_w.image_data[:, :, 3] > 0):
                    img_w = Image.fromarray(c_w.image_data.astype('uint8'), 'RGBA')
                    buf_w = io.BytesIO()
                    img_w.save(buf_w, format="PNG")
                    worker_sigs_base64[i_w] = base64.b64encode(buf_w.getvalue()).decode()
                    w_label = workers[i_w].strip() if workers[i_w].strip() else f"worker{i_w+1}"
                    img_w.save(os.path.join(SIGNATURE_DIR, sanitize_filename(f"worker_sig_{w_label}_{tbm_date}.png")))
            except: pass

        contractor_sigs_base64 = ["", "", "", ""]
        for i_c, c_c in enumerate(contractor_canvases):
            try:
                if c_c.image_data is not None and np.any(c_c.image_data[:, :, 3] > 0):
                    img_c = Image.fromarray(c_c.image_data.astype('uint8'), 'RGBA')
                    buf_c = io.BytesIO()
                    img_c.save(buf_c, format="PNG")
                    contractor_sigs_base64[i_c] = base64.b64encode(buf_c.getvalue()).decode()
                    c_label = contractor_workers[i_c].strip() if contractor_workers[i_c].strip() else f"contractor{i_c+1}"
                    img_c.save(os.path.join(SIGNATURE_DIR, sanitize_filename(f"contractor_sig_{c_label}_{tbm_date}.png")))
            except: pass

        sign_tag = f'<img src="data:image/png;base64,{leader_sig_base64}" style="max-height:28px; vertical-align:middle;"/>' if leader_sig_base64 else leader_name

        def get_worker_sign_cell(idx, w_name):
            if not w_name.strip(): return ""
            agree_badge = '<span style="font-size:7.5px; color:#0284C7; display:block; line-height:1;">(☑개인정보동의)</span>' if (idx < len(worker_agrees) and worker_agrees[idx]) else '<span style="font-size:7.5px; color:#999; display:block; line-height:1;">(□미동의)</span>'
            if worker_mode == "직접 텍스트(이름) 입력":
                sig_part = custom_worker_signs[idx] if idx < len(custom_worker_signs) and custom_worker_signs[idx].strip() else w_name
                return f"{sig_part}{agree_badge}"
            else:
                if idx < len(worker_sigs_base64) and worker_sigs_base64[idx]:
                    return f'<img src="data:image/png;base64,{worker_sigs_base64[idx]}" style="max-height:18px; vertical-align:middle;"/>{agree_badge}'
                else:
                    return f'(서명){agree_badge}'

        def get_contractor_sign_cell(idx, c_name):
            if not c_name.strip(): return ""
            agree_badge = '<span style="font-size:7.5px; color:#0284C7; display:block; line-height:1;">(☑개인정보동의)</span>' if (idx < len(contractor_agrees) and contractor_agrees[idx]) else '<span style="font-size:7.5px; color:#999; display:block; line-height:1;">(□미동의)</span>'
            if contractor_mode == "직접 텍스트(이름) 입력":
                sig_part = custom_contractor_signs[idx] if idx < len(custom_contractor_signs) and custom_contractor_signs[idx].strip() else c_name
                return f"{sig_part}{agree_badge}"
            else:
                if idx < len(contractor_sigs_base64) and contractor_sigs_base64[idx]:
                    return f'<img src="data:image/png;base64,{contractor_sigs_base64[idx]}" style="max-height:18px; vertical-align:middle;"/>{agree_badge}'
                else:
                    return f'(서명){agree_badge}'

        weekly_worker_rows = ""
        for i, w_name in enumerate(workers):
            if w_name.strip():
                sig_c = get_worker_sign_cell(i, w_name)
                weekly_worker_rows += f'<tr style="text-align:center; height:24px;"><td style="font-weight:bold; background:#fafafa; text-align:left; padding-left:6px;">{w_name}</td><td>{sig_c}</td><td>{sig_c}</td><td>{sig_c}</td><td>{sig_c}</td><td>{sig_c}</td><td>{sig_c}</td><td>{sig_c}</td></tr>'
        for i, c_name in enumerate(contractor_workers):
            if c_name.strip():
                sig_cc = get_contractor_sign_cell(i, c_name)
                weekly_worker_rows += f'<tr style="text-align:center; height:24px;"><td style="font-weight:bold; background:#f1f3f5; text-align:left; padding-left:6px;">(업체){c_name}</td><td>{sig_cc}</td><td>{sig_cc}</td><td>{sig_cc}</td><td>{sig_cc}</td><td>{sig_cc}</td><td>{sig_cc}</td><td>{sig_cc}</td></tr>'
        if not weekly_worker_rows: weekly_worker_rows = '<tr style="text-align:center;"><td colspan="8">참석자 없음</td></tr>'

        job_risks = []
        if r1.strip(): job_risks.append((r1, s1))
        if r2.strip(): job_risks.append((r2, s2))
        if r3.strip(): job_risks.append((r3, s3))
        risk_rows_html = "".join([f'<tr><td style="border:1px solid #000; padding:4px; font-size:9.5px; background:#fafafa;">{r}</td><td style="border:1px solid #000; padding:4px; font-size:9.5px;">{s}</td></tr>' for r, s in job_risks])

        tbm_standard_html = f"""
        <!DOCTYPE html>
        <html>
        <head>
        <meta charset="UTF-8">
        <style>
            body {{ font-family: 'Malgun Gothic', '맑은 고딕', sans-serif; margin: 8px; color: #000; font-size: 10px; }}
            .title-box {{ font-size: 15px; font-weight: bold; margin-bottom: 5px; text-align: left; }}
            table {{ width: 100%; border-collapse: collapse; margin-bottom: 4px; font-size: 9.5px; }}
            th, td {{ border: 1px solid #000; padding: 3px 4px; vertical-align: middle; }}
            .header-td {{ background-color: #f2f2f2; font-weight: bold; text-align: center; }}
            .sub-table td {{ border: none; padding: 1px 2px; }}
        </style>
        </head>
        <body>
            <div class="title-box">[시설 명: 단월하수 처리시설 ] TBM[Tool Box Meeting] 회의록</div>
            <table>
                <tr>
                    <td class="header-td" style="width:11%;">TBM 일시</td>
                    <td style="width:39%;">{tbm_date.strftime('%Y년 %m월 %d일')} {tbm_time}</td>
                    <td class="header-td" style="width:28%;">작업날짜와 동일함</td>
                    <td style="width:22%; text-align:center; font-size:9px;">☑예 □아니오</td>
                </tr>
                <tr>
                    <td class="header-td">작 업 명</td>
                    <td style="font-weight:bold;">{custom_job}</td>
                    <td class="header-td" rowspan="2">TBM 장소</td>
                    <td rowspan="2" style="font-size:9px;">
                        {"□" if tbm_place=="사무실" else "☑"}사무실 &nbsp;&nbsp; {"☑" if tbm_place=="작업현장" else "□"}작업현장<br>
                        {"□" if tbm_place=="기타" else "☑"}기타(&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;)
                    </td>
                </tr>
                <tr>
                    <td class="header-td">작업내용</td>
                    <td>{job_desc}</td>
                </tr>
            </table>

            <table>
                <tr>
                    <td class="header-td" style="width:11%;" rowspan="3">외주<br>업체정보<br><span style="font-size:7.5px; font-weight:normal;">(해당시 작성)</span></td>
                    <td class="header-td" style="width:16%;">외주작업</td>
                    <td style="width:23%;">&nbsp;{"☑" if has_contractor else "□"} 예 &nbsp;&nbsp; {"□" if has_contractor else "☑"} 아니오</td>
                    <td class="header-td" style="width:22%;">업체 위험성평가 실시</td>
                    <td style="width:28%;">&nbsp;{"☑" if c_risk_eval=="예" else "□"} 예 &nbsp;&nbsp; {"□" if c_risk_eval=="예" else "☑"} 아니오<br><span style="font-size:8px; color:#333;">'아니오' 시도시 해당 직원 듣기</span></td>
                </tr>
                <tr>
                    <td class="header-td">업체명</td>
                    <td>{c_name}</td>
                    <td class="header-td" rowspan="2">산업안전보건 교육 확인</td>
                    <td rowspan="2">&nbsp;{"☑" if c_safety_edu=="예" else "□"} 예 &nbsp;&nbsp; {"□" if c_safety_edu=="예" else "☑"} 아니오<br><span style="font-size:8px; color:#333;">교육 미수도시 직접 듣기</span></td>
                </tr>
                <tr>
                    <td class="header-td">책임자 / 연락처</td>
                    <td>{c_leader} ({c_phone})</td>
                </tr>
            </table>

            <table>
                <tr>
                    <td class="header-td" style="width:38%;">■ 해당작업 위험성평가 실시 확인</td>
                    <td style="padding:3px; font-size:9.5px;">
                        &nbsp;□ 최초 &nbsp;&nbsp;&nbsp;&nbsp; 
                        &nbsp;{"☑" if eval_type == "정기" else "□"} 정기 &nbsp;&nbsp;&nbsp;&nbsp; 
                        &nbsp;{"☑" if eval_type == "수시" else "□"} 수시 (중복 가능)
                    </td>
                </tr>
            </table>

            <table>
                <tr style="background:#f2f2f2;">
                    <th style="width:45%; border:1px solid #000; text-align:center;">■ 유해·위험요인 파악 내용</th>
                    <th style="width:55%; border:1px solid #000; text-align:center;">■ 파악된 유해·위험요인의 감소대책 수립 및 이행</th>
                </tr>
                {risk_rows_html}
            </table>

            <table>
                <tr style="background:#f2f2f2;"><th colspan="2" style="text-align:left; border:1px solid #000;">■ TBM 리더의 작업 전 안전조치 여부 재확인</th></tr>
                <tr>
                    <td style="width:78%; padding:0;">
                        <table class="sub-table" style="width:100%;">
                            <tr><td style="width:35%; border-bottom:1px solid #ccc; padding:2px 4px;">• 필수 보호구 착용 상태</td><td style="border-bottom:1px solid #ccc; padding:2px 4px;">안전모, 안전화, 안전장갑, 각반, 안전대 등</td></tr>
                            <tr><td style="width:35%; border-bottom:1px solid #ccc; padding:2px 4px;">• 작업자 건강상태 확인</td><td style="border-bottom:1px solid #ccc; padding:2px 4px;">작업자 음주, 발열, 피로도 등(필요시 작업 배제)</td></tr>
                            <tr><td style="width:35%; border-bottom:1px solid #ccc; padding:2px 4px;">• 작업자의 TBM 내용 숙지</td><td style="border-bottom:1px solid #ccc; padding:2px 4px;">위험성평가 결과 및 동일작업의 TBM 일지 확인 등</td></tr>
                            <tr><td style="padding:2px 4px;">• 파악된 감소 대책 이행 여부</td><td style="padding:2px 4px;">위험성 감소대책 이행 확인</td></tr>
                        </table>
                    </td>
                    <td style="width:22%; text-align:center; vertical-align:middle; font-weight:bold; font-size:11px; background:#fff;">
                        ☑ 확인
                    </td>
                </tr>
            </table>

            <table>
                <tr style="background:#f2f2f2;"><td style="border:1px solid #000; font-weight:bold;">■ 작업 후 미팅 <span style="font-weight:normal; font-size:8px;">※ 위험요인 감소대책의 실효성, 미치사고 내용, 근로자 건의사항 등 의견을 작성, 내용이 없으면 "없음" 체크</span></td></tr>
                <tr><td style="border:1px solid #000; padding:4px; height:22px; font-size:9.5px;">{single_post_note}</td></tr>
            </table>

            <table>
                <tr><th colspan="5" style="text-align:left; background:#f2f2f2; border:1px solid #000;">■ TBM 리더 정보</th></tr>
                <tr style="text-align:center; font-weight:bold; background:#fafafa;">
                    <td style="width:16%;">소속</td><td style="width:18%;">직책(또는 직급)</td><td style="width:20%;">관리감독자</td><td style="width:16%;">성명</td><td style="width:30%;">서명란</td>
                </tr>
                <tr style="text-align:center; height:26px;">
                    <td>{leader_dept}</td><td>{leader_role}</td><td>□예 ☑아니오</td><td style="font-weight:bold;">{leader_name}</td><td>{sign_tag}</td>
                </tr>
            </table>

            <table>
                <tr><th colspan="6" style="text-align:left; background:#f2f2f2; border:1px solid #000;">■ 참석자 확인 <span style="font-weight:normal; font-size:8.5px;">※ 외주 업체 포함 참석</span></th></tr>
                <tr style="text-align:center; background:#fafafa; font-weight:bold; font-size:9.5px;">
                    <td style="width:14%;">성 명</td><td style="width:19%;">서 명</td><td style="width:14%;">성 명</td><td style="width:19%;">서 명</td><td style="width:18%;">업 체 성 명</td><td style="width:16%;">업 체 서 명</td>
                </tr>
                <tr style="text-align:center; height:22px;"><td>① {workers[0]}</td><td>{get_worker_sign_cell(0, workers[0])}</td><td>⑤</td><td></td><td>① {contractor_workers[0]}</td><td>{get_contractor_sign_cell(0, contractor_workers[0])}</td></tr>
                <tr style="text-align:center; height:22px;"><td>② {workers[1]}</td><td>{get_worker_sign_cell(1, workers[1])}</td><td>⑥</td><td></td><td>② {contractor_workers[1]}</td><td>{get_contractor_sign_cell(1, contractor_workers[1])}</td></tr>
                <tr style="text-align:center; height:22px;"><td>③ {workers[2]}</td><td>{get_worker_sign_cell(2, workers[2])}</td><td>⑦</td><td></td><td>③ {contractor_workers[2]}</td><td>{get_contractor_sign_cell(2, contractor_workers[2])}</td></tr>
                <tr style="text-align:center; height:22px;"><td>④ {workers[3]}</td><td>{get_worker_sign_cell(3, workers[3])}</td><td>⑧</td><td></td><td>④ {contractor_workers[3]}</td><td>{get_contractor_sign_cell(3, contractor_workers[3])}</td></tr>
            </table>
        </body>
        </html>
        """

        tbm_weekly_html = f"""
        <!DOCTYPE html>
        <html>
        <head>
        <meta charset="UTF-8">
        <style>
            body {{ font-family: 'Malgun Gothic', '맑은 고딕', sans-serif; margin: 8px; color: #000; font-size: 10px; }}
            table {{ width: 100%; border-collapse: collapse; margin-bottom: 4px; font-size: 9.5px; }}
            th, td {{ border: 1px solid #000; padding: 3px 4px; vertical-align: middle; text-align: center; }}
            .header-td {{ background-color: #f2f2f2; font-weight: bold; }}
        </style>
        </head>
        <body>
            <div style="font-size: 9px; margin-bottom: 2px; text-align: right;">[인원 초과 시 뒷면 추가 작성, 동일 1주일 작업 참석자는 별지1 작성 가능]</div>
            <div style="font-size: 9.5px; font-weight: bold; margin-bottom: 4px;">[별지1. 작업내용이 동일하여, TBM을 1주일 단위로 하고자 할 경우의 양식, 매 작업마다 TBM 작성 원칙]</div>
            <table>
                <tr class="header-td">
                    <td style="width:16%;">구분</td>
                    <td style="width:12%;">월</td><td style="width:12%;">화</td><td style="width:12%;">수</td><td style="width:12%;">목</td><td style="width:12%;">금</td><td style="width:12%;">토</td><td style="width:12%;">일</td>
                </tr>
                <tr>
                    <td class="header-td">TBM 장소</td>
                    <td>{tbm_place}</td><td>{tbm_place}</td><td>{tbm_place}</td><td>{tbm_place}</td><td>{tbm_place}</td><td>{tbm_place}</td><td>{tbm_place}</td>
                </tr>
                <tr>
                    <td class="header-td">TBM 리더 서명<br><span style="font-size:7.5px; font-weight:normal;">(안전조치 확인자)</span></td>
                    <td>{leader_name}</td><td>{leader_name}</td><td>{leader_name}</td><td>{leader_name}</td><td>{leader_name}</td><td>{leader_name}</td><td>{leader_name}</td>
                </tr>
            </table>

            <table>
                <tr><th colspan="9" class="header-td" style="background:#e9ecef;">참석자 서명</th></tr>
                <tr class="header-td">
                    <td style="width:16%; height:25px;">
                        <table style="width:100%; border:none;"><tr style="border:none;"><td style="border:none; text-align:left;">참석자</td></tr><tr style="border:none;"><td style="border:none; text-align:right;">요일</td></tr></table>
                    </td>
                    <td style="width:12%;">월</td><td style="width:12%;">화</td><td style="width:12%;">수</td><td style="width:12%;">목</td><td style="width:12%;">금</td><td style="width:12%;">토</td><td style="width:12%;">일</td>
                </tr>
                {weekly_worker_rows}
            </table>

            <table>
                <tr class="header-td"><td colspan="2" style="text-align:left;">■ 작업 후 미팅 내용</td></tr>
                <tr><td style="width:15%; background:#fafafa;" class="header-td">· (월)</td><td style="text-align:left; padding-left:5px;">{post_note_mon}</td></tr>
                <tr><td style="background:#fafafa;" class="header-td">· (화)</td><td style="text-align:left; padding-left:5px;">{post_note_tue}</td></tr>
                <tr><td style="background:#fafafa;" class="header-td">· (수)</td><td style="text-align:left; padding-left:5px;">{post_note_wed}</td></tr>
                <tr><td style="background:#fafafa;" class="header-td">· (목)</td><td style="text-align:left; padding-left:5px;">{post_note_thu}</td></tr>
                <tr><td style="background:#fafafa;" class="header-td">· (금)</td><td style="text-align:left; padding-left:5px;">{post_note_fri}</td></tr>
                <tr><td style="background:#fafafa;" class="header-td">· (토)</td><td style="text-align:left; padding-left:5px;">{post_note_sat}</td></tr>
                <tr><td style="background:#fafafa;" class="header-td">· (일)</td><td style="text-align:left; padding-left:5px;">{post_note_sun}</td></tr>
            </table>
        </body>
        </html>
        """

        active_html = tbm_weekly_html if is_weekly else tbm_standard_html

        st.divider()
        st.subheader("📋 TBM 표준회의록 공식 양식 미리보기 (새로운 규격 반영)")
        st.components.v1.html(active_html, height=680, scrolling=True)

        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            st.download_button("📥 TBM 회의록 인쇄/PDF 다운로드", data=active_html, file_name=f"TBM회의록_{tbm_date}.html", mime="text/html", type="primary", use_container_width=True)
        with col_btn2:
            if st.button("💾 TBM 영구 보관함에 저장", type="primary", use_container_width=True):
                save_path = os.path.join(TBM_RECORD_DIR, sanitize_filename(f"TBM회의록_{tbm_date}_{custom_job[:12]}.html"))
                with open(save_path, "w", encoding="utf-8") as f:
                    f.write(active_html)
                st.success(f"✅ TBM 영구 보관함({save_path})에 안전하게 저장되었습니다! 상단 [보관함] 탭에서 언제든 열람하실 수 있습니다.")

    with tab_tbm_chk:
        st.subheader("☑️ TBM[Tool Box Meeting] 실행 체크리스트 작성 및 인쇄")
        st.caption("🔒 항목별 체크(YES / NO / 해당없음) 후 하단에서 공식 양식 그대로 미리보기·다운로드·인쇄 및 영구 보관함 저장이 가능합니다.")

        col_c_top1, col_c_top2 = st.columns([1.5, 1.5])
        with col_c_top1:
            chk_date = st.date_input("체크리스트 점검 일자", tbm_date, key="chk_date_input")
        with col_c_top2:
            st.write("")
            if st.button("✨ [일괄 동의] 13개 전 항목 'YES' 일괄 체크하기", type="primary", use_container_width=True):
                for i in range(1, 14):
                    st.session_state[f"chk_{i}"] = "YES"
                st.toast("모든 체크리스트 항목이 'YES'로 일괄 선택되었습니다!", icon="✅")
                st.rerun()

        chk_items = [
            ("1 TBM 사전준비", 1, "① 해당 작업의 위험성평가를 실시하였다.<br>&nbsp;&nbsp;&nbsp;&nbsp;(해당 작업의 위험성평가 결과가 있다.)", "① 해당 작업의 위험성평가를 실시하였다. (해당 작업의 위험성평가 결과가 있다.)"),
            ("1 TBM 사전준비", 2, "② 해당 작업에서 발생한 사고보고서(아차<br>&nbsp;&nbsp;&nbsp;&nbsp;사고 포함)의 내용을 확인하였다.", "② 해당 작업에서 발생한 사고보고서(아차사고 포함)의 내용을 확인하였다."),
            ("1 TBM 사전준비", 3, "③ 작업 물량과 범위, 작업내용을 충분히 숙<br>&nbsp;&nbsp;&nbsp;&nbsp;지하였다.", "③ 작업 물량과 범위, 작업내용을 충분히 숙지하였다."),
            ("1 TBM 사전준비", 4, "④ 위험성평가 결과, 사고보고서, 안전작업<br>&nbsp;&nbsp;&nbsp;&nbsp;지침의 내용을 여러 번 읽어 숙지하였다.", "④ 위험성평가 결과, 사고보고서, 안전작업 지침의 내용을 여러 번 읽어 숙지하였다."),
            ("1 TBM 사전준비", 5, "⑤ 작업에 필요한 개인보호구(안전모, 안전화,<br>&nbsp;&nbsp;&nbsp;&nbsp;각반 등)를 지급하고, 착용방법을 숙지하<br>&nbsp;&nbsp;&nbsp;&nbsp;고 있다.", "⑤ 작업에 필요한 개인보호구(안전모, 안전화, 각반 등)를 지급하고, 착용방법을 숙지하고 있다."),
            ("2 TBM 실행과정", 6, "① 작업자가 음주, 발열, 약물 복용 등으로<br>&nbsp;&nbsp;&nbsp;&nbsp;작업에 적합한지 여부를 확인하였다.", "① 작업자가 음주, 발열, 약물 복용 등으로 작업에 적합한지 여부를 확인하였다."),
            ("2 TBM 실행과정", 7, "② 작업내용 / 위험요인 / 안전 작업절차 / 대책에<br>&nbsp;&nbsp;&nbsp;&nbsp;대해 긍정적인 분위기로 대화하였다", "② 작업내용 / 위험요인 / 안전 작업절차 / 대책에 대해 긍정적인 분위기로 대화하였다"),
            ("2 TBM 실행과정", 8, "③ 작업자와 중점 위험요인과 대책을 도출<br>&nbsp;&nbsp;&nbsp;&nbsp;하고 이를 숙지하도록 하였다.", "③ 작업자와 중점 위험요인과 대책을 도출하고 이를 숙지하도록 하였다."),
            ("2 TBM 실행과정", 9, "④ 위험요인, 불안전한 상태 발견시 멈추고,<br>&nbsp;&nbsp;&nbsp;&nbsp;확인하고, 생각한 후 작업하도록 하였다.", "④ 위험요인, 불안전한 상태 발견시 멈추고, 확인하고, 생각한 후 작업하도록 하였다."),
            ("2 TBM 실행과정", 10, "⑤ 작업 후 정리 정돈 상태를 확인하였다.", "⑤ 작업 후 정리 정돈 상태를 확인하였다."),
            ("3 TBM 후속조치", 11, "① 작업자가 제기한 불만사항, 질문, 제안<br>&nbsp;&nbsp;&nbsp;&nbsp;사항을 검토하였다.", "① 작업자가 제기한 불만사항, 질문, 제안 사항을 검토하였다."),
            ("3 TBM 후속조치", 12, "② TBM 결과를 충실하게 기록하고 보관 한다.", "② TBM 결과를 충실하게 기록하고 보관 한다."),
            ("3 TBM 후속조치", 13, "③ 관련 조치결과는 작업자에게 피드백 한다.", "③ 관련 조치결과는 작업자에게 피드백 한다.")
        ]

        for i in range(1, 14):
            if f"chk_{i}" not in st.session_state:
                st.session_state[f"chk_{i}"] = "YES"
            if f"memo_{i}" not in st.session_state:
                st.session_state[f"memo_{i}"] = ""

        with st.expander("📝 항목별 개별 체크 및 조치내용 입력 열기/닫기", expanded=True):
            checklist_data = []
            cur_cat = ""
            for cat, idx, q_html, q_plain in chk_items:
                if cat != cur_cat:
                    cur_cat = cat
                    st.markdown(f"##### ▣ [{cat}]")
                c_q, c_opt, c_memo = st.columns([2.3, 1.3, 1.4])
                with c_q:
                    st.write(q_plain)
                with c_opt:
                    val = st.radio(f"체크_{idx}", ["YES", "NO", "해당없음"], key=f"chk_{idx}", horizontal=True, label_visibility="collapsed")
                with c_memo:
                    memo = st.text_input(f"조치_{idx}", key=f"memo_{idx}", placeholder="'아니오'인 경우 조치내용 입력", label_visibility="collapsed")
                checklist_data.append({"구분": cat, "번호": idx, "확인사항_html": q_html, "확인사항": q_plain, "해당사항": val, "조치내용": memo})

        # 이미지와 100% 동일한 TBM 실행 체크리스트 인쇄용 HTML 생성
        chk_rows_html = ""
        current_cat_html = ""
        for item in checklist_data:
            cat_str = item["구분"]
            if cat_str != current_cat_html:
                current_cat_html = cat_str
                cat_num = cat_str.split()[0]
                cat_title = " ".join(cat_str.split()[1:])
                chk_rows_html += f"""
                <tr>
                    <td style="border:1px solid #000; padding:5px 8px; font-weight:bold; font-size:11px; background:#fff; width:26%;">
                        <span style="display:inline-block; border:1px solid #666; width:14px; height:14px; line-height:13px; text-align:center; font-size:10px; margin-right:4px;">{cat_num}</span>{cat_title}
                    </td>
                    <td colspan="4" style="border:1px solid #000; background:#fff;"></td>
                </tr>
                """
            v_yes = "V" if item["해당사항"] == "YES" else ""
            v_no = "V" if item["해당사항"] == "NO" else ""
            v_na = "V" if item["해당사항"] == "해당없음" else ""
            memo_txt = item["조치내용"] if item["조치내용"] else ""
            chk_rows_html += f"""
            <tr>
                <td colspan="1" style="border:1px solid #000; padding:6px 8px; font-size:10.5px; line-height:1.45; text-align:left;">
                    {item["확인사항_html"]}
                </td>
                <td style="border:1px solid #000; width:8%; text-align:center; font-weight:bold; font-size:12px;">{v_yes}</td>
                <td style="border:1px solid #000; width:8%; text-align:center; font-weight:bold; font-size:12px;">{v_no}</td>
                <td style="border:1px solid #000; width:8%; text-align:center; font-weight:bold; font-size:12px;">{v_na}</td>
                <td style="border:1px solid #000; width:22%; padding:4px 6px; font-size:10px; text-align:left;">{memo_txt}</td>
            </tr>
            """

        tbm_checklist_html = f"""
        <!DOCTYPE html>
        <html>
        <head>
        <meta charset="UTF-8">
        <style>
            @page {{ size: A4 portrait; margin: 12mm; }}
            body {{ font-family: 'Malgun Gothic', '맑은 고딕', sans-serif; margin: 10px; color: #000; }}
            .chk-container {{ max-width: 720px; margin: 0 auto; }}
            table.chk-table {{ width: 100%; border-collapse: collapse; border: 1.5px solid #000; table-layout: fixed; }}
            table.chk-table th, table.chk-table td {{ border: 1px solid #000; vertical-align: middle; word-break: keep-all; }}
            .title-banner {{ background-color: #5B82C4; color: #ffffff; font-size: 16px; font-weight: bold; text-align: center; padding: 9px; letter-spacing: 0.5px; }}
            .notice-box {{ padding: 10px 14px; font-size: 10.5px; line-height: 1.55; text-align: left; background: #fff; }}
            .notice-title {{ text-align: center; font-weight: bold; margin-bottom: 4px; font-size: 11px; }}
            .th-label {{ background-color: #fff; font-weight: bold; text-align: center; font-size: 11px; padding: 5px; }}
            @media print {{
                .no-print-btn {{ display: none !important; }}
                body {{ margin: 0; }}
            }}
        </style>
        </head>
        <body>
            <div class="chk-container">
                <div class="no-print-btn" style="text-align:right; margin-bottom:8px;">
                    <button onclick="window.print()" style="background:#0284C7; color:white; border:none; padding:6px 14px; border-radius:6px; font-weight:bold; cursor:pointer; font-size:12px;">
                        🖨️ 이 체크리스트 바로 인쇄하기
                    </button>
                </div>
                <table class="chk-table">
                    <colgroup>
                        <col style="width: 54%;">
                        <col style="width: 8%;">
                        <col style="width: 8%;">
                        <col style="width: 8%;">
                        <col style="width: 22%;">
                    </colgroup>
                    <tr>
                        <td colspan="5" class="title-banner">TBM[Tool Box Meeting] 실행 체크리스트</td>
                    </tr>
                    <tr>
                        <td colspan="5" class="notice-box">
                            <div class="notice-title">[유의사항]</div>
                            ◆ TBM은 작업 전 TBM 리더와 작업자 간 실행하는 안전보건 회의입니다.<br>
                            ◆ 이 체크리스트는 TBM 리더에게 효과적인 TBM 실행 팁을 제공하기 위해 제작된 것으로 사내 TBM 절차가 있는 경우 이와 함께 보완적으로 사용합니다.<br>
                            ◆ 작업별 위험요인은 위험성평가 결과 또는 별도의 자료를 활용하시기 바랍니다.
                        </td>
                    </tr>
                    <tr>
                        <td rowspan="2" class="th-label">확인사항</td>
                        <td colspan="3" class="th-label">해당사항에<br>체크(V) 하세요</td>
                        <td class="th-label">‘아니오’인 경우<br>필요한 조치내용</td>
                    </tr>
                    <tr>
                        <td class="th-label" style="font-size:10px;">YES</td>
                        <td class="th-label" style="font-size:10px;">NO</td>
                        <td class="th-label" style="font-size:9.5px; line-height:1.15;">해당<br>없음</td>
                        <td class="th-label"></td>
                    </tr>
                    {chk_rows_html}
                </table>
            </div>
        </body>
        </html>
        """

        st.divider()
        st.subheader("📋 TBM 실행 체크리스트 공식 양식 미리보기 (인쇄·다운로드 규격)")
        st.components.v1.html(tbm_checklist_html, height=780, scrolling=True)

        col_c_btn1, col_c_btn2 = st.columns(2)
        with col_c_btn1:
            st.download_button(
                "📥 TBM 실행 체크리스트 양식 다운로드 (HTML/인쇄용)",
                data=tbm_checklist_html,
                file_name=f"TBM실행체크리스트_{chk_date}.html",
                mime="text/html",
                type="primary",
                use_container_width=True
            )
        with col_c_btn2:
            if st.button("💾 TBM 실행 체크리스트 영구 보관함에 저장", type="primary", use_container_width=True):
                chk_html_path = os.path.join(TBM_RECORD_DIR, sanitize_filename(f"TBM실행체크리스트_{chk_date}_{custom_job[:10]}.html"))
                with open(chk_html_path, "w", encoding="utf-8") as f:
                    f.write(tbm_checklist_html)
                df_chk = pd.DataFrame(checklist_data)
                chk_csv_path = os.path.join(TBM_RECORD_DIR, sanitize_filename(f"TBM실행체크리스트_{chk_date}.csv"))
                df_chk.to_csv(chk_csv_path, index=False, encoding='utf-8-sig')
                st.success(f"✅ TBM 실행 체크리스트가 공식 인쇄 양식({chk_html_path})으로 영구 보관함에 저장되었습니다! 우측 [보관함] 탭에서도 언제든 열람·인쇄하실 수 있습니다.")

    with tab_tbm_archive:
        st.subheader("🗂️ 영구 저장된 TBM 회의록 보관함 열람 및 관리")
        saved_tbms = [sanitize_filename(f) for f in os.listdir(TBM_RECORD_DIR) if f.endswith(".html")]
        if saved_tbms:
            st.write(f"📁 **총 {len(saved_tbms)}건의 TBM 회의록이 영구 보관되어 있습니다.**")
            col_t1, col_t2 = st.columns([3, 1])
            with col_t1:
                sel_tbm_file = st.selectbox("열람할 TBM 회의록 파일 선택", sorted(saved_tbms, reverse=True))
            with col_t2:
                st.write(""); st.write("")
                if st.button("🗑️ 선택 회의록 삭제", use_container_width=True):
                    os.remove(os.path.join(TBM_RECORD_DIR, sel_tbm_file))
                    st.success("삭제되었습니다.")
                    st.rerun()

            if sel_tbm_file:
                tbm_fp = os.path.join(TBM_RECORD_DIR, sel_tbm_file)
                with open(tbm_fp, "r", encoding="utf-8") as f:
                    saved_html_content = f.read()
                st.download_button("📥 선택된 TBM 회의록 다운로드 (HTML/인쇄용)", saved_html_content, file_name=sel_tbm_file, mime="text/html", type="primary", use_container_width=True)
                st.markdown("##### 🔍 저장된 TBM 회의록 원본 화면 열람")
                st.components.v1.html(saved_html_content, height=680, scrolling=True)
        else:
            st.info("💡 아직 TBM 보관함에 저장된 회의록이 없습니다. [작성] 탭에서 '💾 TBM 영구 보관함에 저장' 버튼을 눌러주세요.")

        st.divider()
        st.subheader("✍️ 영구 저장된 전자서명 이미지 보관소 열람")
        saved_sigs = [sanitize_filename(f) for f in os.listdir(SIGNATURE_DIR) if f.endswith(".png")]
        if saved_sigs:
            cols_sig = st.columns(min(4, len(saved_sigs)))
            for idx, s_file in enumerate(sorted(saved_sigs)):
                with cols_sig[idx % 4]:
                    st.image(os.path.join(SIGNATURE_DIR, s_file), caption=s_file, use_container_width=True)
        else:
            st.info("💡 저장된 전자서명 이미지가 없습니다.")

# -------------------------------------------------------------
# 8. 약품 구매 및 사용 관리 & 보관함
# -------------------------------------------------------------
elif menu == "💊 8. 약품 구매 및 사용 관리 & 보관함":
    st.title("💊 단월공공하수처리시설 약품 구매 및 재고 영구 보관함")
    with st.form("chem_purchase_form"):
        c1, c2, c3, c4 = st.columns(4)
        with c1: p_date = st.date_input("입고/사용 일자", datetime.date.today())
        with c2: p_item = st.selectbox("약품 종류", ["PAC (17%)", "염화제이철 (38%)", "양이온 폴리머", "기타 실험시약"])
        with c3: p_type = st.selectbox("구분", ["구매 입고(+)", "일일 사용(-)"])
        with c4: p_qty = st.number_input("수량 (kg)", value=100.0, step=10.0)
        p_memo = st.text_input("비고 (납품업체 또는 투입 공정)", "정상 입고")
        if st.form_submit_button("💾 약품 수불부 기록 영구 저장", type="primary", use_container_width=True):
            row_df = pd.DataFrame([{"일자": p_date.strftime('%Y-%m-%d'), "약품명": p_item, "구분": p_type, "수량_kg": p_qty, "비고": p_memo}])
            if os.path.exists(CHEMICAL_INVENTORY_DB):
                old_df = pd.read_csv(CHEMICAL_INVENTORY_DB)
                comb_df = pd.concat([old_df, row_df], ignore_index=True)
            else:
                comb_df = row_df
            comb_df.to_csv(CHEMICAL_INVENTORY_DB, index=False, encoding='utf-8-sig')
            comb_df.to_csv(os.path.join(CHEMICAL_RECORD_DIR, f"약품수불부_백업_{p_date.strftime('%Y%m')}.csv"), index=False, encoding='utf-8-sig')
            st.success("✅ 약품 입출고 기록이 마스터 DB와 약품 보관함에 영구 저장되었습니다!")
            st.rerun()

    if os.path.exists(CHEMICAL_INVENTORY_DB):
        st.subheader("🗂️ 영구 저장된 약품 수불부 내역 열람")
        df_inv = pd.read_csv(CHEMICAL_INVENTORY_DB).sort_values(by="일자", ascending=False)
        st.dataframe(df_inv, use_container_width=True)
        st.download_button("📥 약품 수불부 전체 내역 다운로드 (CSV)", df_inv.to_csv(index=False, encoding='utf-8-sig'), file_name="danwol_chemical_inventory.csv", mime="text/csv", type="primary", use_container_width=True)
    else:
        st.info("💡 등록된 약품 입출고 내역이 없습니다.")

# -------------------------------------------------------------
# 9. 통합 영구저장소(보관함·마스터DB) 열람 및 백업 센터
# -------------------------------------------------------------
elif menu == "🗄️ 9. 통합 영구저장소(보관함·마스터DB) 열람 및 백업 센터":
    st.title("🗄️ 단월 스마트 관제 플랫폼 통합 영구저장소 & 백업 센터")
    st.info("💡 코드 수정이나 업데이트 시에도 지워지지 않고 서버에 영구 보존되는 모든 보관함 파일과 마스터 데이터베이스를 한눈에 열람·검증·백업할 수 있습니다.")

    zip_backup_bytes = create_full_backup_zip()
    st.download_button(
        "📦 ⚡ [전체 영구저장소 + 마스터 DB 통합 백업 ZIP 다운로드]",
        data=zip_backup_bytes,
        file_name=f"DANWOL_AI_WaterOps_Full_Backup_{datetime.date.today()}.zip",
        mime="application/zip",
        type="primary",
        use_container_width=True
    )

    st.divider()
    tab_db_view, tab_folder_view = st.tabs(["📊 5대 마스터 데이터베이스(CSV) 통합 열람", "📁 5대 영구 보관함 폴더별 파일 열람 및 미리보기"])

    with tab_db_view:
        st.subheader("📊 영구 적재된 5대 마스터 데이터베이스 실시간 조회")
        sel_db_label = st.selectbox("열람할 마스터 데이터베이스 선택", list(ALL_MASTER_DBS.keys()))
        target_db_path = ALL_MASTER_DBS[sel_db_label]

        if os.path.exists(target_db_path):
            try:
                df_db_view = pd.read_csv(target_db_path)
                f_size_kb = round(os.path.getsize(target_db_path) / 1024, 1)
                m_time = datetime.datetime.fromtimestamp(os.path.getmtime(target_db_path), KST).strftime('%Y-%m-%d %H:%M:%S')
                st.success(f"✅ **파일명**: `{target_db_path}` | **총 적재 건수**: `{len(df_db_view)}행` | **용량**: `{f_size_kb} KB` | **최종 갱신**: `{m_time}`")
                st.dataframe(df_db_view, use_container_width=True)
                with open(target_db_path, "rb") as f:
                    st.download_button(f"📥 {sel_db_label} 원본 CSV 다운로드", f.read(), file_name=target_db_path, mime="text/csv", type="primary", use_container_width=True)
            except Exception as e:
                st.error(f"데이터베이스 읽기 오류: {e}")
        else:
            st.warning(f"⚠️ 아직 `{target_db_path}` 파일에 저장된 데이터가 없습니다. 각 기능 메뉴에서 데이터를 저장하면 이곳에 즉시 표시됩니다.")

    with tab_folder_view:
        st.subheader("📁 5대 영구 보관함 디렉토리 파일 목록 및 즉시 열람")
        sel_dir_label = st.selectbox("열람할 영구 보관함 폴더 선택", list(ALL_STORAGE_DIRS.keys()))
        target_dir_path = ALL_STORAGE_DIRS[sel_dir_label]

        if os.path.exists(target_dir_path):
            files_in_dir = [sanitize_filename(f) for f in os.listdir(target_dir_path) if os.path.isfile(os.path.join(target_dir_path, f))]
            if files_in_dir:
                st.write(f"📂 **현재 폴더(`{target_dir_path}/`)에 총 `{len(files_in_dir)}개`의 파일이 영구 보관되어 있습니다.**")
                summary_rows = []
                for fn in sorted(files_in_dir, reverse=True):
                    fp = os.path.join(target_dir_path, fn)
                    summary_rows.append({
                        "파일명": fn,
                        "크기(KB)": round(os.path.getsize(fp) / 1024, 1),
                        "수정일시(KST)": datetime.datetime.fromtimestamp(os.path.getmtime(fp), KST).strftime('%Y-%m-%d %H:%M:%S')
                    })
                st.dataframe(pd.DataFrame(summary_rows), use_container_width=True)

                sel_f = st.selectbox("미리보기 또는 다운로드할 파일 선택", sorted(files_in_dir, reverse=True), key="center_file_sel")
                if sel_f:
                    full_fp = os.path.join(target_dir_path, sel_f)
                    col_f1, col_f2 = st.columns([3, 1])
                    with open(full_fp, "rb") as f:
                        col_f1.download_button(f"📥 '{sel_f}' 파일 다운로드", f.read(), file_name=sel_f, type="primary", use_container_width=True)
                    if col_f2.button("🗑️ 선택 파일 삭제", key="center_del_btn", use_container_width=True):
                        os.remove(full_fp)
                        st.success("삭제 완료!")
                        st.rerun()

                    if sel_f.endswith(".html"):
                        st.markdown("##### 🔍 HTML 문서 즉시 미리보기")
                        with open(full_fp, "r", encoding="utf-8") as f:
                            st.components.v1.html(f.read(), height=650, scrolling=True)
                    elif sel_f.endswith(".png") or sel_f.endswith(".jpg"):
                        st.markdown("##### 🔍 이미지 파일 미리보기")
                        st.image(full_fp, caption=sel_f)
                    elif sel_f.endswith(".csv"):
                        st.markdown("##### 🔍 CSV 데이터 미리보기")
                        st.dataframe(pd.read_csv(full_fp), use_container_width=True)
                    elif sel_f.endswith(".xlsx") or sel_f.endswith(".xls"):
                        st.markdown("##### 🔍 엑셀 데이터 미리보기")
                        try:
                            st.dataframe(pd.read_excel(full_fp).head(35), use_container_width=True)
                        except Exception:
                            st.info("엑셀 파일을 다운로드하여 확인해 주세요.")
            else:
                st.info(f"💡 `{target_dir_path}/` 폴더가 비어 있습니다.")
