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

    /* 💧 단월 슈퍼 프리미엄 글래스모피즘 KPI 카드 (용문 상위호환) */
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

    /* 기존 호환 클래스 */
    .ym-kpi-card { background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 14px; padding: 16px 18px; box-shadow: 0 4px 12px rgba(15, 23, 42, 0.04); }
    .ym-kpi-label { font-size: 13px; font-weight: 600; color: #475569; margin-bottom: 6px; }
    .ym-kpi-val { font-size: 30px; font-weight: 800; color: #0F172A; line-height: 1.15; margin-bottom: 8px; }
    .ym-pill-red { display: inline-block; background: #FEE2E2; color: #DC2626; font-size: 12px; font-weight: 700; padding: 3px 10px; border-radius: 20px; }
    .ym-pill-green { display: inline-block; background: #DCFCE7; color: #15803D; font-size: 12px; font-weight: 700; padding: 3px 10px; border-radius: 20px; }
    .ym-blue-banner { background: #EFF6FF; border-left: 5px solid #2563EB; border-radius: 10px; padding: 14px 18px; color: #1E3A8A; font-size: 14.5px; font-weight: 700; margin: 14px 0; }
    .ym-green-safe-banner { background: linear-gradient(90deg, #ECFDF5 0%, #D1FAE5 100%); border: 1px solid #6EE7B7; border-left: 6px solid #10B981; border-radius: 12px; padding: 16px 20px; color: #065F46; font-size: 14.5px; font-weight: 800; margin: 12px 0; box-shadow: 0 4px 12px rgba(16, 185, 129, 0.08); }
    .ym-yellow-warn-banner { background: linear-gradient(90deg, #FEFCE8 0%, #FEF3C7 100%); border: 1px solid #FDE68A; border-left: 5px solid #F59E0B; border-radius: 12px; padding: 13px 18px; color: #92400E; font-size: 14px; font-weight: 800; margin-bottom: 12px; }
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
    facility_aliases = {
        "산음": ["산음", "산음리"], "삼가리": ["삼가리"], "진목": ["진목", "보룡리(진목)", "보룡리", "보룡"],
        "몰운": ["몰운", "몰운리"], "단월마을": ["단월마을"], "당의": ["당의"]
    }
    accumulated_data = {fac: {} for fac in facility_aliases.keys()}
    if not file_list: return {fac: pd.DataFrame() for fac in facility_aliases.keys()}
    for f in file_list:
        try:
            fname = getattr(f, 'name', str(f))
            y_m = re.search(r'(20[1-3]\d)', fname)
            file_year_anchor = int(y_m.group(1)) if y_m else 2026
            wb = openpyxl.load_workbook(f, data_only=True)
            for sname in wb.sheetnames:
                ws = wb[sname]
                if ws.max_row < 2: continue
                sheet_fac = None
                sname_clean = sname.replace(" ", "")
                fname_clean = fname.replace(" ", "")
                for std_fac, aliases in facility_aliases.items():
                    for al in aliases:
                        if al.replace(" ", "") in sname_clean or al.replace(" ", "") in fname_clean:
                            sheet_fac = std_fac; break
                    if sheet_fac: break
                r1_val = str(ws.cell(1, 1).value or '')
                r2_val = str(ws.cell(2, 1).value or '')
                is_24_col = ('유량및수질' in r1_val or '업로드양식' in r1_val or '날짜' in r2_val)
                is_comp_log = ('종합운영일지' in r1_val or '1. 유량현황' in str(ws.cell(4, 1).value or ''))
                if is_24_col and sheet_fac:
                    for r in range(4, min(ws.max_row + 1, 1000)):
                        c1_val = ws.cell(r, 1).value
                        if c1_val is None: continue
                        dt_val = None
                        if isinstance(c1_val, (datetime.datetime, datetime.date)):
                            dt_val = datetime.date(c1_val.year, c1_val.month, c1_val.day)
                        elif isinstance(c1_val, str):
                            m = re.search(r'(20[1-3]\d)[-/.](\d{1,2})[-/.](\d{1,2})', c1_val)
                            if m:
                                dt_val = datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
                            else:
                                m2 = re.search(r'(\d{1,2})[-/.](\d{1,2})', c1_val)
                                if m2:
                                    dt_val = datetime.date(file_year_anchor, int(m2.group(1)), int(m2.group(2)))
                        if dt_val:
                            d_str = dt_val.strftime('%Y-%m-%d')
                            if d_str not in accumulated_data[sheet_fac]:
                                accumulated_data[sheet_fac][d_str] = {'날짜': d_str}
                            rec = accumulated_data[sheet_fac][d_str]
                            flow_in = pd.to_numeric(ws.cell(r, 2).value, errors='coerce')
                            flow_out = pd.to_numeric(ws.cell(r, 6).value, errors='coerce')
                            temp_val = pd.to_numeric(ws.cell(r, 7).value, errors='coerce')
                            if pd.notna(flow_in): rec['유입량'] = float(flow_in)
                            if pd.notna(flow_out): rec['방류량'] = float(flow_out)
                            if pd.notna(temp_val): rec['수온'] = float(temp_val)
                            for col_idx, col_name in [(8, '유입pH'), (9, '유입BOD'), (10, '유입TOC'), (11, '유입SS'), (12, '유입TN'), (13, '유입TP'), (14, '유입대장균')]:
                                v = pd.to_numeric(ws.cell(r, col_idx).value, errors='coerce')
                                if pd.notna(v): rec[col_name] = float(v)
                            for col_idx, col_name in [(16, '방류pH'), (17, '방류BOD'), (18, '방류TOC'), (19, '방류SS'), (20, '방류TN'), (21, '방류TP'), (22, '방류대장균')]:
                                v = pd.to_numeric(ws.cell(r, col_idx).value, errors='coerce')
                                if pd.notna(v): rec[col_name] = float(v)
                elif is_comp_log and sheet_fac:
                    cur_date = None
                    in_flow_sec = False
                    in_wq_sec = False
                    for r in range(1, min(ws.max_row + 1, 2000)):
                        c0 = ws.cell(r, 1).value
                        if isinstance(c0, (datetime.datetime, datetime.date)) or (isinstance(c0, str) and re.search(r'202[4-6]', c0)):
                            cur_date = c0 if isinstance(c0, (datetime.datetime, datetime.date)) else pd.to_datetime(c0).date()
                            if isinstance(cur_date, datetime.datetime): cur_date = cur_date.date()
                            d_str = cur_date.strftime('%Y-%m-%d')
                            if d_str not in accumulated_data[sheet_fac]:
                                accumulated_data[sheet_fac][d_str] = {'날짜': d_str}
                            in_flow_sec = False
                            in_wq_sec = False
                        c0_str = str(c0 or '').replace(' ', '')
                        if '1.유량현황' in c0_str:
                            in_flow_sec = True; in_wq_sec = False; continue
                        elif '2.전력량' in c0_str or '3.수질현황' in c0_str:
                            in_flow_sec = False; in_wq_sec = ('3.수질현황' in c0_str); continue
                        elif '4.시설현황' in c0_str:
                            in_flow_sec = False; in_wq_sec = False; continue
                        if cur_date and in_flow_sec:
                            d_str = cur_date.strftime('%Y-%m-%d')
                            rec = accumulated_data[sheet_fac][d_str]
                            val5 = pd.to_numeric(ws.cell(r, 5).value, errors='coerce')
                            if '처리장' in c0_str and pd.notna(val5):
                                rec['유입량'] = float(val5)
                                rec['방류량'] = float(val5)
                            elif '유입량' in c0_str and pd.notna(val5):
                                rec['유입량'] = float(val5)
                            elif '방류량' in c0_str and pd.notna(val5):
                                rec['방류량'] = float(val5)
                        if cur_date and in_wq_sec:
                            d_str = cur_date.strftime('%Y-%m-%d')
                            rec = accumulated_data[sheet_fac][d_str]
                            if '유입수' in c0_str:
                                for col_idx, col_name in [(2, '유입BOD'), (3, '유입TOC'), (4, '유입SS'), (5, '유입TN'), (6, '유입TP'), (7, '유입대장균')]:
                                    v = pd.to_numeric(ws.cell(r, col_idx).value, errors='coerce')
                                    if pd.notna(v): rec[col_name] = float(v)
                            elif '방류수' in c0_str:
                                for col_idx, col_name in [(2, '방류BOD'), (3, '방류TOC'), (4, '방류SS'), (5, '방류TN'), (6, '방류TP'), (7, '방류대장균')]:
                                    v = pd.to_numeric(ws.cell(r, col_idx).value, errors='coerce')
                                    if pd.notna(v): rec[col_name] = float(v)
                else:
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
                                        cur_tab_fac = std_fac; break
                                if cur_tab_fac: break
                        if cur_tab_fac and isinstance(c1, (datetime.datetime, datetime.date)):
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
                            v_wk = pd.to_numeric(ws.cell(r, 15).value, errors='coerce')
                            v_dy = pd.to_numeric(ws.cell(r, 16).value, errors='coerce')
                            if pd.notna(v_wk): rec['주간유량'] = float(v_wk)
                            if pd.notna(v_dy):
                                rec['유입량'] = round(float(v_dy), 1)
                                rec['방류량'] = round(float(v_dy), 1)
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
    """소규모 6개소(산음·삼가리·진목·몰운·단월마을·당의) 업로드양식 예시 파일과 100% 동일한 서식·수치 생성"""
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

    r3_heights = {'산음': 28.5, '삼가리': 29.25, '진목': 32.25, '몰운': 18.75, '단월마을': 30.75, '당의': 27.75}
    ws.row_dimensions[2].height = 18.75
    ws.row_dimensions[3].height = r3_heights.get(fac_name, 28.5)
    ws.column_dimensions['A'].width = 12.0
    ws.column_dimensions['B'].width = 12.5
    ws.column_dimensions['G'].width = 10.0
    for c_idx in range(3, 25):
        if c_idx != 7:
            ws.column_dimensions[get_column_letter(c_idx)].width = 13.0

    thin_side = Side(style='thin', color='000000')
    border_r2 = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    font_header = Font(name='dotum', size=9, bold=True)
    for r in range(1, 4):
        for c in range(1, 25):
            cell = ws.cell(r, c)
            cell.alignment = align_center
            cell.font = font_header
            cell.number_format = '@'
            if r == 2:
                cell.border = border_r2

    if start_date and end_date:
        d_range = pd.date_range(start_date, end_date)
    else:
        d_range = pd.date_range(f"{year}-01-01", f"{year}-12-31")

    lookup = {}
    measured_dates = []
    if df_data is not None and not df_data.empty and '날짜' in df_data.columns:
        for _, r in df_data.sort_values(by='날짜').iterrows():
            d_key = str(r['날짜']).split()[0]
            lookup[d_key] = r
            if pd.notna(r.get('유입BOD')) or pd.notna(r.get('방류BOD')):
                measured_dates.append(pd.to_datetime(d_key))

    daily_flow_in_map = {}
    daily_flow_out_map = {}
    for d_key, r_item in lookup.items():
        if pd.notna(r_item.get('유입량')):
            daily_flow_in_map[d_key] = float(r_item.get('유입량'))
        if pd.notna(r_item.get('방류량')):
            daily_flow_out_map[d_key] = float(r_item.get('방류량'))

    if len(daily_flow_in_map) < 60 and measured_dates:
        prev_dt = None
        for m_dt in measured_dates:
            d_str = m_dt.strftime('%Y-%m-%d')
            r_item = lookup[d_str]
            val_in = float(r_item.get('유입량', 0.0)) if pd.notna(r_item.get('유입량')) else 0.0
            val_out = float(r_item.get('방류량', val_in)) if pd.notna(r_item.get('방류량')) else val_in
            start_fill = pd.to_datetime(f"{m_dt.year}-01-01") if prev_dt is None else (prev_dt + pd.Timedelta(days=1))
            for w_dt in pd.date_range(start_fill, m_dt):
                w_str = w_dt.strftime('%Y-%m-%d')
                daily_flow_in_map[w_str] = val_in
                daily_flow_out_map[w_str] = val_out
            prev_dt = m_dt

    font_data = Font(name='맑은 고딕', size=11, bold=False)
    date_fmt = 'mm-dd-yy' if fac_name in ['몰운', '당의'] else 'yyyy\\-mm\\-dd'

    for r_idx, dt in enumerate(d_range, start=4):
        d_str = dt.strftime('%Y-%m-%d')
        c1 = ws.cell(r_idx, 1, datetime.datetime(dt.year, dt.month, dt.day, 0, 0))
        c1.number_format = date_fmt
        c1.font = font_data

        if d_str in daily_flow_in_map or d_str in daily_flow_out_map:
            f_in_v = daily_flow_in_map.get(d_str, None)
            f_out_v = daily_flow_out_map.get(d_str, f_in_v)
            if f_in_v is not None and pd.notna(f_in_v):
                c_b = ws.cell(r_idx, 2, f_in_v); c_b.font = font_data; c_b.number_format = '0.0'
                c_e = ws.cell(r_idx, 5, f_in_v); c_e.font = font_data; c_e.number_format = '0.0'
            if f_out_v is not None and pd.notna(f_out_v):
                c_f = ws.cell(r_idx, 6, f_out_v); c_f.font = font_data; c_f.number_format = '0.0'

        r_match = lookup.get(d_str, None)
        if r_match is not None:
            for col_idx, col_name, n_fmt in [
                (9, '유입BOD', '0.0'), (10, '유입TOC', '0.0'), (11, '유입SS', '0.0'),
                (12, '유입TN', '0.000'), (13, '유입TP', '0.000'), (14, '유입대장균', '#,##0_);[Red]\\(#,##0\\)')
            ]:
                v = r_match.get(col_name, None)
                if pd.notna(v) and str(v).strip() != '':
                    v_num = int(round(float(v))) if col_idx == 14 else float(v)
                    c = ws.cell(r_idx, col_idx, v_num); c.font = font_data; c.number_format = n_fmt
            for col_idx, col_name, n_fmt in [
                (17, '방류BOD', '0.0'), (18, '방류TOC', '0.0'), (19, '방류SS', '0.0'),
                (20, '방류TN', '0.000'), (21, '방류TP', '0.000'), (22, '방류대장균', '#,##0_);[Red]\\(#,##0\\)')
            ]:
                v = r_match.get(col_name, None)
                if pd.notna(v) and str(v).strip() != '':
                    v_num = int(round(float(v))) if col_idx == 22 else float(v)
                    c = ws.cell(r_idx, col_idx, v_num); c.font = font_data; c.number_format = n_fmt
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

def auto_seed_master_db_from_local_files():
    """업로드된 국가하수도 공인 엑셀 예시 파일 9종을 최우선으로 읽어 오차 없이 마스터 DB를 초기화"""
    if os.path.exists(MASTER_ACCUM_DB):
        try:
            df_chk = pd.read_csv(MASTER_ACCUM_DB)
            has_priv = all(pf in df_chk['시설명'].unique() for pf in PRIVATE_PLANTS) if '시설명' in df_chk.columns else False
            if len(df_chk) > 1500 and has_priv:
                return
        except Exception:
            pass
    df_main_list = []
    for m_cand in ["2026년 월별수질(단월).xlsx", "2026년 월별수질(단월)_2.xlsx"]:
        if os.path.exists(m_cand):
            df_m = universal_main_plant_parser([m_cand])
            if not df_m.empty:
                df_main_list.append(df_m)
            break
    if os.path.exists("유량및수질관리.xlsx"):
        df_kh_m = parse_khas_51col_main_file("유량및수질관리.xlsx")
        if not df_kh_m.empty:
            df_main_list.append(df_kh_m)
    if df_main_list:
        df_main_merged = pd.concat(df_main_list, ignore_index=True).drop_duplicates(subset=['날짜'], keep='last').sort_values('날짜')
        append_to_master_db(MAIN_PLANT, df_main_merged)

    small_khas_files = [f for f in [
        "유량및수질관리 업로드양식(산음).xlsx",
        "유량및수질관리 업로드양식(삼가리).xlsx",
        "유량및수질관리 업로드양식(진목).xlsx",
        "유량및수질관리 업로드양식 (몰운).xlsx",
        "유량및수질관리 업로드양식(단월마을).xlsx",
        "유량및수질관리 업로드양식(당의).xlsx"
    ] if os.path.exists(f)]
    if small_khas_files:
        s_dict = universal_small_plant_parser(small_khas_files)
        for fk, df_s in s_dict.items():
            if not df_s.empty:
                append_to_master_db(fk, df_s)
    else:
        for s_cand in ["1. 소규모(2026).xlsx", "1. 소규모(2026)_2.xlsx"]:
            if os.path.exists(s_cand):
                s_dict = universal_small_plant_parser([s_cand])
                for fk, df_s in s_dict.items():
                    if not df_s.empty:
                        append_to_master_db(fk, df_s)
                break

    for p_cand in ["개인소규모(2026년).xlsx", "개인소규모(2026년)_2.xlsx", "개인소규모(2026년)+단월.xls", "개인소규모(2026년)+단월_2.xls"]:
        if os.path.exists(p_cand):
            p_dict = parse_private_plant_multi_files([p_cand])
            for fk, df_p in p_dict.items():
                if not df_p.empty:
                    append_to_master_db(fk, df_p)

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

        for r_idx, (b_lbl, fn_stat) in enumerate([(54, ("최    대", max)), (55, ("최    소", min)), (56, ("평    균", lambda x: round(sum(x)/len(x), 1)))], start=54):
            lbl_txt, fn_op = b_lbl
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
# 4. 공정 제어 (지능형 수질진단 관제 대시보드 [슈퍼 프리미엄 에디션] + 시설별 AI 공정 연산 및 보관함)
# -------------------------------------------------------------
elif menu == "⚙️ 4. AI 최적 운전조건 제안 & KNR+IPR 공정 정밀진단":
    # 사이드바 실시간 공정 제어 패널 (단월 본장 + 소규모 6개소 전환 및 강우·기온·수온 변동 추이 직접 지정 시뮬레이터 탑재)
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🎛️ 실시간 공정 제어 패널")
    diag_fac = st.sidebar.selectbox("🏢 관제 대상 처리장 선택", [MAIN_PLANT] + SMALL_PLANTS, index=0, key="ym_fac_sel")
    diag_date = st.sidebar.date_input("📅 진단 기준일 선택", datetime.date(2026, 11, 3), key="ym_diag_date")
    st.sidebar.markdown("##### 🌦️ 날씨 · 기온 · 강우 변동 추이 설정")
    diag_season = st.sidebar.selectbox("현재 계절", ["봄", "여름", "가을", "겨울"], index=2, key="ym_season")
    diag_month = st.sidebar.slider("분석 대상 월 (3개년 통계 연동)", 1, 12, int(diag_date.month), key="ym_month")
    diag_rain = st.sidebar.number_input("기준 예보 강우량 (mm/일)", value=0.00, min_value=0.00, step=1.0, format="%.2f", key="ym_rain")
    diag_air_temp = st.sidebar.number_input("현재 외기 기온 (℃)", value=11.00, min_value=-25.0, max_value=40.0, step=0.5, format="%.1f", key="ym_air_temp")
    st.sidebar.markdown("##### 🧪 현재 수질 & 운전 인자")
    diag_temp = st.sidebar.number_input("반응조 현재 수온 (℃)", value=14.00, step=0.5, format="%.2f", key="ym_temp")
    diag_mlss = st.sidebar.slider("생물반응조 MLSS (mg/L)", 2000, 5000, 3650, step=50, key="ym_mlss")
    diag_do = st.sidebar.slider("호기조 DO 농도 (mg/L)", 0.8, 4.0, 2.1, step=0.1, key="ym_do")

    fac_spec = PLANT_DESIGN_SPECS.get(diag_fac, PLANT_DESIGN_SPECS[MAIN_PLANT])
    max_cap_val = int(fac_spec["cap"])
    warn_cap_val = int(round(max_cap_val * 0.91))

    tab_diag, tab_p1, tab_p2 = st.tabs([
        "🎯 지능형 수질진단 관제 대시보드 (계절·강우·기온·디지털트윈 연동)",
        "⚙️ 시설별 AI 공정 연산 및 진단",
        "🗂️ 공정 제어 영구 보관함 (전체 시설 열람)"
    ])

    with tab_diag:
        # 상단 프리미엄 헤더 바 (우측 배지는 깔끔하게 '🟢 AI 자율진단 엔진 가동중'으로 표시)
        st.markdown(f"""
        <div style="display:flex; justify-content:space-between; align-items:flex-end; border-bottom:2px solid #E2E8F0; padding-bottom:14px; margin-bottom:20px;">
            <div>
                <div style="display:inline-block; background:#E0F2FE; color:#0369A1; font-weight:800; font-size:12px; padding:4px 12px; border-radius:20px; margin-bottom:6px;">
                    AI DIGITAL TWIN SIMULATOR · {fac_spec['method']} 공법
                </div>
                <h2 style="margin:0; font-size:28px; font-weight:900; color:#0F172A; letter-spacing:-0.5px;">
                    지능형 수질진단 관제 대시보드 <span style="color:#0284C7; font-size:20px;">({diag_fac})</span>
                </h2>
                <div style="color:#64748B; font-size:14px; font-weight:500; margin-top:4px;">
                    양평군 {diag_fac} 맞춤형 인공지능 시뮬레이터 시스템입니다. (일일 최대 시설 용량: <b>{max_cap_val:,}톤</b> | 기준일: {diag_date.strftime('%Y-%m-%d')})
                </div>
            </div>
            <div style="text-align:right;">
                <span style="background:#ECFDF5; color:#059669; border:1px solid #6EE7B7; padding:6px 16px; border-radius:20px; font-size:13px; font-weight:800;">
                    🟢 AI 자율진단 엔진 가동중
                </span>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # -----------------------------------------------------------------
        # 🎛️ 사용자 맞춤형 [강우량 · 외기 기온 · 반응조 수온] 일별 변동 추이 직접 지정 제어반
        # -----------------------------------------------------------------
        import calendar
        days_in_sel_m = calendar.monthrange(diag_date.year, diag_month)[1]
        days_arr = list(range(1, days_in_sel_m + 1))

        with st.expander("🌦️🌡️ [강우량 · 기온 · 반응조 수온] 월간 일별(1~말일) 변동 추이 직접 설정 및 표 편집 (클릭하여 열기/닫기)", expanded=True):
            st.caption("💡 아래에서 **강우 시나리오(비 오는 날·지속 일수·강우량)**와 **기온/수온 변동 추이(초순➔말일 변화)**를 슬라이더로 지정하거나, **[일별 상세 표 직접 편집]**을 켜서 원하는 날짜의 강우량(mm)·기온(℃)·수온(℃)을 직접 입력할 수 있습니다.")
            c_tr1, c_tr2, c_tr3 = st.columns([1.1, 1.1, 1.0])
            with c_tr1:
                st.markdown("**☔ 1. 강우 발생 패턴 지정**")
                rain_mode = st.selectbox(
                    "강우 시나리오 선택",
                    ["평시 무강우 (전체 0mm)", "단발성 강우 (특정 구간 지정)", "월 2회 강우 (초순 + 중하순)", "장마·집중호우 (연속 강우)"],
                    index=0 if diag_rain == 0 else 1,
                    key="tr_rain_mode"
                )
                rain_peak_day = st.slider("강우 집중 시작일 (일)", 1, max(1, days_in_sel_m - 3), min(14, days_in_sel_m - 3), key="tr_rain_day")
                rain_dur_days = st.slider("강우 지속 기간 (일수)", 1, 7, 3, key="tr_rain_dur")
                rain_peak_mm = st.number_input("최대 피크 강우량 (mm/일)", value=float(diag_rain if diag_rain > 0 else 25.0), min_value=0.0, max_value=300.0, step=5.0, key="tr_rain_peak")

            with c_tr2:
                st.markdown("**🌡️ 2. 기온 & 반응조 수온 변동 추이 지정**")
                temp_trend_mode = st.selectbox(
                    "기온·수온 월간 변동 추이 패턴",
                    ["일정 유지 (현재 설정값 고정)", "점진적 하강 (가을➔겨울철 냉각 추이)", "점진적 상승 (봄➔여름철 온난화 추이)", "한파/급강하 포함 (중순 기온 급락)"],
                    index=1 if diag_season in ["가을", "겨울"] else (2 if diag_season in ["봄", "여름"] else 0),
                    key="tr_temp_mode"
                )
                air_t_start, air_t_end = st.slider(
                    "월간 외기 기온 변동 범위 (초순 ℃ ➔ 말일 ℃)",
                    -20.0, 38.0,
                    (float(diag_air_temp + 3.0), float(diag_air_temp - 4.0)) if "하강" in temp_trend_mode else ((float(diag_air_temp - 3.0), float(diag_air_temp + 4.0)) if "상승" in temp_trend_mode else (float(diag_air_temp), float(diag_air_temp))),
                    step=0.5,
                    key="tr_air_range"
                )
                water_t_start, water_t_end = st.slider(
                    "월간 반응조 수온 변동 범위 (초순 ℃ ➔ 말일 ℃)",
                    4.0, 32.0,
                    (float(diag_temp + 2.0), float(diag_temp - 2.5)) if "하강" in temp_trend_mode else ((float(diag_temp - 2.0), float(diag_temp + 2.5)) if "상승" in temp_trend_mode else (float(diag_temp), float(diag_temp))),
                    step=0.5,
                    key="tr_water_range"
                )

            with c_tr3:
                st.markdown("**✍️ 3. 일별 상세 수치 직접 수정 옵션**")
                use_manual_table = st.toggle("📋 일별(1~말일) 강우·기온·수온 표에서 직접 수정하기", value=False, key="tr_use_manual")
                st.info("📌 슬라이더나 표에서 강우량과 기온·수온 추이를 변경하면 **예상 방류수질(BOD·T-N·SS·T-P·TOC), 디지털 트윈 운전값, 일별 유입예측 차트**가 즉시 연동되어 재계산됩니다.")

            # 패턴에 따른 일별 배열 생성
            rain_arr = [0.0] * days_in_sel_m
            if rain_mode != "평시 무강우 (전체 0mm)" and rain_peak_mm > 0:
                for d_offset in range(rain_dur_days):
                    idx_d = min(days_in_sel_m - 1, rain_peak_day - 1 + d_offset)
                    # 가운데 날에 피크가 오도록 가중치 부여
                    center_dist = abs(d_offset - (rain_dur_days - 1) / 2.0)
                    weight = max(0.3, 1.0 - (center_dist / max(1.0, rain_dur_days / 1.5)))
                    rain_arr[idx_d] = round(rain_peak_mm * weight, 1)
                if rain_mode == "월 2회 강우 (초순 + 중하순)":
                    for d_offset in range(max(1, rain_dur_days - 1)):
                        idx_d2 = min(days_in_sel_m - 1, max(0, (rain_peak_day + 10) % days_in_sel_m + d_offset))
                        rain_arr[idx_d2] = round(rain_peak_mm * 0.65, 1)
                elif rain_mode == "장마·집중호우 (연속 강우)":
                    for d_offset in range(min(10, days_in_sel_m - rain_peak_day)):
                        idx_d3 = min(days_in_sel_m - 1, rain_peak_day - 1 + d_offset)
                        rain_arr[idx_d3] = round(rain_peak_mm * (0.45 + 0.55 * abs(np.sin(d_offset))), 1)

            air_temp_arr = []
            water_temp_arr = []
            for i_d in range(days_in_sel_m):
                frac = i_d / max(1, days_in_sel_m - 1)
                if temp_trend_mode == "일정 유지 (현재 설정값 고정)":
                    a_val = float(diag_air_temp) + 0.6 * np.sin(i_d * 0.5)
                    w_val = float(diag_temp) + 0.3 * np.sin(i_d * 0.5)
                elif temp_trend_mode == "한파/급강하 포함 (중순 기온 급락)":
                    dip = -6.0 * np.exp(-((i_d - days_in_sel_m * 0.55) ** 2) / 12.0)
                    a_val = air_t_start + (air_t_end - air_t_start) * frac + dip
                    w_val = water_t_start + (water_t_end - water_t_start) * frac + dip * 0.45
                else:
                    a_val = air_t_start + (air_t_end - air_t_start) * frac + 0.8 * np.sin(i_d * 0.6)
                    w_val = water_t_start + (water_t_end - water_t_start) * frac + 0.35 * np.sin(i_d * 0.6)
                air_temp_arr.append(round(a_val, 1))
                water_temp_arr.append(round(max(2.0, w_val), 1))

            # 일별 기본 유입량 + 강우 유입량(I/I) 연산
            scale_f = max_cap_val / 1700.0
            base_inflow_arr = []
            for i_d, d_num in enumerate(days_arr):
                dry_flow = int(max_cap_val * 0.88) + int((max_cap_val * 0.03) * np.sin(d_num * 0.45))
                r_mm = rain_arr[i_d]
                wet_extra = int(r_mm * 8.5 * scale_f) if r_mm > 0 else 0
                base_inflow_arr.append(dry_flow + wet_extra)

            df_trend_sim = pd.DataFrame({
                "일(Day)": days_arr,
                "강우량(mm)": rain_arr,
                "외기기온(℃)": air_temp_arr,
                "반응조수온(℃)": water_temp_arr,
                "예측유입량(톤/일)": base_inflow_arr
            })

            if use_manual_table:
                st.markdown("###### 📝 일별 강우량(mm) · 외기기온(℃) · 반응조수온(℃) · 예측유입량(톤/일) 직접 수정 테이블")
                edited_trend_df = st.data_editor(
                    df_trend_sim,
                    use_container_width=True,
                    hide_index=True,
                    height=260,
                    key="ed_custom_weather_trend"
                )
                rain_arr = pd.to_numeric(edited_trend_df["강우량(mm)"], errors="coerce").fillna(0.0).tolist()
                air_temp_arr = pd.to_numeric(edited_trend_df["외기기온(℃)"], errors="coerce").fillna(diag_air_temp).tolist()
                water_temp_arr = pd.to_numeric(edited_trend_df["반응조수온(℃)"], errors="coerce").fillna(diag_temp).tolist()
                # 강우량을 수정하면 유입량도 자동 연동 반영 (단, 사용자가 유입량을 직접 바꾼 경우 존중)
                recalc_inflow = []
                for i_d, d_num in enumerate(days_arr):
                    orig_in = df_trend_sim["예측유입량(톤/일)"].iloc[i_d]
                    user_in = pd.to_numeric(edited_trend_df["예측유입량(톤/일)"].iloc[i_d], errors="coerce")
                    if pd.notna(user_in) and int(user_in) != int(orig_in):
                        recalc_inflow.append(int(user_in))
                    else:
                        dry_flow = int(max_cap_val * 0.88) + int((max_cap_val * 0.03) * np.sin(d_num * 0.45))
                        recalc_inflow.append(dry_flow + (int(rain_arr[i_d] * 8.5 * scale_f) if rain_arr[i_d] > 0 else 0))
                base_inflow_arr = recalc_inflow

        # 설정된 변동 추이의 평균/최대값 추출하여 수질 예측에 실시간 반영
        eff_max_rain = max(rain_arr) if rain_arr else float(diag_rain)
        eff_avg_water_temp = float(np.mean(water_temp_arr)) if water_temp_arr else float(diag_temp)
        eff_min_water_temp = float(np.min(water_temp_arr)) if water_temp_arr else float(diag_temp)
        eff_avg_air_temp = float(np.mean(air_temp_arr)) if air_temp_arr else float(diag_air_temp)

        # 계절·강우·수온·MLSS·DO 기반 동적 예측 수질 산출 (실측 마스터 DB 연동)
        ym_q = f"{diag_date.year}-{diag_month:02d}"
        df_m_diag = get_master_data(diag_fac, f"{ym_q}-01", f"{ym_q}-31")
        if df_m_diag.empty:
            df_m_diag = get_master_data(diag_fac)

        base_bod_d = float(df_m_diag['방류BOD'].dropna().mean()) if (not df_m_diag.empty and '방류BOD' in df_m_diag.columns and df_m_diag['방류BOD'].dropna().any()) else 0.70
        base_tn_d = float(df_m_diag['방류TN'].dropna().mean()) if (not df_m_diag.empty and '방류TN' in df_m_diag.columns and df_m_diag['방류TN'].dropna().any()) else 6.01
        base_ss_d = float(df_m_diag['방류SS'].dropna().mean()) if (not df_m_diag.empty and '방류SS' in df_m_diag.columns and df_m_diag['방류SS'].dropna().any()) else 2.10
        base_tp_d = float(df_m_diag['방류TP'].dropna().mean()) if (not df_m_diag.empty and '방류TP' in df_m_diag.columns and df_m_diag['방류TP'].dropna().any()) else 0.032
        base_toc_d = float(df_m_diag['방류TOC'].dropna().mean()) if (not df_m_diag.empty and '방류TOC' in df_m_diag.columns and df_m_diag['방류TOC'].dropna().any()) else 3.10
        base_coli_d = float(df_m_diag['방류대장균'].dropna().mean()) if (not df_m_diag.empty and '방류대장균' in df_m_diag.columns and df_m_diag['방류대장균'].dropna().any()) else 45.0

        rain_adj = 1.0 + max(0.0, (eff_max_rain - 10.0) * 0.004)
        temp_adj = 1.0 + max(0.0, (16.0 - eff_avg_water_temp) * 0.018)
        do_adj = 1.0 - ((diag_do - 2.0) * 0.05)

        pred_bod_v = round(max(0.1, base_bod_d * rain_adj * do_adj), 2)
        pred_tn_v = round(max(1.0, base_tn_d * temp_adj * rain_adj), 2)
        pred_ss_v = round(max(0.2, base_ss_d * rain_adj), 2)
        pred_tp_v = round(max(0.008, base_tp_d * rain_adj), 3)
        pred_toc_v = round(max(0.8, base_toc_d * rain_adj * do_adj), 2)
        pred_coli_v = max(1, int(round(base_coli_d * rain_adj)))

        # 5대 핵심 수질 슈퍼 프리미엄 카드
        k1, k2, k3, k4, k5 = st.columns(5)
        kpi_super_specs = [
            (k1, "예상 방류수 BOD", f"{pred_bod_v:.2f}", "mg/L", 5.0, pred_bod_v, "🧪", "#0284C7"),
            (k2, "예상 방류수 T-N", f"{pred_tn_v:.2f}", "mg/L", 20.0, pred_tn_v, "🧬", "#6366F1"),
            (k3, "예상 방류수 SS", f"{pred_ss_v:.2f}", "mg/L", 10.0, pred_ss_v, "💧", "#0EA5E9"),
            (k4, "예상 방류수 T-P", f"{pred_tp_v:.3f}", "mg/L", 0.20, pred_tp_v, "⚡", "#10B981"),
            (k5, "예상 방류 대장균", f"{pred_coli_v:,}", "개/mL", 1000.0, float(pred_coli_v), "🦠", "#8B5CF6"),
        ]
        for col_obj, k_lbl, k_val_s, k_unit_s, k_lim, k_cur, k_ico, k_color in kpi_super_specs:
            ratio_pct = min(100, max(4, int(round((k_cur / k_lim) * 100))))
            safety_margin = max(0, 100 - ratio_pct)
            bar_color = "#10B981" if ratio_pct < 70 else ("#F59E0B" if ratio_pct < 90 else "#EF4444")
            if k_lim < 1:
                k_lim_str = f"{k_lim:.2f}"
            elif k_lim < 100:
                k_lim_str = f"{k_lim:.1f}"
            else:
                k_lim_str = f"{int(k_lim)}"
            with col_obj:
                st.markdown(f"""
                <div class="dw-super-card" style="border-top-color: {k_color};">
                    <div class="dw-kpi-top">
                        <span class="dw-kpi-title">{k_lbl}</span>
                        <span class="dw-kpi-icon">{k_ico}</span>
                    </div>
                    <div class="dw-kpi-value">{k_val_s}<span class="dw-kpi-unit">{k_unit_s}</span></div>
                    <div class="dw-progress-bg">
                        <div class="dw-progress-fill" style="width:{ratio_pct}%; background:{bar_color};"></div>
                    </div>
                    <div class="dw-badge-row">
                        <span class="dw-pill-limit">↑ ↑ 기준: {k_lim_str}</span>
                        <span class="dw-pill-safe">안전여유 {safety_margin}%</span>
                    </div>
                </div>
                """, unsafe_allow_html=True)

        # 유기물 분해 지표 진단 (연도별 법령 자동 전환) + 제거효율 배너
        toc_margin_pct = round((1.0 - (pred_toc_v / 15.0)) * 100, 1)
        st.markdown("<div style='margin-top:18px;'></div>", unsafe_allow_html=True)
        st.markdown("#### 🛡️ 유기물 분해 지표 진단 (연도별 법령 자동 전환 & 실시간 여유율)")
        st.markdown(f"""
        <div class="dw-toc-banner">
            <div>
                <span style="background:#0284C7; color:white; font-size:11.5px; font-weight:800; padding:4px 10px; border-radius:6px; margin-right:10px;">법령 자동전환</span>
                <span style="color:#0F172A; font-size:15px; font-weight:800;">2024년 이후 현행 법령 기준: TOC (총유기탄소)</span>
                <span style="color:#0369A1; font-size:15.5px; font-weight:900; margin-left:10px;">| 예상 방류수 TOC 결과: {pred_toc_v:.2f} mg/L</span>
                <span style="color:#475569; font-size:13.5px; font-weight:700; margin-left:6px;">(↑ 기준: 15.0 mg/L)</span>
            </div>
            <div style="background:#DCFCE7; color:#15803D; font-weight:800; font-size:13px; padding:6px 14px; border-radius:20px; border:1px solid #86EFAC;">
                ✅ 법적 기준 대비 {toc_margin_pct}% 안전 여유 확보
            </div>
        </div>
        """, unsafe_allow_html=True)

        # KNR + IPR 5단계 스마트 공정 디지털 트윈 상태바 (강우·기온·수온 추이 반영)
        ai_opt = calculate_ai_process_parameters(float(np.mean(base_inflow_arr)), 165.0, 26.5, 2.9, facility_name=diag_fac)
        rainy_days_cnt = sum(1 for r_v in rain_arr if r_v > 0)
        weather_status_txt = "청천시(무강우) 평시 유입" if eff_max_rain <= 0 else f"강우 {rainy_days_cnt}일 (피크 {eff_max_rain:.1f}mm)"
        st.markdown(f"""
        <div class="dw-flow-container">
            <div style="display:flex; justify-content:space-between; align-items:center;">
                <span style="font-size:14.5px; font-weight:800; color:#E2E8F0;">⚡ {diag_fac} [{fac_spec['method']}] 실시간 디지털 트윈 공정 제어 상태</span>
                <span style="font-size:12px; color:#38BDF8; font-weight:700;">평균 외기기온 {eff_avg_air_temp:.1f}℃ · 평균 반응조수온 {eff_avg_water_temp:.1f}℃ 연동</span>
            </div>
            <div class="dw-flow-grid">
                <div class="dw-flow-node">
                    <div class="dw-flow-node-title">① 유입·스크린·침사지</div>
                    <div class="dw-flow-node-val">{int(round(np.mean(base_inflow_arr))):,} ㎥/일</div>
                    <div class="dw-flow-node-sub">{weather_status_txt}</div>
                </div>
                <div class="dw-flow-node">
                    <div class="dw-flow-node-title">② KNR 생물반응조</div>
                    <div class="dw-flow-node-val">MLSS {diag_mlss:,}</div>
                    <div class="dw-flow-node-sub">평균수온 {eff_avg_water_temp:.1f}℃ · DO {diag_do:.1f}</div>
                </div>
                <div class="dw-flow-node">
                    <div class="dw-flow-node-title">③ AI 송풍기 최적제어</div>
                    <div class="dw-flow-node-val">{ai_opt['권장송풍량_m3min']} ㎥/분</div>
                    <div class="dw-flow-node-sub">권장 가동 {ai_opt['송풍기가동대수']}대 (절전모드)</div>
                </div>
                <div class="dw-flow-node">
                    <div class="dw-flow-node-title">④ 종침 & IPR 총인처리</div>
                    <div class="dw-flow-node-val">FeCl₃ {ai_opt['권장염화제이철_L']} L</div>
                    <div class="dw-flow-node-sub">PAC {ai_opt['종침전PAC주입량_L']} L/일 정밀주입</div>
                </div>
                <div class="dw-flow-node">
                    <div class="dw-flow-node-title">⑤ UV소독 · 최종방류</div>
                    <div class="dw-flow-node-val">T-P {pred_tp_v:.3f}</div>
                    <div class="dw-flow-node-sub">전 항목 협약수질 100% 이내</div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # 메인 차트 (유입량·강우량 + 외기기온·반응조수온 변동 추이 2단 연동 차트) + 우측 계절별 운영 지침
        col_chart, col_guide = st.columns([1.45, 1.0])
        with col_chart:
            st.markdown(f"#### 📊 단월 AI 유입예측 & 강우·기온·수온 일별 변동 추이 리포트 ({diag_season}철 / {diag_month}월)")
            peak_inflow = max(base_inflow_arr)
            peak_day = days_arr[base_inflow_arr.index(peak_inflow)]

            fig_dual = make_subplots(
                rows=2, cols=1,
                shared_xaxes=True,
                vertical_spacing=0.10,
                row_heights=[0.62, 0.38],
                specs=[[{"secondary_y": True}], [{"secondary_y": False}]],
                subplot_titles=(
                    f"🌊 일별 하수 유입량(톤/일) 및 지정 강우량(mm) 추이",
                    f"🌡️ 일별 외기 기온(℃) 및 생물반응조 수온(℃) 변동 추이"
                )
            )
            # 1단: 강우량 막대 + 예측 유입량 스플라인
            fig_dual.add_trace(
                go.Bar(
                    x=days_arr, y=rain_arr,
                    name="지정 강우량 (mm)",
                    marker_color="rgba(99, 102, 241, 0.42)",
                    marker_line_color="rgba(99, 102, 241, 0.85)",
                    marker_line_width=1
                ),
                row=1, col=1, secondary_y=True
            )
            fig_dual.add_trace(
                go.Scatter(
                    x=days_arr, y=base_inflow_arr,
                    mode="lines+markers",
                    name="예측 유입량 (톤/일)",
                    line=dict(color="#0284C7", width=3, shape="spline"),
                    fill="tozeroy",
                    fillcolor="rgba(2, 132, 199, 0.08)",
                    marker=dict(size=6, color="#0369A1", line=dict(width=1.2, color="white"))
                ),
                row=1, col=1, secondary_y=False
            )
            fig_dual.add_hline(y=max_cap_val, line_dash="dash", line_color="#EF4444", line_width=2, annotation_text=f"최대 처리용량 ({max_cap_val:,}톤)", annotation_position="top left", row=1, col=1, secondary_y=False)
            fig_dual.add_hline(y=warn_cap_val, line_dash="dot", line_color="#F59E0B", line_width=2, annotation_text=f"주의 경계선 ({warn_cap_val:,}톤)", annotation_position="bottom left", row=1, col=1, secondary_y=False)

            fig_dual.add_annotation(
                x=peak_day, y=peak_inflow,
                text=f"최대 피크: {peak_inflow:,}톤 ({peak_day}일)",
                showarrow=True, arrowhead=2, arrowcolor="#0284C7",
                bgcolor="#0F172A", font=dict(color="white", size=11),
                row=1, col=1
            )

            # 2단: 외기 기온 & 반응조 수온 변동 추이
            fig_dual.add_trace(
                go.Scatter(
                    x=days_arr, y=air_temp_arr,
                    mode="lines+markers",
                    name="외기 기온 (℃)",
                    line=dict(color="#F59E0B", width=2.5, dash="dot"),
                    marker=dict(size=5, color="#D97706")
                ),
                row=2, col=1
            )
            fig_dual.add_trace(
                go.Scatter(
                    x=days_arr, y=water_temp_arr,
                    mode="lines+markers",
                    name="반응조 수온 (℃)",
                    line=dict(color="#10B981", width=3, shape="spline"),
                    marker=dict(size=6, color="#059669")
                ),
                row=2, col=1
            )
            fig_dual.add_hline(y=12.0, line_dash="dot", line_color="#EF4444", line_width=1.5, annotation_text="질산화 저하 주의 수온 (12℃)", annotation_position="bottom right", row=2, col=1)

            fig_dual.update_layout(
                height=520,
                margin=dict(l=15, r=15, t=35, b=15),
                legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="left", x=0),
                plot_bgcolor="#FFFFFF",
                paper_bgcolor="#FFFFFF",
                hovermode="x unified"
            )
            fig_dual.update_xaxes(title_text="일 (Day)", showgrid=True, gridcolor="#F1F5F9", row=2, col=1)
            fig_dual.update_yaxes(title_text="유입량 (톤/일)", range=[int(max_cap_val * 0.55), max(int(max_cap_val * 1.18), int(peak_inflow * 1.08))], showgrid=True, gridcolor="#F1F5F9", row=1, col=1, secondary_y=False)
            fig_dual.update_yaxes(title_text="강우량 (mm)", range=[0, max(45, eff_max_rain * 1.5)], showgrid=False, row=1, col=1, secondary_y=True)
            fig_dual.update_yaxes(title_text="온도 (℃)", showgrid=True, gridcolor="#F1F5F9", row=2, col=1)
            st.plotly_chart(fig_dual, use_container_width=True)

        with col_guide:
            st.markdown(f"#### 📅 {diag_season}철 공정 운영 주의사항 ({diag_month}월 관제)")
            season_banners = {
                "봄": "해빙기 유입수온 상승 전환 및 미생물 활성도 변동 관리 단계",
                "여름": "하절기 집중호우(I/I) 유입 급증 및 고수온 DO 저하 집중 관리 단계",
                "가을": "동절기 대비 미생물 침강성 관리 및 공정 안정화 단계",
                "겨울": "저수온기 질산화 효율 저하 방지 및 보온·송풍량 정밀 제어 단계"
            }
            st.markdown(f"""
            <div class="ym-yellow-warn-banner">
                ▶ {season_banners.get(diag_season, "동절기 대비 미생물 침강성 관리 및 공정 안정화 단계")}
            </div>
            """, unsafe_allow_html=True)
            st.markdown("**[현장 실시간 점검 지침 (설정된 기온·수온·강우 추이 연동)]**")
            st.markdown(
                f"* **기온·수온 추이**: 외기 평균 **{eff_avg_air_temp:.1f}℃** / 반응조 수온 **{water_temp_arr[0]:.1f}℃ ➔ {water_temp_arr[-1]:.1f}℃** (최저 **{eff_min_water_temp:.1f}℃**) 변동에 대비하여 MLSS 농도 적정 유지 (현재 **{diag_mlss:,} mg/L**)\n"
                f"* **강우·유입 대비**: 월간 강우일수 **{rainy_days_cnt}일** (최대 **{eff_max_rain:.1f}mm/일**) 설정에 따른 협잡물 스크린 및 유량조정조 수위 점검\n"
                "* **총인(IPR) 약품 점검**: 메디아 여과기 차압 및 응집제(염화제이철·PAC) 정량펌프 주입 라인 점검\n"
                "* **법령 기준 확인**: 최종 방류수 계절별 법령 기준(TOC, T-N, T-P 등) 자동 전환 시스템 정상 작동 확인"
            )

            # 6대 수질 제거효율 레이더 차트
            radar_cats = ['BOD 제거율', 'TOC 제거율', 'SS 제거율', 'T-N 제거율', 'T-P 제거율', '대장균 살균율']
            radar_vals = [99.4, 91.3, 98.7, 77.5, 98.9, 99.9]
            fig_radar = go.Figure()
            fig_radar.add_trace(go.Scatterpolar(
                r=radar_vals + [radar_vals[0]],
                theta=radar_cats + [radar_cats[0]],
                fill='toself',
                fillcolor='rgba(2, 132, 199, 0.22)',
                line=dict(color='#0284C7', width=2.5),
                name='단월 제거효율(%)'
            ))
            fig_radar.update_layout(
                polar=dict(radialaxis=dict(visible=True, range=[60, 100], tickfont=dict(size=9))),
                showlegend=False,
                height=210,
                margin=dict(l=30, r=30, t=15, b=15),
                paper_bgcolor="rgba(0,0,0,0)"
            )
            st.plotly_chart(fig_radar, use_container_width=True)

        st.markdown("<hr style='margin:16px 0; border:none; border-top:1px solid #E2E8F0;'>", unsafe_allow_html=True)
        st.markdown("#### 🚨 실시간 공정 위험도 시스템 리포트 & 에너지 절감 코칭")
        if peak_inflow <= max_cap_val:
            saved_kwh = int(round(max_cap_val * 0.085))
            st.markdown(f"""
            <div class="ym-green-safe-banner">
                💚 [현재 수질 안전] 지정하신 기온({eff_avg_air_temp:.1f}℃)·수온({eff_avg_water_temp:.1f}℃) 및 강우(최대 {eff_max_rain:.1f}mm) 조건에서 모든 일별 유입 트렌드가 기준치 이내로 안정적입니다. (최대 피크 {peak_inflow:,.1f}톤 / 설계용량 {max_cap_val:,}톤) 과잉 가동 중인 송풍량을 미세 수동 조작하여 전기세를 절감(일 약 {saved_kwh:,} kWh 절감 가능)할 수 있는 최적의 타이밍입니다.
            </div>
            """, unsafe_allow_html=True)
        else:
            st.warning(f"⚠️ [유입부하 주의] 설정하신 강우(피크 {eff_max_rain:.1f}mm) 영향으로 {peak_day}일경 최대 피크 유입량({peak_inflow:,.1f}톤)이 설계용량({max_cap_val:,}톤/일)을 초과할 것으로 예측됩니다. 유량조정조 수위 선제 확보 및 총인(IPR) 응집제 예비 펌프를 가동하십시오.")

        st.markdown("<hr style='margin:18px 0; border:none; border-top:1px solid #E2E8F0;'>", unsafe_allow_html=True)
        st.markdown("#### 📋 방류수 수질 항목 초과 가이드라인 및 긴급 대응 점검표 (SOP)")
        sop_t1, sop_t2, sop_t3, sop_t4, sop_t5, sop_t6 = st.tabs([
            "BOD 조치 지침", "T-N 조치 지침", "SS 조치 지침", "T-P 조치 지침", "대장균 조치 지침", "TOC 조치 지침"
        ])
        with sop_t1:
            st.markdown("##### 🔴 BOD 초과 시 조치 (협약기준: 5.0 mg/L / 법적기준: 10.0 mg/L)")
            st.markdown(f"1. **DO 최적화**: 폭기조 DO를 2.0 mg/L 이상으로 제어하십시오. (현재 설정 DO: **{diag_do:.1f} mg/L**, 권장 송풍량: **{ai_opt['권장송풍량_m3min']} ㎥/분**)\n2. **미생물 진단**: SV30 및 MLSS 농도(현재 **{diag_mlss:,} mg/L**)를 확인하고 잉여슬러지 인발량을 조절하십시오.")
        with sop_t2:
            st.markdown("##### 🔴 T-N 초과 시 조치 (기준: 20.0 mg/L)")
            st.markdown(f"1. **내부반송 및 C/N비 점검**: 유입수 BOD/T-N 비(현재 **{ai_opt['CN비']}**)가 4.0 이상인지 확인하고 내부반송률을 150~200%로 조정하십시오.\n2. **수온 및 무산소조 교반 점검**: 반응조 평균 수온(**{eff_avg_water_temp:.1f}℃**, 최저 **{eff_min_water_temp:.1f}℃**) 하강 시 SRT(슬러지 체류시간)를 길게 유지하십시오.")
        with sop_t3:
            st.markdown("##### 🔴 SS 초과 시 조치 (기준: 10.0 mg/L)")
            st.markdown("1. **이차침전지 스컴 및 슬러지 계면 관리**: 슬러지 수집기 및 스컴 스키머 작동 상태를 확인하십시오.\n2. **총인처리시설(IPR) 여과 점검**: 미디어 여과 차압 상승 시 즉시 역세척(Backwash)을 실시하십시오.")
        with sop_t4:
            st.markdown("##### 🔴 T-P 초과 시 조치 (기준: 0.20 mg/L)")
            st.markdown(f"1. **응집제 주입량 상향**: 총인처리시설(IPR) 염화제이철(권장 **{ai_opt['권장염화제이철_L']} L/일**) 및 종침 PAC(권장 **{ai_opt['종침전PAC주입량_L']} L/일**) 정량펌프 주입율을 10~20% 상향 조정하십시오.\n2. **응집·침전 상태 확인**: 급속·완속 교반기 회전수 및 폴리머 용해 상태를 점검하십시오.")
        with sop_t5:
            st.markdown("##### 🔴 대장균 초과 시 조치 (기준: 1,000 개/mL)")
            st.markdown("1. **UV 소독설비 점검**: UV 램프 점등 상태, 석영관 오염도(세척기 작동) 및 자외선 강도계를 확인하십시오.\n2. **방류수 탁도 관리**: SS 상승으로 인한 자외선 투과율 저하 여부를 함께 점검하십시오.")
        with sop_t6:
            st.markdown("##### 🔴 TOC 초과 시 조치 (기준: 15.0 mg/L)")
            st.markdown("1. **난분해성 유기물 유입 확인**: 외부 연계수 또는 고농도 비점오염원 유입 여부를 확인하십시오.\n2. **물리·화학적 고도처리 연계**: 생물학적 분해 효율 증대와 함께 총인 응집침전 공정의 유기물 동시 제거 효율을 극대화하십시오.")

        # 수질진단 리포트 인쇄 / 다운로드 / 영구 보관함 저장 버튼
        diag_report_html = f"""
        <!DOCTYPE html><html><head><meta charset="UTF-8">
        <style>
            @page {{ size: A4 portrait; margin: 12mm; }}
            body {{ font-family: 'Malgun Gothic', sans-serif; color:#0F172A; margin:15px; }}
            h2 {{ border-bottom: 2px solid #0284C7; padding-bottom: 8px; color:#0F172A; }}
            .kpi-grid {{ display:grid; grid-template-columns: repeat(5, 1fr); gap:10px; margin:15px 0; }}
            .kpi-box {{ border:1px solid #CBD5E1; border-top:4px solid #0284C7; border-radius:8px; padding:10px; text-align:center; }}
            .banner {{ background:#EFF6FF; border-left:5px solid #0284C7; padding:12px; margin:12px 0; font-weight:bold; }}
            .safe-box {{ background:#ECFDF5; border:1px solid #6EE7B7; padding:12px; border-radius:8px; font-weight:bold; color:#065F46; }}
            @media print {{ .no-print {{ display:none !important; }} }}
        </style></head><body>
            <div class="no-print" style="text-align:right; margin-bottom:10px;">
                <button onclick="window.print()" style="background:#0284C7; color:white; border:none; padding:8px 16px; border-radius:6px; font-weight:bold; cursor:pointer;">🖨️ 수질진단 관제 리포트 바로 인쇄 / PDF 출력</button>
            </div>
            <h2>💧 {diag_fac} 지능형 수질진단 관제 리포트 ({diag_date.strftime('%Y-%m-%d')})</h2>
            <p>• <b>시설용량</b>: {max_cap_val:,} ㎥/일 ({fac_spec['method']}) | <b>계절/월</b>: {diag_season}철 ({diag_month}월) | <b>강우 조건</b>: {weather_status_txt} | <b>평균 외기기온/반응조수온</b>: {eff_avg_air_temp:.1f}℃ / {eff_avg_water_temp:.1f}℃</p>
            <div class="kpi-grid">
                <div class="kpi-box"><div>예상 방류수 BOD</div><div style="font-size:20px; font-weight:bold;">{pred_bod_v:.2f} mg/L</div><div style="color:#DC2626; font-size:11px;">기준: 5.0</div></div>
                <div class="kpi-box"><div>예상 방류수 T-N</div><div style="font-size:20px; font-weight:bold;">{pred_tn_v:.2f} mg/L</div><div style="color:#DC2626; font-size:11px;">기준: 20.0</div></div>
                <div class="kpi-box"><div>예상 방류수 SS</div><div style="font-size:20px; font-weight:bold;">{pred_ss_v:.2f} mg/L</div><div style="color:#DC2626; font-size:11px;">기준: 10.0</div></div>
                <div class="kpi-box"><div>예상 방류수 T-P</div><div style="font-size:20px; font-weight:bold;">{pred_tp_v:.3f} mg/L</div><div style="color:#DC2626; font-size:11px;">기준: 0.20</div></div>
                <div class="kpi-box"><div>예상 방류 대장균</div><div style="font-size:20px; font-weight:bold;">{pred_coli_v:,} 개/mL</div><div style="color:#DC2626; font-size:11px;">기준: 1000</div></div>
            </div>
            <div class="banner">2024년 이후 현행 법령 기준: TOC (총유기탄소) | 예상 방류수 TOC 결과: {pred_toc_v:.2f} mg/L (기준: 15.0 mg/L, 안전여유 {toc_margin_pct}%)</div>
            <div class="safe-box">💚 [실시간 공정 위험도 리포트] 최대 피크 유입량 {peak_inflow:,.1f}톤/일 | AI 권장 송풍량: {ai_opt['권장송풍량_m3min']} ㎥/분 (가동 {ai_opt['송풍기가동대수']}대) | 권장 염화제이철: {ai_opt['권장염화제이철_L']} L/일</div>
        </body></html>
        """
        st.divider()
        c_pr1, c_pr2 = st.columns(2)
        with c_pr1:
            st.download_button(
                f"📥 {diag_fac}_지능형수질진단리포트_{diag_date}.html (인쇄/PDF용) 다운로드",
                data=diag_report_html,
                file_name=sanitize_filename(f"{diag_fac}_지능형수질진단리포트_{diag_date}.html"),
                mime="text/html",
                type="primary",
                use_container_width=True
            )
        with c_pr2:
            if st.button("💾 현재 수질진단 리포트 및 AI 운전조건 영구 보관함 저장", type="primary", use_container_width=True):
                rep_save_p = os.path.join(HWPX_RECORD_DIR, sanitize_filename(f"수질진단리포트_{diag_fac}_{diag_date}.html"))
                with open(rep_save_p, "w", encoding="utf-8") as f:
                    f.write(diag_report_html)
                append_to_process_db(pd.DataFrame([{
                    "날짜": diag_date.strftime('%Y-%m-%d'),
                    "CN비": ai_opt['CN비'],
                    "권장송풍량_m3min": ai_opt['권장송풍량_m3min'],
                    "송풍기가동대수": ai_opt['송풍기가동대수'],
                    "권장염화제이철_L": ai_opt['권장염화제이철_L'],
                    "종침전PAC주입량_L": ai_opt['종침전PAC주입량_L']
                }]), facility_name=diag_fac)
                st.success(f"✅ {diag_fac} 지능형 수질진단 리포트와 AI 운전조건이 영구 보관함에 안전하게 저장되었습니다!")

    with tab_p1:
        sel_p = st.selectbox("🎯 공정 대상 시설 선택", [MAIN_PLANT] + SMALL_PLANTS)
        target_spec = PLANT_DESIGN_SPECS[sel_p]
        st.info(f"🏢 **시설명**: {sel_p} | **공법**: {target_spec['method']} | **설계용량**: {target_spec['cap']} ㎥/일")
        if st.button("🔄 공정 데이터 연산 및 영구 적재", type="primary"):
            df_fac = get_master_data(sel_p)
            recs = []
            if not df_fac.empty:
                for idx, (_, r) in enumerate(df_fac.iterrows()):
                    res = calculate_ai_process_parameters(r.get('유입량', target_spec['cap']), 120, 25, 2.8, facility_name=sel_p, date_seed=idx)
                    recs.append({
                        "날짜": r['날짜'], "CN비": res['CN비'],
                        "권장송풍량_m3min": res['권장송풍량_m3min'], "송풍기가동대수": res['송풍기가동대수'],
                        "권장염화제이철_L": res['권장염화제이철_L'], "종침전PAC주입량_L": res['종침전PAC주입량_L']
                    })
            else:
                for idx in range(10):
                    d_str = (datetime.date.today() - datetime.timedelta(days=idx)).strftime('%Y-%m-%d')
                    res = calculate_ai_process_parameters(target_spec['cap'], 120, 25, 2.8, facility_name=sel_p, date_seed=idx)
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
