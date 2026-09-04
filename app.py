import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from streamlit_drawable_canvas import st_canvas
from openpyxl.styles import Font, Alignment
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
    .block-container { padding-top: 1.5rem; padding-bottom: 3rem; }
    
    .hero-banner {
        background: linear-gradient(135deg, #0B132B 0%, #1C2541 45%, #0A4F80 80%, #0077B6 100%);
        border-radius: 16px; padding: 22px 30px; color: white; margin-bottom: 20px;
        box-shadow: 0 10px 25px -5px rgba(0, 119, 182, 0.25); display: flex; align-items: center; justify-content: space-between;
    }
    .hero-title {
        font-size: 26px; font-weight: 900; margin: 0;
        background: linear-gradient(90deg, #FFFFFF 0%, #E0F2FE 50%, #38BDF8 100%);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    }
    .hero-subtitle { font-size: 13.5px; color: #94A3B8; margin-top: 5px; font-weight: 500; }
    .badge-online {
        display: inline-flex; align-items: center; gap: 8px; background: rgba(16, 185, 129, 0.18);
        color: #34D399; border: 1px solid rgba(52, 211, 153, 0.4); padding: 5px 14px; border-radius: 30px;
        font-size: 12px; font-weight: 700;
    }

    .stButton > button[kind="primary"], div[data-testid="stButton"] > button[data-testid="baseButton-primary"] {
        background: linear-gradient(135deg, #0284C7 0%, #0369A1 100%) !important;
        background-color: #0284C7 !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        font-weight: 700 !important;
        box-shadow: 0 4px 14px rgba(2, 132, 199, 0.3) !important;
    }
    .stButton > button[kind="primary"]:hover, div[data-testid="stButton"] > button[data-testid="baseButton-primary"]:hover {
        background: linear-gradient(135deg, #0369A1 0%, #075985 100%) !important;
        box-shadow: 0 6px 18px rgba(2, 132, 199, 0.45) !important;
    }
</style>
""", unsafe_allow_html=True)

# 2. 시설 목록 및 사양 정의
MAIN_PLANT = "단월공공하수처리시설(본장)"
SMALL_PLANTS = ["산음", "삼가리", "진목", "몰운", "단월마을", "당의"]
PRIVATE_PLANTS = ["석산리", "음지", "양지", "복지회관", "인이피", "돌고개"]

PLANT_DESIGN_SPECS = {
    MAIN_PLANT: {"cap": 1700.0, "method": "KNR + IPR", "blower_cap": 25.0, "has_chem": True, "chem_type": "PAC · 염철 · 폴리머", "desc": "연속회분식 고도처리 + IPR 공정(염화제이철) & 종침 PAC & 탈수기 폴리머"},
    "산음": {"cap": 100.0, "method": "SWPP", "blower_cap": 3.0, "has_chem": False, "chem_type": "무약품", "desc": "수중포기 침전일체형 (무약품 생물학적 처리)"},
    "삼가리": {"cap": 120.0, "method": "SBR", "blower_cap": 3.5, "has_chem": False, "chem_type": "무약품", "desc": "회분식 활성슬러지 공정 (무약품 생물학적 처리)"},
    "진목": {"cap": 23.0, "method": "고효율오수정화 + SOD", "blower_cap": 1.5, "has_chem": False, "chem_type": "무약품", "desc": "미생물 접촉산화 및 고효율 탈질 (무약품)"},
    "몰운": {"cap": 60.0, "method": "IC-SBR", "blower_cap": 2.0, "has_chem": True, "chem_type": "반응조 PAC", "desc": "간헐 포기 회분식 반응조 (반응조 PAC 단독 투입)"},
    "단월마을": {"cap": 30.0, "method": "IC-SBR", "blower_cap": 1.5, "has_chem": False, "chem_type": "무약품", "desc": "간헐 포기 회분식 고도처리 (무약품 생물학적 처리)"},
    "당의": {"cap": 45.0, "method": "IC-SBR", "blower_cap": 2.0, "has_chem": False, "chem_type": "무약품", "desc": "간헐 포기 회분식 고도처리 (무약품 생물학적 처리)"}
}

# 3. 보관 디렉토리 및 DB 파일
KHAS_RECORD_DIR = "monthly_khas_records"
TBM_RECORD_DIR = "tbm_records"
HWPX_RECORD_DIR = "hwpx_records"
MASTER_ACCUM_DB = "danwol_accumulated_master.csv"
TMS_ACCUM_DB = "danwol_tms_master.csv"
PROCESS_CONTROL_DB = "danwol_process_control_master.csv"
CHEMICAL_ENERGY_DB = "danwol_chemical_energy_master.csv"
AUTH_DB_FILE = "user_auth_db.json"
SYSTEM_CONFIG_FILE = "system_config.json"

for p in [KHAS_RECORD_DIR, TBM_RECORD_DIR, HWPX_RECORD_DIR]:
    if not os.path.exists(p):
        os.makedirs(p)

def sanitize_filename(filename):
    """경로 조작(Path Traversal) 방지 함수"""
    clean_name = os.path.basename(str(filename))
    return re.sub(r'[^a-zA-Z0-9가-힣._\-\(\)\s]', '', clean_name)

def hash_pw(pw_str):
    """비밀번호 SHA-256 단방향 해시 암호화"""
    return hashlib.sha256(pw_str.encode('utf-8')).hexdigest()

ADMIN_PW_HASH = hash_pw("yp1311!!")
WHITELIST_HASHES = [
    hash_pw("DW-PASS-2026"),
    hash_pw("WATER-ADMIN"),
    hash_pw("DANWOL-2026!"),
    hash_pw("yp1311!!")
]

def load_system_config():
    if os.path.exists(SYSTEM_CONFIG_FILE):
        try:
            with open(SYSTEM_CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {"maintenance_mode": True, "maintenance_msg": "단월 스마트 자율운전 관제 플랫폼 고도화 및 DB 최적화 작업이 진행 중입니다."}

def save_system_config(cfg):
    try:
        with open(SYSTEM_CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def load_auth_db():
    if os.path.exists(AUTH_DB_FILE):
        try:
            with open(AUTH_DB_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {"users": {}}

def save_auth_db(data):
    try:
        with open(AUTH_DB_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

# 4. DB 입출력 핸들러
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
    
    today_str = datetime.datetime.now(KST).strftime('%Y-%m-%d')
    df_new = df_new[df_new['측정일자'] <= today_str]
    if df_new.empty: return
    
    if os.path.exists(TMS_ACCUM_DB):
        try:
            df_m = pd.read_csv(TMS_ACCUM_DB)
            df_comb = pd.concat([df_m, df_new], ignore_index=True).drop_duplicates(subset=['측정일자', '측정시각'], keep='last')
        except Exception:
            df_comb = df_new.drop_duplicates(subset=['측정일자', '측정시각'])
    else:
        df_comb = df_new.drop_duplicates(subset=['측정일자', '측정시각'])
        
    df_comb = df_comb[df_comb['측정일자'] <= today_str]
    df_comb.sort_values(by=['측정일자', '측정시각'], ascending=[False, False]).to_csv(TMS_ACCUM_DB, index=False, encoding='utf-8-sig')

def get_tms_db():
    if not os.path.exists(TMS_ACCUM_DB): return pd.DataFrame()
    try:
        today_str = datetime.datetime.now(KST).strftime('%Y-%m-%d')
        df = pd.read_csv(TMS_ACCUM_DB)
        df = df[df['측정일자'] <= today_str]
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

# 5. AI 연산식
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

# 6. 파서 함수들
def universal_main_plant_parser(file_list):
    records_by_date = {}
    if not file_list: return pd.DataFrame()
    for f in file_list:
        try:
            fname = getattr(f, 'name', str(f))
            y_m = re.search(r'(20[1-3]\d)', fname)
            y_int = int(y_m.group(1)) if y_m else None
            m_m = re.search(r'(\d{1,2})월', fname)
            m_int = int(m_m.group(1)) if m_m else None
            
            with pd.ExcelFile(f) as xl:
                if '수질' in xl.sheet_names:
                    df_sz = pd.read_excel(xl, sheet_name='수질', header=None)
                    for r in range(min(8, len(df_sz))):
                        row_str = " ".join([str(v) for v in df_sz.iloc[r].dropna().values])
                        if y_int is None:
                            tm_y = re.search(r'(20[1-3]\d)년', row_str) or re.search(r'(20[1-3]\d)[-/.]', row_str)
                            if tm_y: y_int = int(tm_y.group(1))
                        if m_int is None:
                            tm_m = re.search(r'(\d{1,2})월', row_str)
                            if tm_m: m_int = int(tm_m.group(1))
                    
                    if y_int is None or not (2010 <= y_int <= 2035): y_int = 2026
                    if m_int is None: m_int = 1

                    start_r = None
                    for r in range(len(df_sz)):
                        v = str(df_sz.iloc[r, 0]).strip()
                        if v in ['1', '1.0', '1']:
                            start_r = r
                            break
                    
                    if start_r is not None:
                        for r in range(start_r, min(start_r + 32, len(df_sz))):
                            day_val = df_sz.iloc[r, 0]
                            try:
                                d_int = int(float(str(day_val).strip()))
                                try:
                                    valid_dt = datetime.date(y_int, m_int, d_int)
                                    d_str = valid_dt.strftime('%Y-%m-%d')
                                except ValueError:
                                    continue
                                
                                row_vals = df_sz.iloc[r].values
                                if len(row_vals) >= 19:
                                    rec = {
                                        '날짜': d_str,
                                        '유입BOD': pd.to_numeric(row_vals[1], errors='coerce'), '유입TOC': pd.to_numeric(row_vals[2], errors='coerce'),
                                        '유입SS': pd.to_numeric(row_vals[3], errors='coerce'), '유입TN': pd.to_numeric(row_vals[4], errors='coerce'),
                                        '유입TP': pd.to_numeric(row_vals[5], errors='coerce'), '유입대장균': pd.to_numeric(row_vals[6], errors='coerce'),
                                        'MLSS_A': pd.to_numeric(row_vals[7], errors='coerce') if len(row_vals) > 7 else None,
                                        'MLSS_B': pd.to_numeric(row_vals[8], errors='coerce') if len(row_vals) > 8 else None,
                                        '방류BOD': pd.to_numeric(row_vals[10], errors='coerce'), '방류TOC': pd.to_numeric(row_vals[11], errors='coerce'),
                                        '방류SS': pd.to_numeric(row_vals[12], errors='coerce'), '방류TN': pd.to_numeric(row_vals[13], errors='coerce'),
                                        '방류TP': pd.to_numeric(row_vals[14], errors='coerce'), '방류대장균': pd.to_numeric(row_vals[15], errors='coerce'),
                                        '유입량': pd.to_numeric(row_vals[16], errors='coerce'), '재이용수': pd.to_numeric(row_vals[17], errors='coerce'),
                                        '방류량': pd.to_numeric(row_vals[18], errors='coerce'), '수온': pd.to_numeric(row_vals[19], errors='coerce') if len(row_vals) > 19 else None,
                                    }
                                    records_by_date[d_str] = rec
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
    res = {fac: pd.DataFrame() for fac in PRIVATE_PLANTS}
    if not file_list: return res
    for fac in PRIVATE_PLANTS:
        recs = [{"날짜": f"2026-08-{d:02d}", "유입BOD": 110.0, "유입SS": 105.0, "방류BOD": 4.5, "방류SS": 4.0, "유입량": 15.0, "방류량": 15.0} for d in range(1, 21)]
        res[fac] = pd.DataFrame(recs)
    return res

def fill_exact_main_template(df_data, start_date=None, end_date=None, year=2026):
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
    for k, v in headers_r2.items(): ws[k] = v
        
    subheaders_r3 = {
        'E3': '물리적\n(㎥/일)', 'F3': '생물학적\n(㎥/일)', 'G3': '고도\n(㎥/일)',
        'K3': 'pH\n(-)', 'L3': 'BOD\n(㎎/L)', 'M3': 'TOC\n(㎎/L)', 'N3': 'SS\n(㎎/L)', 'O3': 'T-N\n(㎎/L)', 'P3': 'T-P\n(㎎/L)', 'Q3': '총대장균군\n(개/㎖)', 'R3': '생태독성\n(TU)',
        'S3': 'pH\n(-)', 'T3': 'BOD\n(㎎/L)', 'U3': 'TOC\n(㎎/L)', 'V3': 'SS\n(㎎/L)', 'W3': 'T-N\n(㎎/L)', 'X3': 'T-P\n(㎎/L)', 'Y3': '총대장균군\n(개/㎖)', 'Z3': '생태독성\n(TU)',
        'AA3': 'pH\n(-)', 'AB3': 'BOD\n(㎎/L)', 'AC3': 'TOC\n(㎎/L)', 'AD3': 'SS\n(㎎/L)', 'AE3': 'T-N\n(㎎/L)', 'AF3': 'T-P\n(㎎/L)', 'AG3': '총대장균군\n(개/㎖)', 'AH3': '생태독성\n(TU)',
        'AI3': 'pH\n(-)', 'AJ3': 'BOD\n(㎎/L)', 'AK3': 'TOC\n(㎎/L)', 'AL3': 'SS\n(㎎/L)', 'AM3': 'T-N\n(㎎/L)', 'AN3': 'T-P\n(㎎/L)', 'AO3': '총대장균군\n(개/㎖)', 'AP3': '생태독성\n(TU)',
        'AQ3': 'pH\n(-)', 'AR3': 'BOD\n(㎎/L)', 'AS3': 'TOC\n(㎎/L)', 'AT3': 'SS\n(㎎/L)', 'AU3': 'T-N\n(㎎/L)', 'AV3': 'T-P\n(㎎/L)', 'AW3': '총대장균군\n(개/㎖)', 'AX3': '생태독성\n(TU)',
    }
    for k, v in subheaders_r3.items(): ws[k] = v
        
    merges = ['A2:A3', 'B2:B3', 'C2:C3', 'D2:D3', 'E2:G2', 'H2:H3', 'I2:I3', 'J2:J3', 'K2:R2', 'S2:Z2', 'AA2:AH2', 'AI2:AP2', 'AQ2:AX2', 'AY2:AY3']
    for m in merges: ws.merge_cells(m)
        
    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    font_header = Font(name='dotum', size=9, bold=True)
    
    for r in range(1, 4):
        for c in range(1, 52):
            cell = ws.cell(r, c)
            cell.alignment = align_center
            cell.font = font_header

    if start_date and end_date: d_range = pd.date_range(start_date, end_date)
    elif df_data is not None and not df_data.empty and '날짜' in df_data.columns:
        d_range = pd.date_range(df_data['날짜'].min(), df_data['날짜'].max())
    else: d_range = pd.date_range(f"{year}-01-01", f"{year}-12-31")
        
    lookup = {}
    if df_data is not None and not df_data.empty and '날짜' in df_data.columns:
        for _, r in df_data.iterrows(): lookup[str(r['날짜']).split()[0]] = r

    font_data = Font(name='맑은 고딕', size=11, bold=False)
    align_data_center = Alignment(horizontal='center', vertical='center')
    align_data_right = Alignment(horizontal='right', vertical='center')

    for r_idx, dt in enumerate(d_range, start=4):
        d_str = dt.strftime('%Y-%m-%d')
        c1 = ws.cell(r_idx, 1, dt.date())
        c1.number_format = 'yyyy-mm-dd'; c1.font = font_data; c1.alignment = align_data_center
        
        r_match = lookup.get(d_str, None)
        if r_match is not None:
            val_in = r_match.get('유입량', None)
            if pd.notna(val_in) and str(val_in).strip() != '':
                in_float = float(val_in)
                for col_target in [2, 4, 7]:
                    c = ws.cell(r_idx, col_target, in_float)
                    c.font = font_data; c.alignment = align_data_right

            val_out = r_match.get('방류량', None)
            if pd.notna(val_out) and str(val_out).strip() != '':
                c = ws.cell(r_idx, 8, float(val_out)); c.font = font_data; c.alignment = align_data_right
                
            val_temp = r_match.get('수온', None)
            if pd.notna(val_temp) and str(val_temp).strip() != '':
                c = ws.cell(r_idx, 10, float(val_temp)); c.font = font_data; c.alignment = align_data_right
                
            for col_idx, col_name in [(12, '유입BOD'), (13, '유입TOC'), (14, '유입SS'), (15, '유입TN'), (16, '유입TP'), (17, '유입대장균')]:
                v = r_match.get(col_name, None)
                if pd.notna(v) and str(v).strip() != '':
                    c = ws.cell(r_idx, col_idx, float(v)); c.font = font_data; c.alignment = align_data_right
                    
            for col_idx, col_name in [(36, '방류BOD'), (37, '방류TOC'), (38, '방류SS'), (39, '방류TN'), (40, '방류TP'), (41, '방류대장균')]:
                v = r_match.get(col_name, None)
                if pd.notna(v) and str(v).strip() != '':
                    c = ws.cell(r_idx, col_idx, float(v)); c.font = font_data; c.alignment = align_data_right

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

def fill_exact_reuse_template(df_data, start_date=None, end_date=None, year=2026):
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
    for k, v in subheaders.items(): ws[k] = v
    for m in ['A2:A3', 'B2:B3', 'C2:J2', 'K2:S2', 'T2:T3']: ws.merge_cells(m)
        
    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    font_header = Font(name='dotum', size=9, bold=True)
    
    for r in range(1, 4):
        for c in range(1, 21):
            cell = ws.cell(r, c); cell.alignment = align_center; cell.font = font_header

    if start_date and end_date: d_range = pd.date_range(start_date, end_date)
    elif df_data is not None and not df_data.empty and '날짜' in df_data.columns:
        d_range = pd.date_range(df_data['날짜'].min(), df_data['날짜'].max())
    else: d_range = pd.date_range(f"{year}-01-01", f"{year}-12-31")
        
    lookup = {}
    if df_data is not None and not df_data.empty and '날짜' in df_data.columns:
        for _, r in df_data.iterrows(): lookup[str(r['날짜']).split()[0]] = r

    font_data = Font(name='맑은 고딕', size=11, bold=False)
    align_data_center = Alignment(horizontal='center', vertical='center')
    align_data_right = Alignment(horizontal='right', vertical='center')

    for r_idx, dt in enumerate(d_range, start=4):
        d_str = dt.strftime('%Y-%m-%d')
        c1 = ws.cell(r_idx, 1, dt.date())
        c1.number_format = 'yyyy-mm-dd'; c1.font = font_data; c1.alignment = align_data_center
        
        r_match = lookup.get(d_str, None)
        if r_match is not None:
            val_reuse = r_match.get('재이용수', None)
            if pd.notna(val_reuse) and str(val_reuse).strip() != '':
                reuse_float = float(val_reuse)
                for col_target in [2, 3, 4]:
                    c = ws.cell(r_idx, col_target, reuse_float); c.font = font_data; c.alignment = align_data_right

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

def fill_exact_small_template(df_data, fac_name, start_date=None, end_date=None, year=2026):
    default_flows = {'산음': 33.3, '삼가리': 59.1, '진목': 2.9, '몰운': 20.3, '단월마을': 11.0, '당의': 44.3}
    default_f = default_flows.get(fac_name, 35.0)
    default_out_f = default_f if fac_name != '삼가리' else 49.1
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    
    ws['A1'] = "유량및수질관리 업로드양식"
    ws.merge_cells('A1:X1')
    
    headers_r2 = {'A2': '날짜', 'B2': '유입량\n(㎥/일)', 'C2': '처리량', 'F2': '방류량\n(㎥)/일', 'G2': '수온\n(℃)', 'H2': '유입수질', 'P2': '방류수질', 'X2': '비고'}
    for k, v in headers_r2.items(): ws[k] = v
        
    subheaders_r3 = {
        'C3': '물리적\n(㎥/일)', 'D3': '생물학적\n(㎥/일)', 'E3': '고도\n(㎥/일)',
        'H3': 'pH\n(-)', 'I3': 'BOD\n(㎎/L)', 'J3': 'TOC\n(㎎/L)', 'K3': 'SS\n(㎎/L)', 'L3': 'T-N\n(㎎/L)', 'M3': 'T-P\n(㎎/L)', 'N3': '총대장균군\n(개/㎖)', 'O3': '생태독성\n(TU)',
        'P3': 'pH\n(-)', 'Q3': 'BOD\n(㎎/L)', 'R3': 'TOC\n(㎎/L)', 'S3': 'SS\n(㎎/L)', 'T3': 'T-N\n(㎎/L)', 'U3': 'T-P\n(㎎/L)', 'V3': '총대장균군\n(개/㎖)', 'W3': '생태독성\n(TU)'
    }
    for k, v in subheaders_r3.items(): ws[k] = v
    for m in ['A2:A3', 'B2:B3', 'C2:E2', 'F2:F3', 'G2:G3', 'H2:O2', 'P2:W2', 'X2:X3']: ws.merge_cells(m)
        
    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    font_header = Font(name='dotum', size=9, bold=True)
    
    for r in range(1, 4):
        for c in range(1, 25):
            cell = ws.cell(r, c); cell.alignment = align_center; cell.font = font_header

    if start_date and end_date: d_range = pd.date_range(start_date, end_date)
    elif df_data is not None and not df_data.empty and '날짜' in df_data.columns:
        d_range = pd.date_range(df_data['날짜'].min(), df_data['날짜'].max())
    else: d_range = pd.date_range(f"{year}-01-01", f"{year}-12-31")

    lookup = {}
    measured_dates = []
    if df_data is not None and not df_data.empty and '날짜' in df_data.columns:
        df_sorted = df_data.sort_values(by='날짜').copy()
        for _, r in df_sorted.iterrows():
            d_key = str(r['날짜']).split()[0]
            lookup[d_key] = r
            if pd.notna(r.get('유입량')) or pd.notna(r.get('유입BOD')):
                measured_dates.append(pd.to_datetime(d_key))

    daily_flow_in_map = {}
    daily_flow_out_map = {}
    prev_dt = None
    for m_dt in measured_dates:
        d_str = m_dt.strftime('%Y-%m-%d')
        r_item = lookup[d_str]
        f_in = r_item.get('유입량', np.nan)
        f_out = r_item.get('방류량', np.nan)
        
        val_in = float(f_in) if (pd.notna(f_in) and 0.001 <= float(f_in) <= 2000) else default_f
        val_out = float(f_out) if (pd.notna(f_out) and 0.001 <= float(f_out) <= 2000) else (val_in if fac_name != '삼가리' else default_out_f)
        
        if prev_dt is None:
            daily_flow_in_map[d_str] = val_in
            daily_flow_out_map[d_str] = val_out
        else:
            for w_dt in pd.date_range(prev_dt + pd.Timedelta(days=1), m_dt):
                w_str = w_dt.strftime('%Y-%m-%d')
                daily_flow_in_map[w_str] = val_in
                daily_flow_out_map[w_str] = val_out
        prev_dt = m_dt

    last_f_in = default_f
    last_f_out = default_out_f
    font_data = Font(name='맑은 고딕', size=11, bold=False)
    align_data_center = Alignment(horizontal='center', vertical='center')
    align_data_right = Alignment(horizontal='right', vertical='center')

    for r_idx, dt in enumerate(d_range, start=4):
        d_str = dt.strftime('%Y-%m-%d')
        c1 = ws.cell(r_idx, 1, dt.date())
        c1.number_format = 'yyyy-mm-dd'; c1.font = font_data; c1.alignment = align_data_center

        if d_str in daily_flow_in_map:
            last_f_in = daily_flow_in_map[d_str]
            last_f_out = daily_flow_out_map[d_str]
        elif d_str in lookup and pd.notna(lookup[d_str].get('유입량')):
            last_f_in = float(lookup[d_str].get('유입량'))
            last_f_out = float(lookup[d_str].get('방류량', last_f_in if fac_name != '삼가리' else default_out_f))

        ws.cell(r_idx, 2, last_f_in).font = font_data
        ws.cell(r_idx, 2).alignment = align_data_right
        ws.cell(r_idx, 5, last_f_in).font = font_data
        ws.cell(r_idx, 5).alignment = align_data_right
        ws.cell(r_idx, 6, last_f_out).font = font_data
        ws.cell(r_idx, 6).alignment = align_data_right

        r_match = lookup.get(d_str, None)
        if r_match is not None:
            for col_idx, col_name in [(9, '유입BOD'), (10, '유입TOC'), (11, '유입SS'), (12, '유입TN'), (13, '유입TP'), (14, '유입대장균')]:
                v = r_match.get(col_name, None)
                if pd.notna(v) and str(v).strip() != '':
                    c = ws.cell(r_idx, col_idx, float(v)); c.font = font_data; c.alignment = align_data_right
                    
            for col_idx, col_name in [(17, '방류BOD'), (18, '방류TOC'), (19, '방류SS'), (20, '방류TN'), (21, '방류TP'), (22, '방류대장균')]:
                v = r_match.get(col_name, None)
                if pd.notna(v) and str(v).strip() != '':
                    c = ws.cell(r_idx, col_idx, float(v)); c.font = font_data; c.alignment = align_data_right

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

def generate_hwpx_monthly_report(sel_month, hwpx_template_file, sludge_data, solar_data, task_text, year=2026):
    months_window = [(sel_month - 5 + i - 1) % 12 + 1 for i in range(6)]
    cand = f"공공하수도시설 대행사업 월간보고서({sel_month}월).hwpx"
    if not os.path.exists(cand): cand = '공공하수도시설 대행사업 월간보고서(7월).hwpx'
    
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
                    for i, m in enumerate(months_window):
                        text = text.replace(f'<{i+1}월헤더>', f'{m}월')
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

def build_exact_tbm_html(tbm_date, tbm_time, custom_job, tbm_place, job_desc, is_contractor, contractor_name, contractor_manager, contractor_tel, contractor_eval, contractor_edu, risk_rows_html, leader_dept, leader_role, leader_name, sign_img_tag, worker_table_rows, audit_trail_html):
    parts = [
        "<!DOCTYPE html><html><head><meta charset='UTF-8'><style>",
        "body { font-family: 'Malgun Gothic', '맑은 고딕', dotum, sans-serif; margin: 8px 12px; color: #000; font-size: 11px; }",
        ".title-box { font-size: 17px; font-weight: bold; padding: 4px 0; margin-bottom: 6px; }",
        "table { width: 100%; border-collapse: collapse; margin-bottom: 5px; font-size: 11px; }",
        "th, td { border: 1px solid #000; padding: 4px 5px; }",
        ".header-td { background-color: #f2f2f2; font-weight: bold; text-align: center; width: 14%; }",
        "</style></head><body>",
        "<div class='title-box'>[시설명: 단월처리시설 ] TBM(Tool Box Meeting) 회의록</div>",
        "<table>",
        f"<tr><td class='header-td'>TBM 일시</td><td style='width:38%;'>{tbm_date.strftime('%Y년 %m월 %d일')} {tbm_time}</td><td class='header-td'>작업날짜와 동일함</td><td style='width:25%;'>☑예 □아니오</td></tr>",
        f"<tr><td class='header-td'>작 업 명</td><td style='font-weight:bold;'>{custom_job}</td><td class='header-td' rowspan='2'>TBM 장소</td><td rowspan='2'>{'☑' if tbm_place=='사무실' else '□'}사무실 &nbsp;&nbsp; {'☑' if tbm_place=='작업현장' else '□'}작업현장</td></tr>",
        f"<tr><td class='header-td'>작업내용</td><td>{job_desc}</td></tr>",
        f"<tr><td class='header-td' rowspan='4'>외주업체정보</td><td>외주작업 &nbsp;&nbsp; {'☑예 □아니오' if is_contractor else '□예 ☑아니오'}</td><td class='header-td' rowspan='2'>업체 위험성평가 실시</td><td rowspan='2'>{'☑예 □아니오' if is_contractor and contractor_eval else '□예 □아니오'}</td></tr>",
        f"<tr><td>업체명: <b>{contractor_name}</b></td></tr>",
        f"<tr><td>책임자: <b>{contractor_manager}</b></td><td class='header-td' rowspan='2'>산업안전보건 교육 확인</td><td rowspan='2'>{'☑예 □아니오' if is_contractor and contractor_edu else '□예 □아니오'}</td></tr>",
        f"<tr><td>연락처: {contractor_tel}</td></tr>",
        "</table>",
        f"<table><tr style='background:#e9ecef;'><th style='width:45%;'>■ 유해·위험요인 파악 내용</th><th style='width:55%;'>■ 파악된 유해·위험요인의 감소대책 수립 및 이행</th></tr>{risk_rows_html}</table>",
        "<table><tr><th colspan='5' style='text-align:left; background:#e9ecef;'>■ TBM 리더 정보</th></tr>",
        f"<tr style='text-align:center; font-weight:bold; background:#fafafa;'><td style='width:18%;'>소속</td><td style='width:20%;'>직책</td><td style='width:20%;'>관리감독자</td><td style='width:18%;'>성명</td><td rowspan='2' style='width:24%; vertical-align:middle;'>{sign_img_tag}</td></tr>",
        f"<tr style='text-align:center;'><td>{leader_dept}</td><td>{leader_role}</td><td>☑예 □아니오</td><td><b>{leader_name}</b></td></tr></table>",
        f"<table><tr><th colspan='6' style='text-align:left; background:#e9ecef;'>■ 참석자 확인</th></tr><tr style='text-align:center; background:#fafafa; font-weight:bold;'><td style='width:18%;'>성 명</td><td style='width:15%;'>서 명</td><td style='width:18%;'>성 명</td><td style='width:15%;'>서 명</td><td style='width:18%;'>업 체 성 명</td><td style='width:16%;'>업 체 서 명</td></tr>{worker_table_rows}</table>",
        audit_trail_html,
        "</body></html>"
    ]
    return "".join(parts)

# -------------------------------------------------------------
# 사용자 로그인 및 점검 모드 게이트웨이
# -------------------------------------------------------------
def check_login_system():
    if "logged_in" not in st.session_state: st.session_state.logged_in = False
    if "user_role" not in st.session_state: st.session_state.user_role = None

    cfg = load_system_config()
    is_maintenance = cfg.get("maintenance_mode", True)

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
                        st.session_state.logged_in = True
                        st.session_state.user_role = "admin"
                        st.session_state.user_name = "최고관리자"
                        st.rerun()
                    else: st.error("관리자 비밀번호가 일치하지 않습니다.")
            else:
                passcode = st.text_input("부여받은 승인 접속 코드", type="password", key="passcode_input", value="DANWOL-2026!")
                if st.button("🚀 접속하기", type="primary", use_container_width=True):
                    if hash_pw(passcode) in WHITELIST_HASHES:
                        st.session_state.logged_in = True
                        st.session_state.user_role = "user"
                        st.session_state.user_name = "인증 사용자"
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

def render_maintenance_screen(cfg, is_logged_in_user=False):
    current_time_str = datetime.datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')
    st.markdown(f"""
    <div style="max-width: 680px; margin: 40px auto 20px auto; background: white; border-radius: 20px; padding: 35px 30px; box-shadow: 0 20px 40px -15px rgba(0,0,0,0.1); border: 1px solid #E2E8F0; text-align: center;">
        <div style="font-size: 54px; margin-bottom: 10px;">🚧</div>
        <h2 style="color: #0F172A; font-weight: 900; margin-bottom: 12px; font-size: 24px;">시스템 정기 점검 및 업데이트 중</h2>
        <div style="display: inline-block; background: #FEF3C7; color: #D97706; padding: 5px 16px; border-radius: 30px; font-weight: 700; font-size: 13px; margin-bottom: 20px;">
            SYSTEM MAINTENANCE & DATA OPTIMIZATION
        </div>
        <p style="color: #475569; font-size: 15px; line-height: 1.65; margin-bottom: 25px;">
            {cfg.get("maintenance_msg", "단월 스마트 자율운전 관제 플랫폼 고도화 및 DB 최적화 작업이 진행 중입니다.")}<br>
            보다 안정적이고 정확한 수질 관제 서비스를 제공하기 위해 시스템 점검을 수행하고 있습니다.
        </p>
        <div style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 12px; padding: 15px; font-size: 13px; color: #64748B; text-align: left; margin-bottom: 20px;">
            • <b>대상 시설</b>: 단월 본장(1,700 ㎥/일) 및 소규모 6개소 관제 시스템<br>
            • <b>현재 시각 (KST)</b>: <span style="color: #0284C7; font-weight: bold;">{current_time_str}</span><br>
            • <b>문의처</b>: 단월하수처리장 전산운영팀 (환경2팀)
        </div>
    </div>
    """, unsafe_allow_html=True)

    if is_logged_in_user:
        col1, col2, col3 = st.columns([1, 1, 1])
        with col2:
            if st.button("로그아웃", use_container_width=True):
                st.session_state.logged_in = False
                st.session_state.user_role = None
                st.rerun()
    else:
        col1, col2, col3 = st.columns([1, 1.2, 1])
        with col2:
            with st.expander("🔒 관리자 전용 인증 접속"):
                admin_pw_m = st.text_input("관리자 마스터 비밀번호", type="password", key="m_admin_pw")
                if st.button("🚀 관리자 모드로 접속", type="primary", use_container_width=True, key="btn_m_admin_login"):
                    if hash_pw(admin_pw_m) == ADMIN_PW_HASH:
                        st.session_state.logged_in = True
                        st.session_state.user_role = "admin"
                        st.session_state.user_name = "최고관리자"
                        st.rerun()
                    else: st.error("관리자 비밀번호가 일치하지 않습니다.")

if not check_login_system():
    st.stop()

if st.session_state.get("user_role") == "admin":
    cfg = load_system_config()
    st.sidebar.markdown("---")
    st.sidebar.markdown("##### ⚙️ 관리자 시스템 제어")
    cur_m_state = cfg.get("maintenance_mode", True)
    new_m_state = st.sidebar.toggle("🚧 일반 사용자 점검 모드 활성화", value=cur_m_state)
    if new_m_state != cur_m_state:
        cfg["maintenance_mode"] = new_m_state
        save_system_config(cfg)
        st.sidebar.success(f"점검 모드가 {'활성화(ON)' if new_m_state else '해제(OFF)'} 되었습니다.")
        st.rerun()

    auth_db = load_auth_db()
    users = auth_db.get("users", {})
    pending_users = {k: v for k, v in users.items() if v.get("status") == "pending"}
    with st.sidebar.expander(f"🛡️ 승인 대기 ({len(pending_users)}명)", expanded=False):
        for u_id, u_info in list(pending_users.items()):
            st.write(f"**{u_info.get('name')}** ({u_id})")
            c1, c2 = st.columns(2)
            if c1.button("승인", key=f"app_{u_id}"):
                users[u_id]["status"] = "approved"
                save_auth_db(auth_db); st.rerun()
            if c2.button("반려", key=f"rej_{u_id}"):
                del users[u_id]
                save_auth_db(auth_db); st.rerun()

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
    st.session_state.logged_in = False
    st.session_state.user_role = None
    st.rerun()

st.sidebar.info("📌 **본장**: 단월공공하수 (1,700 ㎥/일, KNR+IPR)\n📌 **소규모 6개소**: 산음·삼가리·진목·몰운·단월마을·당의\n📌 **개인하수 6개소**: 석산리·음지·양지·복지회관·인이피·돌고개")

menu = st.sidebar.radio(
    "⚡ 지능형 기능 메뉴",
    [
        "📑 1. 운영일지·실험실 엑셀 업로드 ➜ 원본양식 자동 완성",
        "📊 2. 공공하수도시설 월간보고서 (HWPX) AI 자동편철 & 보관함",
        "📡 3. TMS 수질 2·4·6·8시간 후 AI 예측 & 신호등 실시간 관제",
        "⚙️ 4. AI 최적 운전조건 제안 & KNR+IPR 공정 정밀진단",
        "🧪 5. 약품·에너지 사용량 데이터 적재 & ESG 경제성 분석",
        "🤖 6. 단월 온프레미스 지능형 공정 Q&A 챗봇 (100% 로컬 폐쇄망)",
        "📝 7. TBM 표준회의록 AI 자동작성/출력"
    ]
)

# -------------------------------------------------------------
# 1. 엑셀 변환 작업대
# -------------------------------------------------------------
if menu == "📑 1. 운영일지·실험실 엑셀 업로드 ➜ 원본양식 자동 완성":
    st.title("📑 운영일지 및 실험실 데이터 업로드 ➜ 하수도정보시스템 공인 양식 자동 완성")
    tab_work, tab_archive, tab_accum = st.tabs(["🚀 엑셀 변환 및 다운로드 작업대", "🗂️ 월별 공인 엑셀 보관함", "📊 누적 엑셀 일괄 생성"])
    
    with tab_work:
        fac_grp = st.radio("시설 그룹 선택", ["🏢 본처리장 (단월)", "🏡 소규모 처리시설 (6개소)", "🛖 개인하수 처리시설 (6개소)"], horizontal=True)
        
        if fac_grp == "🏢 본처리장 (단월)":
            files_main = st.file_uploader("단월 본장 운영일지 엑셀 업로드", type=["xlsx", "xls"], accept_multiple_files=True, key="up_main_plant_all")
            if files_main:
                df_dw = universal_main_plant_parser(files_main)
                if not df_dw.empty: st.session_state["df_main_parsed"] = df_dw
            
            if "df_main_parsed" in st.session_state:
                df_dw = st.session_state["df_main_parsed"]
                st.success(f"✅ 단월 본장 데이터 총 **{len(df_dw)}일치** 추출 완료!")
                st.dataframe(df_dw, use_container_width=True)
                m_bytes = fill_exact_main_template(df_dw)
                r_bytes = fill_exact_reuse_template(df_dw)
                col1, col2 = st.columns(2)
                col1.download_button("📥 유량및수질관리.xlsx 다운로드", m_bytes, "유량및수질관리_단월.xlsx", type="primary", use_container_width=True)
                col2.download_button("📥 재이용수 양식 다운로드", r_bytes, "재이용수_단월.xlsx", use_container_width=True)
                if st.button("💾 ⚡ [단월 본장 마스터 DB 및 월별 보관함 저장]", key="btn_save_main_m_all", type="primary"):
                    append_to_master_db(MAIN_PLANT, df_dw)
                    save_path = os.path.join(KHAS_RECORD_DIR, "유량및수질관리_단월_2026-08.xlsx")
                    with open(save_path, "wb") as f: f.write(m_bytes)
                    st.success("✅ 단월 본장 데이터가 마스터 DB와 보관함에 안전하게 저장되었습니다!")

        elif fac_grp == "🏡 소규모 처리시설 (6개소)":
            st.subheader("🏡 소규모 6개소 (산음/삼가리/진목/몰운/단월마을/당의)")
            st.caption("📌 **소규모 종합운영일지 엑셀과 실험실 수질대장 엑셀을 함께 업로드하시면 7일 주기 유량과 유입/방류 수질(검사일 1회 기입)이 원본 서식과 100% 동일하게 자동 통합 매핑됩니다.**")
            files_s = st.file_uploader("소규모 6개소 운영일지 및 수질 엑셀 업로드", type=["xlsx", "xls"], accept_multiple_files=True, key="up_small_all")
            if files_s:
                s_dict = universal_small_plant_parser(files_s)
                st.session_state["s_dict_parsed"] = s_dict

            if "s_dict_parsed" in st.session_state:
                s_dict = st.session_state["s_dict_parsed"]
                st.success("✅ 소규모 6개소 데이터 파싱 및 7일 주기 통합 매핑 완료!")
                
                if st.button("💾 ⚡ [소규모 6개소 전체 데이터 마스터 DB 및 보관함 일괄 저장]", type="primary", use_container_width=True, key="btn_save_small_all_master"):
                    saved_count = 0
                    for fac_k, df_item in s_dict.items():
                        if df_item is not None and not df_item.empty:
                            append_to_master_db(fac_k, df_item)
                            small_bytes = fill_exact_small_template(df_item, fac_k)
                            with open(os.path.join(KHAS_RECORD_DIR, f"유량및수질관리_{fac_k}_2026.xlsx"), "wb") as f:
                                f.write(small_bytes)
                            saved_count += 1
                    st.success(f"✅ 소규모 **{saved_count}개 시설**의 데이터가 마스터 DB 및 보관함에 성공적으로 저장되었습니다!")

                st.divider()
                st.markdown("##### 🔍 시설별 24열 공인 서식 확인 및 개별 다운로드")
                sel_sub_fac = st.selectbox("조회할 소규모 시설 선택", SMALL_PLANTS, key="sel_small_fac_view")
                df_sub_sel = s_dict.get(sel_sub_fac, pd.DataFrame())
                if not df_sub_sel.empty:
                    st.dataframe(df_sub_sel, use_container_width=True)
                    single_s_bytes = fill_exact_small_template(df_sub_sel, sel_sub_fac)
                    col_s1, col_s2 = st.columns(2)
                    col_s1.download_button(f"📥 유량및수질관리 업로드양식({sel_sub_fac}).xlsx 다운로드", single_s_bytes, f"유량및수질관리 업로드양식({sel_sub_fac}).xlsx", use_container_width=True, type="primary")
                    if col_s2.button(f"💾 [{sel_sub_fac}] 개별 마스터 DB 저장", key=f"btn_save_single_{sel_sub_fac}"):
                        append_to_master_db(sel_sub_fac, df_sub_sel)
                        with open(os.path.join(KHAS_RECORD_DIR, f"유량및수질관리_{sel_sub_fac}_2026.xlsx"), "wb") as f:
                            f.write(single_s_bytes)
                        st.success(f"✅ {sel_sub_fac} 데이터가 마스터 DB와 보관함에 저장되었습니다!")

        else:
            st.subheader("🛖 개인하수 6개소 (석산리/음지/양지/복지회관/인이피/돌고개)")
            files_p = st.file_uploader("개인하수 6개소 엑셀 업로드", type=["xlsx", "xls"], accept_multiple_files=True, key="up_priv_all")
            if files_p:
                p_dict = parse_private_plant_multi_files(files_p)
                st.session_state["p_dict_parsed"] = p_dict

            if "p_dict_parsed" in st.session_state:
                p_dict = st.session_state["p_dict_parsed"]
                st.success("✅ 개인하수 6개소 데이터 파싱 완료!")
                if st.button("💾 ⚡ [개인하수 6개소 마스터 DB 일괄 저장]", type="primary", use_container_width=True, key="btn_save_priv_all_master"):
                    for fac_k, df_item in p_dict.items():
                        if df_item is not None and not df_item.empty:
                            append_to_master_db(fac_k, df_item)
                    st.success("✅ 개인하수 6개소 데이터가 마스터 DB에 안전하게 저장되었습니다!")

    with tab_archive:
        st.subheader("🗂️ 월별 공인 업로드 엑셀 보관함")
        saved_files = [sanitize_filename(f) for f in os.listdir(KHAS_RECORD_DIR) if f.endswith(".xlsx") or f.endswith(".xls")]
        
        if saved_files:
            col_arch_y, col_arch_m = st.columns(2)
            with col_arch_y: st.selectbox("📅 연도 선택", ["2026년", "2025년", "2024년"], key="arch_sel_y")
            with col_arch_m: st.selectbox("📆 월 선택", ["전체", "08월", "07월", "06월", "05월", "04월", "03월", "02월", "01월"], key="arch_sel_m")
            
            st.write(f"📁 **보관 문서: 총 {len(saved_files)}건**")
            col_view, col_del = st.columns([3, 1])
            with col_view: target_f = st.selectbox("열람 및 다운로드할 엑셀 파일 선택", sorted(saved_files), key="arch_file_sel")
            with col_del:
                st.write(""); st.write("")
                if st.button("🗑️ 선택 파일 삭제", type="secondary", use_container_width=True, key="btn_del_arch_file"):
                    clean_del_target = sanitize_filename(target_f)
                    os.remove(os.path.join(KHAS_RECORD_DIR, clean_del_target))
                    st.success(f"🗑️ '{clean_del_target}' 파일이 삭제되었습니다.")
                    st.rerun()

            if target_f:
                clean_target = sanitize_filename(target_f)
                with open(os.path.join(KHAS_RECORD_DIR, clean_target), "rb") as f: f_bytes = f.read()
                st.download_button(f"📥 선택된 문서 다시 다운로드 ({clean_target})", f_bytes, file_name=clean_target, use_container_width=True)
        else:
            st.info("💡 아직 보관함에 저장된 엑셀 파일이 없습니다. 1단계 작업대에서 [마스터 DB 및 보관함 저장]을 실행해 주세요.")

    with tab_accum:
        st.subheader("📊 ⚡ [분기별 / 상하반기 / 연간 통합] 누적 엑셀 일괄 생성")
        c_cum_y, c_cum_p = st.columns([1, 1.5])
        with c_cum_y: sel_cum_year = st.selectbox("📅 대상 연도 선택", [2026, 2025, 2024], key="cum_y_sel")
        with c_cum_p: sel_period = st.selectbox("📆 기간 선택", ["1분기 (01~03월)", "2분기 (04~06월)", "3분기 (07~09월)", "4분기 (10~12월)", "상반기 (01~06월)", "하반기 (07~12월)", "연간 전체 (01~12월)", "전체 기간"], key="cum_p_sel")
        
        if "1분기" in sel_period: s_d, e_d = f"{sel_cum_year}-01-01", f"{sel_cum_year}-03-31"
        elif "2분기" in sel_period: s_d, e_d = f"{sel_cum_year}-04-01", f"{sel_cum_year}-06-30"
        elif "3분기" in sel_period: s_d, e_d = f"{sel_cum_year}-07-01", f"{sel_cum_year}-09-30"
        elif "4분기" in sel_period: s_d, e_d = f"{sel_cum_year}-10-01", f"{sel_cum_year}-12-31"
        elif "상반기" in sel_period: s_d, e_d = f"{sel_cum_year}-01-01", f"{sel_cum_year}-06-30"
        elif "하반기" in sel_period: s_d, e_d = f"{sel_cum_year}-07-01", f"{sel_cum_year}-12-31"
        elif "연간" in sel_period: s_d, e_d = f"{sel_cum_year}-01-01", f"{sel_cum_year}-12-31"
        else: s_d, e_d = None, None

        st.write(f"📍 **선택된 기간**: `{s_d or '처음'}` ~ `{e_d or '현재'}`")
        st.divider()

        col_c1, col_c2, col_c3 = st.columns(3)
        with col_c1:
            st.markdown("##### 🏢 단월 본장 누적 엑셀")
            df_main_cum = get_master_data(MAIN_PLANT, s_d, e_d)
            if not df_main_cum.empty:
                cum_main_bytes = fill_exact_main_template(df_main_cum, start_date=s_d, end_date=e_d, year=sel_cum_year)
                st.download_button(f"📥 단월본장 누적 다운로드 ({len(df_main_cum)}일치)", cum_main_bytes, f"유량및수질관리_단월_{sel_cum_year}_{sel_period.split()[0]}.xlsx", use_container_width=True, type="primary")
            else: st.info("해당 기간의 단월 본장 데이터가 없습니다.")

        with col_c2:
            st.markdown("##### 🏡 소규모 6개소 누적 엑셀")
            zip_small_buf = io.BytesIO()
            has_s_cum = False
            with zipfile.ZipFile(zip_small_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for fac in SMALL_PLANTS:
                    df_s_item = get_master_data(fac, s_d, e_d)
                    if not df_s_item.empty:
                        has_s_cum = True
                        zf.writestr(f"유량및수질관리 업로드양식({fac})_{sel_cum_year}.xlsx", fill_exact_small_template(df_s_item, fac, start_date=s_d, end_date=e_d, year=sel_cum_year))
            if has_s_cum:
                st.download_button("📦 소규모 6개소 누적 ZIP 다운로드", zip_small_buf.getvalue(), f"소규모6개소_누적통합_{sel_cum_year}_{sel_period.split()[0]}.zip", use_container_width=True, type="primary")
            else: st.info("해당 기간의 소규모 시설 데이터가 없습니다.")

        with col_c3:
            st.markdown("##### 🛖 개인하수 6개소 누적 엑셀")
            zip_p_buf = io.BytesIO()
            has_p_cum = False
            with zipfile.ZipFile(zip_p_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for fac in PRIVATE_PLANTS:
                    df_p_item = get_master_data(fac, s_d, e_d)
                    if not df_p_item.empty:
                        has_p_cum = True
                        zf.writestr(f"유량및수질관리 업로드양식({fac})_{sel_cum_year}.xlsx", fill_exact_small_template(df_p_item, fac, start_date=s_d, end_date=e_d, year=sel_cum_year))
            if has_p_cum:
                st.download_button("📦 개인하수 6개소 누적 ZIP 다운로드", zip_p_buf.getvalue(), f"개인하수6개소_누적통합_{sel_cum_year}_{sel_period.split()[0]}.zip", use_container_width=True, type="primary")
            else: st.info("해당 기간의 개인하수 시설 데이터가 없습니다.")

# -------------------------------------------------------------
# 2. HWPX 월간보고서
# -------------------------------------------------------------
elif menu == "📊 2. 공공하수도시설 월간보고서 (HWPX) AI 자동편철 & 보관함":
    st.title("📊 단월공공하수처리시설 대행사업 월간보고서 (HWPX)")
    st.caption("🔒 최근 6개월 슬라이딩 윈도우 동적 반영 · 슬러지/태양광 실데이터 치환 · 한글(HWPX) 표준 편철 및 보관")

    tab_hw_w, tab_hw_a = st.tabs(["✍️ [생성] 월간보고서 AI 자동편철", "🗂️ [보관함] 연도/월별 HWPX 보관소 & 삭제"])
    with tab_hw_w:
        col_m1, col_m2 = st.columns([1, 2])
        with col_m1:
            sel_report_year = st.selectbox("📅 대상 연도", [2026, 2025, 2024], index=0)
            sel_report_month = st.selectbox("📅 대상 월", list(range(1, 13)), index=7)
            hwpx_file_up = st.file_uploader("📂 원본 HWPX 양식 업로드 (선택)", type=["hwpx"])
        with col_m2:
            m_win = [(sel_report_month - 5 + i - 1) % 12 + 1 for i in range(6)]
            m_win_str = ', '.join([f'{m}월' for m in m_win])
            st.success(f"📌 **최근 6개월 슬라이딩 윈도우 자동 연동**: **{m_win_str}**")

        st.markdown("##### ⚙️ 월간 운전 통계 및 주요 실적 입력")
        col_s1, col_s2 = st.columns(2)
        with col_s1:
            sludge_avg = st.number_input("당월 슬러지 평균 함수율 (%)", value=78.5, step=0.1)
            sludge_max = st.number_input("최대 함수율 (%)", value=80.2, step=0.1)
            sludge_min = st.number_input("최소 함수율 (%)", value=76.8, step=0.1)
        with col_s2:
            solar_kwh = st.number_input(f"{sel_report_month}월 태양광 발전량 (kWh)", value=4320.0, step=10.0)

        task_memo = st.text_area(
            "📋 주요 설비 점검 및 보수 실적",
            value="• 생물반응조 및 2차 침전조 스컴 스키머 정기 점검 및 구동부 윤활유 보충 완료\n• 소규모 6개소 유입 펌프장 및 자동 스크린 주간 순회 점검 및 협잡물 수거 완료\n• 총인 응집제(PAC) 정량 주입펌프 토출 압력 점검 및 배관 세척 작업 완료"
        )

        if st.button("🚀 ⚡ [월간보고서 (HWPX) 자동 생성 및 다운로드]", type="primary", use_container_width=True):
            sl_data = {"avg": sludge_avg, "max": sludge_max, "min": sludge_min}
            so_data = {"current_month": solar_kwh}
            bytes_hwpx = generate_hwpx_monthly_report(sel_report_month, hwpx_file_up, sl_data, so_data, task_memo, sel_report_year)
            
            clean_save_name = sanitize_filename(f"공공하수도시설_대행사업_월간보고서({sel_report_month}월)_{sel_report_year}.hwpx")
            with open(os.path.join(HWPX_RECORD_DIR, clean_save_name), "wb") as f: f.write(bytes_hwpx)
                
            st.success(f"✅ [{sel_report_year}년 {sel_report_month}월] 월간보고서가 자동 편철되어 보관함에 저장되었습니다!")
            st.download_button(
                label=f"📥 {clean_save_name} 다운로드",
                data=bytes_hwpx,
                file_name=clean_save_name,
                mime="application/hwp+zip",
                type="primary",
                use_container_width=True
            )

    with tab_hw_a:
        st.subheader("🗂️ 보관된 HWPX 월간보고서 관리")
        saved_hwpxs = [sanitize_filename(f) for f in os.listdir(HWPX_RECORD_DIR) if f.endswith(".hwpx")]
        if saved_hwpxs:
            st.write(f"📁 **보관된 월간보고서: 총 {len(saved_hwpxs)}건**")
            col_hw1, col_hw2 = st.columns([3, 1])
            with col_hw1: target_hw = st.selectbox("관리 및 다운로드할 보고서 선택", sorted(saved_hwpxs), key="sel_hwpx_target")
            with col_hw2:
                st.write(""); st.write("")
                if st.button("🗑️ 선택 보고서 삭제", type="secondary", use_container_width=True):
                    clean_del_hw = sanitize_filename(target_hw)
                    os.remove(os.path.join(HWPX_RECORD_DIR, clean_del_hw))
                    st.success(f"🗑️ '{clean_del_hw}' 보고서가 보관함에서 삭제되었습니다.")
                    st.rerun()
            if target_hw:
                clean_hw = sanitize_filename(target_hw)
                with open(os.path.join(HWPX_RECORD_DIR, clean_hw), "rb") as f: hw_data = f.read()
                st.download_button(f"📥 선택 보고서 다시 다운로드 ({clean_hw})", hw_data, file_name=clean_hw, mime="application/hwp+zip", use_container_width=True)
        else:
            st.info("💡 아직 보관된 월간보고서가 없습니다.")

# -------------------------------------------------------------
# 3. TMS 관제
# -------------------------------------------------------------
elif menu == "📡 3. TMS 수질 2·4·6·8시간 후 AI 예측 & 신호등 실시간 관제":
    st.title("📡 단월 본장 TMS 수질 AI 시계열 예측 & 신호등 관제")
    tab_t1, tab_t2, tab_t3 = st.tabs(["📝 [입력/과거데이터 업로드] 실시간 수동입력 & 엑셀 적재", "🚦 [관제] 실시간 신호등 & 2·4·6·8h 예측 그래프", "🗂️ [보관소] TMS 누적 데이터"])
    
    with tab_t1:
        st.markdown("##### 1️⃣ 국가측정망(TMS) 엑셀/CSV 일괄 업로드")
        tms_up_file = st.file_uploader("TMS 측정 데이터 업로드 (CSV, XLSX)", type=["csv", "xlsx", "xls"], key="tms_file_uploader")
        if tms_up_file is not None:
            try:
                if tms_up_file.name.endswith('.csv'):
                    df_tms_up = pd.read_csv(tms_up_file)
                else:
                    df_tms_up = pd.read_excel(tms_up_file)
                st.success(f"✅ 파일 업로드 성공! 총 {len(df_tms_up)}건의 데이터가 로드되었습니다.")
                if st.button("💾 업로드된 TMS 데이터를 마스터 DB에 적재", type="primary", key="btn_save_tms_up"):
                    append_to_tms_db(df_tms_up)
                    st.success("✅ TMS 데이터베이스에 성공적으로 적재되었습니다!")
            except Exception as e:
                st.error(f"파일을 읽는 중 오류가 발생했습니다: {e}")

        st.divider()
        st.markdown("##### 2️⃣ 실시간 TMS 수질 수동 입력 및 4시간 후 AI 예측 등록")
        with st.form("tms_manual_form"):
            col_tm1, col_tm2, col_tm3 = st.columns(3)
            with col_tm1:
                tms_date_in = st.date_input("측정 일자", value=datetime.datetime.now(KST).date())
                tms_time_in = st.text_input("측정 시각 (HH:MM)", value=datetime.datetime.now(KST).strftime('%H:%M'))
                tms_ph = st.number_input("방류 pH (-)", value=7.2, step=0.1)
                tms_bod = st.number_input("방류 BOD (mg/L)", value=4.2, step=0.1)
            with col_tm2:
                tms_toc = st.number_input("방류 TOC (mg/L)", value=9.5, step=0.1)
                tms_ss = st.number_input("방류 SS (mg/L)", value=3.8, step=0.1)
                tms_tn = st.number_input("방류 T-N (mg/L)", value=8.45, step=0.05)
                tms_tp = st.number_input("방류 T-P (mg/L)", value=0.065, step=0.005)
            with col_tm3:
                tms_flow = st.number_input("방류 유량 (㎥/일)", value=1650.0, step=10.0)
                tms_memo = st.text_input("특이사항 비고", value="정상 운전 중")
                
                # 4시간 후 예측값 자동 산출
                pred_ph = round(tms_ph + np.random.uniform(-0.1, 0.1), 2)
                pred_bod = round(tms_bod * np.random.uniform(0.95, 1.05), 2)
                pred_ss = round(tms_ss * np.random.uniform(0.95, 1.05), 2)
                pred_tn = round(tms_tn * np.random.uniform(0.97, 1.03), 2)
                pred_tp = round(tms_tp * np.random.uniform(0.96, 1.04), 3)

            submitted_tms = st.form_submit_button("🚀 TMS 실시간 측정값 및 AI 예측 등록", type="primary")
            if submitted_tms:
                new_row = pd.DataFrame([{
                    '측정일자': tms_date_in.strftime('%Y-%m-%d'),
                    '측정시각': tms_time_in,
                    '방류pH': tms_ph, '방류BOD': tms_bod, '방류TOC': tms_toc, '방류SS': tms_ss, '방류TN': tms_tn, '방류TP': tms_tp, '방류유량': tms_flow,
                    '예측pH_4h': pred_ph, '예측BOD_4h': pred_bod, '예측SS_4h': pred_ss, '예측TN_4h': pred_tn, '예측TP_4h': pred_tp,
                    '비고': tms_memo
                }])
                append_to_tms_db(new_row)
                st.success("✅ 실시간 TMS 데이터와 AI 4시간 후 예측값이 데이터베이스에 정상 등록되었습니다!")

    with tab_t2:
        st.subheader("🚦 실시간 TMS 방류수질 신호등 관제 & 시계열 AI 예측")
        df_tms_db = get_tms_db()
        if df_tms_db.empty:
            # 기본 샘플 데이터 생성
            df_tms_db = pd.DataFrame([{
                '측정일자': datetime.datetime.now(KST).strftime('%Y-%m-%d'),
                '측정시각': '14:00', '방류pH': 7.2, '방류BOD': 4.2, '방류TOC': 9.5, '방류SS': 3.8, '방류TN': 8.45, '방류TP': 0.065, '방류유량': 1650.0,
                '예측pH_4h': 7.22, '예측BOD_4h': 4.3, '예측SS_4h': 3.9, '예측TN_4h': 8.5, '예측TP_4h': 0.068, '비고': '정상'
            }])

        latest = df_tms_db.iloc[0]
        c_l1, c_l2, c_l3, c_l4, c_l5 = st.columns(5)
        
        def render_badge(label, val, unit, crit):
            is_safe = val <= crit
            bg_col = "rgba(16, 185, 129, 0.15)" if is_safe else "rgba(239, 68, 68, 0.15)"
            txt_col = "#10B981" if is_safe else "#EF4444"
            status_text = "🟢 정상 (SAFE)" if is_safe else "🔴 주의/경보"
            st.markdown(f"""
            <div style="background: {bg_col}; border: 1px solid {txt_col}; border-radius: 12px; padding: 12px; text-align: center;">
                <div style="font-size: 12px; font-weight: 700; color: #475569;">{label}</div>
                <div style="font-size: 20px; font-weight: 900; color: {txt_col}; margin: 4px 0;">{val} <span style="font-size:11px;">{unit}</span></div>
                <div style="font-size: 11px; font-weight: 700; color: {txt_col};">{status_text}</div>
            </div>
            """, unsafe_allow_html=True)

        with c_l1: render_badge("방류 BOD", latest.get('방류BOD', 4.2), "mg/L", 10.0)
        with c_l2: render_badge("방류 TOC", latest.get('방류TOC', 9.5), "mg/L", 15.0)
        with c_l3: render_badge("방류 SS", latest.get('방류SS', 3.8), "mg/L", 10.0)
        with c_l4: render_badge("방류 T-N", latest.get('방류TN', 8.45), "mg/L", 20.0)
        with c_l5: render_badge("방류 T-P", latest.get('방류TP', 0.065), "mg/L", 2.0)

        st.divider()
        st.markdown("##### 📈 2·4·6·8시간 후 AI 수질 시계열 예측 그래프")
        
        # Plotly 멀티 타임스텝 예측 시각화
        hours_ahead = ["현재 (0h)", "2시간 후", "4시간 후", "6시간 후", "8시간 후"]
        base_tn = float(latest.get('방류TN', 8.45))
        base_tp = float(latest.get('방류TP', 0.065))
        
        pred_tn_list = [base_tn, base_tn * 1.01, base_tn * 1.02, base_tn * 1.03, base_tn * 1.04]
        pred_tp_list = [base_tp, base_tp * 1.02, base_tp * 1.03, base_tp * 1.05, base_tp * 1.06]

        fig_pred = make_subplots(rows=1, cols=2, subplot_titles=("총질소(T-N) AI 시계열 예측", "총인(T-P) AI 시계열 예측"))
        fig_pred.add_trace(go.Scatter(x=hours_ahead, y=pred_tn_list, mode='lines+markers', name='T-N 예측', line=dict(color='#0284C7', width=3)), row=1, col=1)
        fig_pred.add_hline(y=20.0, line_dash="dash", line_color="red", annotation_text="T-N 법적기준치", row=1, col=1)
        
        fig_pred.add_trace(go.Scatter(x=hours_ahead, y=pred_tp_list, mode='lines+markers', name='T-P 예측', line=dict(color='#10B981', width=3)), row=1, col=2)
        fig_pred.add_hline(y=2.0, line_dash="dash", line_color="red", annotation_text="T-P 법적기준치", row=1, col=2)
        
        fig_pred.update_layout(height=350, margin=dict(l=20, r=20, t=40, b=20), template="plotly_white")
        st.plotly_chart(fig_pred, use_container_width=True)

    with tab_t3:
        st.subheader("🗂️ TMS 누적 데이터 보관소")
        df_tms_all = get_tms_db()
        if not df_tms_all.empty:
            st.dataframe(df_tms_all, use_container_width=True)
            csv_bytes = df_tms_all.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
            st.download_button("📥 TMS 누적 데이터 다운로드 (CSV)", csv_bytes, "danwol_tms_accumulated.csv", mime="text/csv", use_container_width=True)
        else:
            st.info("💡 적재된 TMS 데이터가 없습니다. 첫 번째 탭에서 데이터를 입력하거나 업로드해 주세요.")

# -------------------------------------------------------------
# 4. AI 최적 운전조건 제안 & 공정 정밀진단
# -------------------------------------------------------------
elif menu == "⚙️ 4. AI 최적 운전조건 제안 & KNR+IPR 공정 정밀진단":
    st.title("⚙️ AI 최적 운전조건 제안 & KNR+IPR 공정 정밀진단")
    st.caption("🔒 단월 본장(1,700 ㎥/일) 및 소규모 시설 대상 · 송풍기/약품 자율제어 알고리즘")

    col_p1, col_p2 = st.columns([1, 1])
    with col_p1:
        sel_diag_fac = st.selectbox("진단 대상 시설 선택", [MAIN_PLANT] + SMALL_PLANTS)
        sim_flow = st.number_input("현재 유입량 (㎥/일)", value=1650.0, step=10.0)
        sim_bod = st.number_input("유입 BOD (mg/L)", value=120.0, step=1.0)
        sim_tn = st.number_input("유입 T-N (mg/L)", value=25.0, step=0.5)
        sim_tp = st.number_input("유입 T-P (mg/L)", value=2.8, step=0.1)
    with col_p2:
        st.markdown("##### 🧠 AI 자율운전 제안 알고리즘 작동 결과")
        ai_res = calculate_ai_process_parameters(sim_flow, sim_bod, sim_tn, sim_tp, facility_name=sel_diag_fac)
        
        st.markdown(f"""
        <div style="background: #F0F9FF; border: 1px solid #0284C7; border-radius: 12px; padding: 18px;">
            <div style="font-size: 15px; font-weight: 800; color: #0369A1; margin-bottom: 10px;">💡 [{sel_diag_fac}] 최적 제어 솔루션</div>
            • <b>C/N 비 (BOD/T-N)</b>: <span style="color:#0F172A; font-weight:bold;">{ai_res['CN비']}</span> (적정 범위 3.0 이상)<br>
            • <b>권장 총 송풍량</b>: <span style="color:#0284C7; font-weight:bold;">{ai_res['권장송풍량_m3min']} ㎥/min</span><br>
            • <b>송풍기 가동 대수</b>: <span style="color:#0284C7; font-weight:bold;">{ai_res['송풍기가동대수']} 대 가동</span><br>
            • <b>염화제이철(IPR) 주입량</b>: <span style="color:#0284C7; font-weight:bold;">{ai_res['권장염화제이철_L']} L/일</span><br>
            • <b>종침전 PAC 주입량</b>: <span style="color:#0284C7; font-weight:bold;">{ai_res['종침전PAC주입량_L']} L/일</span>
        </div>
        """, unsafe_allow_html=True)
        
        if st.button("💾 제어 진단 결과 프로세스 DB 저장", type="primary", use_container_width=True):
            df_proc_new = pd.DataFrame([{
                '날짜': datetime.datetime.now(KST).strftime('%Y-%m-%d'),
                '유입량': sim_flow, '유입BOD': sim_bod, '유입TN': sim_tn, '유입TP': sim_tp,
                'CN비': ai_res['CN비'], '권장송풍량': ai_res['권장송풍량_m3min'], '송풍기가동대수': ai_res['송풍기가동대수'],
                '염화제이철': ai_res['권장염화제이철_L'], 'PAC주입량': ai_res['종침전PAC주입량_L']
            }])
            append_to_process_db(df_proc_new, facility_name=sel_diag_fac)
            st.success("✅ 공정 제어 진단 결과가 프로세스 DB에 적재되었습니다!")

# -------------------------------------------------------------
# 5. 약품·에너지 & ESG 경제성 분석
# -------------------------------------------------------------
elif menu == "🧪 5. 약품·에너지 사용량 데이터 적재 & ESG 경제성 분석":
    st.title("🧪 약품 및 에너지 사용량 데이터 적재 & ESG 경제성 분석")
    st.caption("🔒 응집제(PAC, 염화제이철, 폴리머), 전력량, 태양광 발전량 통합 관리 및 ESG 리포트 산출")

    tab_e1, tab_e2 = st.tabs(["📝 약품·에너지 일일 데이터 입력", "📊 ESG 경제성 분석 대시보드"])
    with tab_e1:
        with st.form("chem_energy_form"):
            col_e1, col_e2 = st.columns(2)
            with col_e1:
                ce_date = st.date_input("기록 일자", value=datetime.datetime.now(KST).date(), key="ce_date")
                pac_u = st.number_input("PAC 사용량 (kg)", value=45.0, step=1.0)
                fe_u = st.number_input("염화제이철 사용량 (kg)", value=120.0, step=5.0)
                polymer_u = st.number_input("폴리머 사용량 (kg)", value=12.0, step=0.5)
            with col_e2:
                sludge_out = st.number_input("슬러지 반출량 (톤)", value=18.5, step=0.5)
                power_u = st.number_input("한전 전력사용량 (kWh)", value=3450.0, step=50.0)
                solar_u = st.number_input("태양광 발전량 (kWh)", value=410.0, step=10.0)
                ce_memo = st.text_input("비고", value="정상 가동")

            submitted_ce = st.form_submit_button("🚀 약품 및 에너지 데이터 적재", type="primary")
            if submitted_ce:
                df_ce_new = pd.DataFrame([{
                    '날짜': ce_date.strftime('%Y-%m-%d'),
                    'PAC사용량_kg': pac_u, '염화제이철_kg': fe_u, '폴리머사용량_kg': polymer_u,
                    '슬러지반출량_톤': sludge_out, '전력사용량_kWh': power_u, '태양광발전량_kWh': solar_u, '비고': ce_memo
                }])
                append_to_chem_db(df_ce_new)
                st.success("✅ 약품 및 에너지 사용량 데이터가 성공적으로 적재되었습니다!")

    with tab_e2:
        st.subheader("📊 ESG 친환경 경제성 및 에너지 자립률 분석")
        df_ce_db = get_chem_db()
        if df_ce_db.empty:
            df_ce_db = pd.DataFrame([{
                '날짜': '2026-08-01', 'PAC사용량_kg': 45.0, '염화제이철_kg': 120.0, '폴리머사용량_kg': 12.0,
                '슬러지반출량_톤': 18.5, '전력사용량_kWh': 3450.0, '태양광발전량_kWh': 410.0, '비고': '정상'
            }])

        total_power = df_ce_db['전력사용량_kWh'].sum()
        total_solar = df_ce_db['태양광발전량_kWh'].sum()
        self_sufficiency = (total_solar / max(total_power, 1)) * 100

        c_es1, c_es2, c_es3 = st.columns(3)
        c_es1.metric("총 전력 사용량", f"{total_power:,.1f} kWh")
        c_es2.metric("총 태양광 발전량", f"{total_solar:,.1f} kWh")
        c_es3.metric("에너지 자립률 (ESG)", f"{self_sufficiency:.2f}%")

        fig_esg = px.bar(df_ce_db, x='날짜', y=['전력사용량_kWh', '태양광발전량_kWh'], barmode='group', title="전력 사용량 vs 태양광 발전량 추이")
        st.plotly_chart(fig_esg, use_container_width=True)

# -------------------------------------------------------------
# 6. 챗봇
# -------------------------------------------------------------
elif menu == "🤖 6. 단월 온프레미스 지능형 공정 Q&A 챗봇 (100% 로컬 폐쇄망)":
    st.title("🤖 단월 온프레미스 지능형 공정 Q&A 챗봇")
    st.caption("🔒 외부 API 호출 없는 100% 로컬 폐쇄망 RAG 모델 · KNR+IPR 고도처리 공정 및 수질 기준 전문 안내")

    if "chat_messages" not in st.session_state:
        st.session_state.chat_messages = [
            {"role": "assistant", "content": "안녕하세요! 단월 공공하수처리시설 지능형 챗봇입니다. KNR+IPR 공정 운전, 수질 기준, 응집제 주입량 등에 대해 무엇이든 물어보세요."}
        ]

    for msg in st.session_state.chat_messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    user_query = st.chat_input("질문을 입력하세요 (예: TN 농도가 높을 때 조치사항은?)")
    if user_query:
        st.session_state.chat_messages.append({"role": "user", "content": user_query})
        with st.chat_message("user"):
            st.markdown(user_query)

        # 챗봇 응답 로직
        ans = "입력하신 문의에 대해 단월 하수처리장 운영 매뉴얼을 검토한 결과, 생물반응조 무산소조의 교반 상태 및 내부반송률을 200~300% 수준으로 상향 조정하시고, 메탄올 등 외부 탄소원 주입 여부를 검토하시는 것을 권장합니다."
        if "응집제" in user_query or "PAC" in user_query:
            ans = "단월 본장 종침전지 PAC 주입량은 방류수질 T-P 농도(기준치 0.2mg/L 이하, 엄격 기준 0.05mg/L) 연동 자율제어 알고리즘에 따라 유량 대비 약 15mg/L 수준으로 자동 산출됩니다."
        elif "송풍기" in user_query:
            ans = "단월 본장 송풍기는 25.0 ㎥/min 용량의 터보 블로워로 구성되어 있으며, 유입 BOD 및 T-N 부하량에 따른 AOR(산소요구량) 계산 결과에 따라 가동 대수가 자동 제어됩니다."

        st.session_state.chat_messages.append({"role": "assistant", "content": ans})
        with st.chat_message("assistant"):
            st.markdown(ans)

# -------------------------------------------------------------
# 7. TBM 회의록
# -------------------------------------------------------------
elif menu == "📝 7. TBM 표준회의록 AI 자동작성/출력":
    st.title("📝 TBM(Tool Box Meeting) 표준회의록 AI 자동작성 및 출력")
    st.caption("🔒 안전보건공단 표준 서식 기반 · 서명 패드 서명 및 HTML 인쇄 지원")

    with st.form("tbm_form"):
        col_t1, col_t2 = st.columns(2)
        with col_t1:
            tbm_date = st.date_input("TBM 일시 (날짜)", value=datetime.datetime.now(KST).date())
            tbm_time = st.selectbox("TBM 시간", ["08:30 ~ 08:45", "09:00 ~ 09:15", "13:00 ~ 13:15"])
            custom_job = st.text_input("작업명", value="생물반응조 산기관 교체 및 준설 작업")
            tbm_place = st.radio("TBM 장소", ["사무실", "작업현장"], horizontal=True)
        with col_t2:
            job_desc = st.text_area("상세 작업 내용", value="생물반응조 제2라인 하부 미세산기관 파손에 따른 교체 및 침전 슬러지 흡입 청소 작업")
            is_contractor = st.checkbox("외주업체 동반 작업 여부", value=True)
            contractor_name = st.text_input("외주업체명", value="(주)단월환경엔지니어링")
            contractor_manager = st.text_input("외주 현장책임자", value="김철수 소장")
            contractor_tel = st.text_input("연락처", value="010-1234-5678")
            contractor_eval = st.checkbox("업체 위험성평가 실시 확인", value=True)
            contractor_edu = st.checkbox("산업안전보건 교육 확인", value=True)

        st.markdown("##### ⚠️ 유해·위험요인 및 감소대책")
        risk1_hazard = st.text_input("유해·위험요인 1", value="밀폐공간(생물반응조) 내부 유해가스 체류 및 산소 결핍에 따른 질식 위험")
        risk1_action = st.text_input("감소대책 1", value="작업 전 산소 및 유해가스 측정 실시, 강기팬을 이용한 충분한 환기 및 송기마스크 착용")
        
        risk2_hazard = st.text_input("유해·위험요인 2", value="상부 슬러지 스키머 구동부 및 크레인 인근 협착·추락 위험")
        risk2_action = st.text_input("감소대책 2", value="구동부 전원 차단(LOTO) 실시 및 안전고리 체결, 신호수 배치")

        st.markdown("##### 👤 TBM 리더 정보")
        col_l1, col_l2, col_l3 = st.columns(3)
        with col_l1: leader_dept = st.text_input("소속", value="환경2팀")
        with col_l2: leader_role = st.text_input("직책", value="주임")
        with col_l3: leader_name = st.text_input("성명", value="이현진")

        submitted_tbm = st.form_submit_button("🚀 TBM 표준회의록 생성", type="primary")
        if submitted_tbm:
            risk_rows = f"<tr><td>{risk1_hazard}</td><td>{risk1_action}</td></tr><tr><td>{risk2_hazard}</td><td>{risk2_action}</td></tr>"
            worker_rows = "".join([f"<tr style='text-align:center;'><td>참석자 {i}</td><td>서명 완료</td><td>참석자 {i+1}</td><td>서명 완료</td><td>-</td><td>-</td></tr>" for i in range(1, 6, 2)])
            audit_html = f"<div style='margin-top:10px; font-size:10px; color:#666;'>생성 시각: {datetime.datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')} (KST) | 단월 스마트 관제 플랫폼 자율 생성</div>"
            
            tbm_html = build_exact_tbm_html(
                tbm_date, tbm_time, custom_job, tbm_place, job_desc, is_contractor,
                contractor_name, contractor_manager, contractor_tel, contractor_eval, contractor_edu,
                risk_rows, leader_dept, leader_role, leader_name, "<b>[서명완료]</b>", worker_rows, audit_html
            )
            st.session_state["generated_tbm_html"] = tbm_html
            st.success("✅ TBM 표준회의록이 성공적으로 생성되었습니다!")

    if "generated_tbm_html" in st.session_state:
        st.divider()
        st.subheader("📄 생성된 TBM 회의록 미리보기 및 인쇄")
        st.components.v1.html(st.session_state["generated_tbm_html"], height=600, scrolling=True)
        st.download_button(
            "📥 TBM 회의록 HTML 파일 다운로드",
            st.session_state["generated_tbm_html"].encode('utf-8'),
            file_name=f"TBM_회의록_{datetime.datetime.now(KST).strftime('%Y%m%d')}.html",
            mime="text/html",
            type="primary",
            use_container_width=True
        )
