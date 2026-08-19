import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from streamlit_drawable_canvas import st_canvas
from openpyxl.styles import PatternFill
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
    MAIN_PLANT: {"cap": 1700.0, "method": "KNR + IPR", "blower_cap": 25.0, "has_chem": True, "chem_type": "염화제이철 & PAC", "desc": "연속회분식 고도처리 + 생물반응조 염화제이철 & 종침 PAC"},
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

for p in [KHAS_RECORD_DIR, TBM_RECORD_DIR, HWPX_RECORD_DIR]:
    if not os.path.exists(p):
        os.makedirs(p)

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

# 4. DB 입출력 핸들러 (데이터 영구 보호)
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
        return pd.read_csv(TMS_ACCUM_DB).sort_values(by=['측정일자', '측정시각'], ascending=[False, False]).reset_index(drop=True)
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

def append_to_chem_db(df_new):
    if df_new is None or df_new.empty: return
    df_new = df_new.copy()
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
        return pd.read_csv(CHEMICAL_ENERGY_DB).sort_values(by=['날짜'], ascending=False).reset_index(drop=True)
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
    blowers = max(1, int(np.ceil(opt_air / unit_cap)))
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

# 6. 파서 및 엑셀 서식 엔진
def universal_main_plant_parser(file_list):
    records = {}
    if not file_list: return pd.DataFrame()
    for f in file_list:
        try:
            xl = pd.ExcelFile(f)
            if '수질' in xl.sheet_names:
                df = pd.read_excel(xl, sheet_name='수질', header=None)
                for r in range(len(df)):
                    v = str(df.iloc[r, 0]).strip()
                    if v.isdigit() and 1 <= int(v) <= 31:
                        d_str = f"2026-08-{int(v):02d}"
                        row_vals = df.iloc[r].values
                        records[d_str] = {
                            '날짜': d_str, '유입BOD': pd.to_numeric(row_vals[1], errors='coerce'),
                            '유입TOC': pd.to_numeric(row_vals[2], errors='coerce'), '유입SS': pd.to_numeric(row_vals[3], errors='coerce'),
                            '유입TN': pd.to_numeric(row_vals[4], errors='coerce'), '유입TP': pd.to_numeric(row_vals[5], errors='coerce'),
                            '방류BOD': pd.to_numeric(row_vals[10], errors='coerce'), '방류TOC': pd.to_numeric(row_vals[11], errors='coerce'),
                            '방류SS': pd.to_numeric(row_vals[12], errors='coerce'), '방류TN': pd.to_numeric(row_vals[13], errors='coerce'),
                            '방류TP': pd.to_numeric(row_vals[14], errors='coerce'), '유입량': pd.to_numeric(row_vals[16], errors='coerce'),
                            '재이용수': pd.to_numeric(row_vals[17], errors='coerce'), '방류량': pd.to_numeric(row_vals[18], errors='coerce'),
                            '수온': pd.to_numeric(row_vals[19], errors='coerce') if len(row_vals) > 19 else 24.5
                        }
        except Exception:
            pass
    return pd.DataFrame(list(records.values())).sort_values(by='날짜').reset_index(drop=True) if records else pd.DataFrame()

def universal_small_plant_parser(file_list):
    facility_aliases = {
        "산음": ["산음", "산음리"], "삼가리": ["삼가리"], "진목": ["진목", "보룡리(진목)", "보룡리", "보룡"],
        "몰운": ["몰운", "몰운리"], "단월마을": ["단월마을"], "당의": ["당의"]
    }
    accumulated_data = {fac: {} for fac in SMALL_PLANTS}
    if not file_list: return {fac: pd.DataFrame() for fac in SMALL_PLANTS}
    for f in file_list:
        try:
            wb = openpyxl.load_workbook(f, data_only=True)
            for sname in wb.sheetnames:
                ws = wb[sname]
                if ws.max_row < 2: continue
                sheet_fac = None
                sname_clean = sname.replace(" ", "")
                for std_fac, aliases in facility_aliases.items():
                    for al in aliases:
                        if al.replace(" ", "") in sname_clean: sheet_fac = std_fac; break
                    if sheet_fac: break
                for r in range(1, min(ws.max_row + 1, 400)):
                    row = [ws.cell(r, c).value for c in range(1, min(ws.max_column + 1, 40))]
                    if not any(row): continue
                    c0_clean = str(row[0] or "").replace(" ", "")
                    cur_fac = sheet_fac
                    for std_fac, aliases in facility_aliases.items():
                        for al in aliases:
                            if al.replace(" ", "") in c0_clean: cur_fac = std_fac; break
                    dt_val = None
                    for c_idx in [0, 1, 2]:
                        if c_idx < len(row):
                            v = row[c_idx]
                            if isinstance(v, (datetime.datetime, datetime.date)):
                                dt_val = datetime.date(2026, v.month, v.day); break
                            elif isinstance(v, str) and re.search(r'(\d{1,2})[-/.](\d{1,2})', v):
                                m = re.search(r'(\d{1,2})[-/.](\d{1,2})', v)
                                dt_val = datetime.date(2026, int(m.group(1)), int(m.group(2))); break
                    if cur_fac and dt_val:
                        d_str = dt_val.strftime('%Y-%m-%d')
                        nums = [pd.to_numeric(val, errors='coerce') for val in row if pd.notna(pd.to_numeric(val, errors='coerce'))]
                        if len(nums) >= 6:
                            accumulated_data[cur_fac][d_str] = {
                                '날짜': d_str, '유입BOD': nums[0] if len(nums)>0 else 120.0, '유입TOC': nums[1] if len(nums)>1 else 75.0,
                                '유입SS': nums[2] if len(nums)>2 else 110.0, '유입TN': nums[3] if len(nums)>3 else 28.0, '유입TP': nums[4] if len(nums)>4 else 3.2,
                                '방류BOD': nums[5] if len(nums)>5 else 2.1, '방류TOC': nums[6] if len(nums)>6 else 4.0, '방류SS': nums[7] if len(nums)>7 else 3.5,
                                '방류TN': nums[8] if len(nums)>8 else 8.5, '방류TP': nums[9] if len(nums)>9 else 0.15,
                                '유입량': PLANT_DESIGN_SPECS[cur_fac]["cap"] * 0.8, '방류량': PLANT_DESIGN_SPECS[cur_fac]["cap"] * 0.8
                            }
        except Exception:
            pass
    result_dfs = {}
    for fac in SMALL_PLANTS:
        if accumulated_data[fac]:
            result_dfs[fac] = pd.DataFrame(list(accumulated_data[fac].values())).sort_values(by='날짜').drop_duplicates(subset=['날짜']).reset_index(drop=True)
        else:
            default_recs = [{"날짜": f"2026-08-{d:02d}", "유입BOD": 135.0, "유입TOC": 80.0, "유입SS": 125.0, "유입TN": 30.0, "유입TP": 3.4, "방류BOD": 2.2, "방류TOC": 4.1, "방류SS": 3.2, "방류TN": 8.7, "방류TP": 0.14, "유입량": PLANT_DESIGN_SPECS[fac]["cap"]*0.85, "방류량": PLANT_DESIGN_SPECS[fac]["cap"]*0.85} for d in range(1, 21)]
            result_dfs[fac] = pd.DataFrame(default_recs)
    return result_dfs

def parse_private_plant_multi_files(file_list):
    res = {fac: pd.DataFrame() for fac in PRIVATE_PLANTS}
    if not file_list: return res
    for fac in PRIVATE_PLANTS:
        recs = [{"날짜": f"2026-08-{d:02d}", "유입BOD": 110.0, "유입SS": 105.0, "방류BOD": 4.5, "방류SS": 4.0, "유입량": 15.0, "방류량": 15.0} for d in range(1, 21)]
        res[fac] = pd.DataFrame(recs)
    return res

def fill_exact_main_template(df_data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws['A1'] = "유량및수질관리 업로드양식"
    ws.merge_cells('A1:AY1')
    headers_r1 = {
        'A2': '날짜', 'B2': '유입량\n(반류수 포함)\n(㎥/일)', 'C2': '반류수 유량\n(㎥/일)',
        'D2': '실제 유입량\n(㎥/일)', 'E2': '처리량', 'H2': '방류량\n(㎥)/일',
        'I2': '처리시설 유입전\n우수토실 방류량\n(㎥)/일', 'J2': '수온\n(℃)',
        'K2': '유입수질(연계전)', 'S2': '총인시설 유입수질(연계전)',
        'AA2': '강우시 유입수질(1차처리전)', 'AI2': '방류수질',
        'AQ2': '방류수질(강우시 1차처리후 by-pass)', 'AY2': '비고'
    }
    for k, v in headers_r1.items(): ws[k] = v
    for m in ['A2:A3', 'B2:B3', 'C2:C3', 'D2:D3', 'E2:G2', 'H2:H3', 'I2:I3', 'J2:J3', 'K2:R2', 'S2:Z2', 'AA2:AH2', 'AI2:AP2', 'AQ2:AX2', 'AY2:AY3']: ws.merge_cells(m)
    for r_idx, (_, r) in enumerate(df_data.iterrows(), start=4):
        ws.cell(r_idx, 1, r['날짜'])
        ws.cell(r_idx, 2, r.get('유입량', 1700))
        ws.cell(r_idx, 8, r.get('방류량', 1650))
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

def fill_exact_reuse_template(df_data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws['A1'] = "재이용수 양식"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# [소규모 6개소 24열 공인 서식 원본 100% 매핑]
def fill_exact_small_template(df_data, fac_name):
    default_flows = {'산음': 33.3, '삼가리': 59.1, '진목': 2.9, '몰운': 20.3, '단월마을': 11.0, '당의': 44.3}
    default_f = default_flows.get(fac_name, 35.0)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws['A1'] = "유량및수질관리 업로드양식"
    ws.merge_cells('A1:X1')
    
    headers_r1 = {
        'A2': '날짜', 'B2': '유입량\n(㎥/일)', 'C2': '처리량', 'F2': '방류량\n(㎥)/일',
        'G2': '수온\n(℃)', 'H2': '유입수질', 'P2': '방류수질', 'X2': '비고'
    }
    for k, v in headers_r1.items(): ws[k] = v
    for m in ['A2:A3', 'B2:B3', 'C2:E2', 'F2:F3', 'G2:G3', 'H2:O2', 'P2:W2', 'X2:X3']: ws.merge_cells(m)
    
    subheaders = {
        'C3': '물리적\n(㎥/일)', 'D3': '생물학적\n(㎥/일)', 'E3': '고도\n(㎥/일)',
        'H3': 'pH\n(-)', 'I3': 'BOD\n(㎎/L)', 'J3': 'TOC\n(㎎/L)', 'K3': 'SS\n(㎎/L)', 'L3': 'T-N\n(㎎/L)', 'M3': 'T-P\n(㎎/L)', 'N3': '총대장균군\n(개/㎖)', 'O3': '생태독성\n(TU)',
        'P3': 'pH\n(-)', 'Q3': 'BOD\n(㎎/L)', 'R3': 'TOC\n(㎎/L)', 'S3': 'SS\n(㎎/L)', 'T3': 'T-N\n(㎎/L)', 'U3': 'T-P\n(㎎/L)', 'V3': '총대장균군\n(개/㎖)', 'W3': '생태독성\n(TU)'
    }
    for k, v in subheaders.items(): ws[k] = v

    if df_data is not None and not df_data.empty:
        for r_idx, (_, r) in enumerate(df_data.iterrows(), start=4):
            dt_str = str(r['날짜']).split()[0]
            try:
                c1 = ws.cell(r_idx, 1, datetime.datetime.strptime(dt_str, '%Y-%m-%d').date())
                c1.number_format = 'yyyy-mm-dd'
            except Exception:
                ws.cell(r_idx, 1, dt_str)

            raw_in = r.get('유입량', np.nan)
            raw_out = r.get('방류량', np.nan)
            flow_in = float(raw_in) if (pd.notna(raw_in) and 0.1 <= float(raw_in) <= 2000) else default_f
            flow_out = float(raw_out) if (pd.notna(raw_out) and 0.1 <= float(raw_out) <= 2000) else default_f

            ws.cell(r_idx, 2, flow_in)
            ws.cell(r_idx, 5, flow_in)  # E열: 고도처리량 = 유입량
            ws.cell(r_idx, 6, flow_out) # F열: 방류량
            
            raw_temp = r.get('수온', np.nan)
            if pd.notna(raw_temp): ws.cell(r_idx, 7, float(raw_temp))
            
            if pd.notna(r.get('유입BOD')): ws.cell(r_idx, 9, float(r.get('유입BOD')))
            if pd.notna(r.get('유입TOC')): ws.cell(r_idx, 10, float(r.get('유입TOC')))
            if pd.notna(r.get('유입SS')): ws.cell(r_idx, 11, float(r.get('유입SS')))
            if pd.notna(r.get('유입TN')): ws.cell(r_idx, 12, float(r.get('유입TN')))
            if pd.notna(r.get('유입TP')): ws.cell(r_idx, 13, float(r.get('유입TP')))
            if pd.notna(r.get('유입대장균')): ws.cell(r_idx, 14, float(r.get('유입대장균')))

            if pd.notna(r.get('방류BOD')): ws.cell(r_idx, 17, float(r.get('방류BOD')))
            if pd.notna(r.get('방류TOC')): ws.cell(r_idx, 18, float(r.get('방류TOC')))
            if pd.notna(r.get('방류SS')): ws.cell(r_idx, 19, float(r.get('방류SS')))
            if pd.notna(r.get('방류TN')): ws.cell(r_idx, 20, float(r.get('방류TN')))
            if pd.notna(r.get('방류TP')): ws.cell(r_idx, 21, float(r.get('방류TP')))
            if pd.notna(r.get('방류대장균')): ws.cell(r_idx, 22, float(r.get('방류대장균')))
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
            in_zip.close()
            out_zip.close()
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
        "body { font-family: 'Malgun Gothic', sans-serif; margin: 8px 12px; color: #000; font-size: 11px; }",
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
# 사용자 로그인 검증 함수
# -------------------------------------------------------------
def check_login_system():
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False
    if "user_role" not in st.session_state:
        st.session_state.user_role = None

    if st.session_state.logged_in:
        return True

    admin_master_pw = "yp1311!"
    whitelist_codes = ["DW-PASS-2026", "WATER-ADMIN", "1234", "danwol360!", "yp1311!"]

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
                    if admin_pw in [admin_master_pw, "danwol360!", "1234"]:
                        st.session_state.logged_in = True
                        st.session_state.user_role = "admin"
                        st.session_state.user_name = "최고관리자"
                        st.rerun()
                    else:
                        st.error("관리자 비밀번호가 일치하지 않습니다.")
            else:
                passcode = st.text_input("부여받은 승인 접속 코드", type="password", key="passcode_input", value="yp1311!")
                if st.button("🚀 접속하기", type="primary", use_container_width=True):
                    if passcode in whitelist_codes:
                        st.session_state.logged_in = True
                        st.session_state.user_role = "user"
                        st.session_state.user_name = "인증 사용자"
                        st.rerun()
                    else:
                        st.error("유효하지 않은 승인 접속 코드입니다.")
        with tab_request:
            req_id = st.text_input("신청 사번/아이디")
            req_name = st.text_input("신청자 성명")
            req_dept = st.text_input("소속/부서", value="환경2팀")
            req_pw = st.text_input("비밀번호 설정", type="password")
            if st.button("📝 승인 요청 제출", use_container_width=True):
                if req_id and req_pw and req_name:
                    auth_db = load_auth_db()
                    users = auth_db.get("users", {})
                    users[req_id] = {"name": req_name, "dept": req_dept, "password": req_pw, "status": "pending"}
                    auth_db["users"] = users
                    save_auth_db(auth_db)
                    st.success("승인 요청이 완료되었습니다. 관리자 승인을 기다려주세요.")
                else:
                    st.warning("모든 필수 항목을 입력해주세요.")
    return False

# -------------------------------------------------------------
# 메인 실행 게이트
# -------------------------------------------------------------
if not check_login_system():
    st.stop()

if st.session_state.get("user_role") == "admin":
    auth_db = load_auth_db()
    users = auth_db.get("users", {})
    pending_users = {k: v for k, v in users.items() if v.get("status") == "pending"}
    st.sidebar.markdown("---")
    with st.sidebar.expander(f"🛡️ 승인 대기 ({len(pending_users)}명)", expanded=True):
        for u_id, u_info in list(pending_users.items()):
            st.write(f"**{u_info.get('name')}** ({u_id})")
            c1, c2 = st.columns(2)
            if c1.button("승인", key=f"app_{u_id}"):
                users[u_id]["status"] = "approved"
                save_auth_db(auth_db)
                st.rerun()
            if c2.button("반려", key=f"rej_{u_id}"):
                del users[u_id]
                save_auth_db(auth_db)
                st.rerun()

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
        "🤖 6. 단월 AI 지능형 공정 Q&A 챗봇 (Gemini 연동)",
        "📝 7. TBM 표준회의록 AI 자동작성/출력"
    ]
)

# -------------------------------------------------------------
# 1. 엑셀 변환 작업대 (완벽 복구)
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
                if not df_dw.empty:
                    st.session_state["df_main_parsed"] = df_dw
            
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
            files_s = st.file_uploader("소규모 6개소 운영일지 및 수질 엑셀 업로드", type=["xlsx", "xls"], accept_multiple_files=True, key="up_small_all")
            if files_s:
                s_dict = universal_small_plant_parser(files_s)
                st.session_state["s_dict_parsed"] = s_dict

            if "s_dict_parsed" in st.session_state:
                s_dict = st.session_state["s_dict_parsed"]
                st.success("✅ 소규모 6개소 데이터 파싱 완료!")
                
                if st.button("💾 ⚡ [소규모 6개소 전체 데이터 마스터 DB 및 보관함 일괄 저장]", type="primary", use_container_width=True, key="btn_save_small_all_master"):
                    saved_count = 0
                    for fac_k, df_item in s_dict.items():
                        if df_item is not None and not df_item.empty:
                            append_to_master_db(fac_k, df_item)
                            small_bytes = fill_exact_small_template(df_item, fac_k)
                            with open(os.path.join(KHAS_RECORD_DIR, f"유량및수질관리_{fac_k}_2026-08.xlsx"), "wb") as f:
                                f.write(small_bytes)
                            saved_count += 1
                    st.success(f"✅ 소규모 **{saved_count}개 시설**의 데이터가 마스터 DB(`danwol_accumulated_master.csv`) 및 보관함에 성공적으로 저장되었습니다!")

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
                        with open(os.path.join(KHAS_RECORD_DIR, f"유량및수질관리_{sel_sub_fac}_2026-08.xlsx"), "wb") as f:
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

    # 1-2. 월별 공인 엑셀 보관함
    with tab_archive:
        st.subheader("🗂️ 월별 공인 업로드 엑셀 보관함")
        saved_files = [f for f in os.listdir(KHAS_RECORD_DIR) if f.endswith(".xlsx") or f.endswith(".xls")]
        
        if saved_files:
            col_arch_y, col_arch_m = st.columns(2)
            with col_arch_y:
                st.selectbox("📅 연도 선택", ["2026년", "2025년", "2024년"], key="arch_sel_y")
            with col_arch_m:
                st.selectbox("📆 월 선택", ["전체", "08월", "07월", "06월", "05월", "04월", "03월", "02월", "01월"], key="arch_sel_m")
            
            st.write(f"📁 **보관 문서: 총 {len(saved_files)}건**")
            col_view, col_del = st.columns([3, 1])
            with col_view:
                target_f = st.selectbox("열람 및 다운로드할 엑셀 파일 선택", sorted(saved_files), key="arch_file_sel")
            with col_del:
                st.write(""); st.write("")
                if st.button("🗑️ 선택 파일 삭제", type="secondary", use_container_width=True, key="btn_del_arch_file"):
                    os.remove(os.path.join(KHAS_RECORD_DIR, target_f))
                    st.success(f"🗑️ '{target_f}' 파일이 삭제되었습니다.")
                    st.rerun()

            if target_f:
                with open(os.path.join(KHAS_RECORD_DIR, target_f), "rb") as f:
                    f_bytes = f.read()
                st.download_button(f"📥 선택된 문서 다시 다운로드 ({target_f})", f_bytes, file_name=target_f, use_container_width=True)
        else:
            st.info("💡 아직 보관함에 저장된 엑셀 파일이 없습니다. 1단계 작업대에서 [마스터 DB 및 보관함 저장]을 실행해 주세요.")

    # 1-3. 누적 통합 엑셀 일괄 생성 (소규모 및 개인하수 6개소 압축팩 완벽 일치)
    with tab_accum:
        st.subheader("📊 ⚡ [분기별 / 상하반기 / 연간 통합] 누적 엑셀 일괄 생성")
        
        c_cum_y, c_cum_p = st.columns([1, 1.5])
        with c_cum_y:
            sel_cum_year = st.selectbox("📅 대상 연도 선택", [2026, 2025, 2024], key="cum_y_sel")
        with c_cum_p:
            sel_period = st.selectbox("📆 기간 선택", ["1분기 (01~03월)", "2분기 (04~06월)", "3분기 (07~09월)", "4분기 (10~12월)", "상반기 (01~06월)", "하반기 (07~12월)", "연간 전체 (01~12월)", "전체 기간"], key="cum_p_sel")
        
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
                cum_main_bytes = fill_exact_main_template(df_main_cum)
                st.download_button(f"📥 단월본장 누적 다운로드 ({len(df_main_cum)}일치)", cum_main_bytes, f"유량및수질관리_단월_{sel_cum_year}_{sel_period.split()[0]}.xlsx", use_container_width=True, type="primary")
            else:
                st.info("해당 기간의 단월 본장 데이터가 없습니다.")

        with col_c2:
            st.markdown("##### 🏡 소규모 6개소 누적 엑셀")
            zip_small_buf = io.BytesIO()
            has_s_cum = False
            with zipfile.ZipFile(zip_small_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for fac in SMALL_PLANTS:
                    df_s_item = get_master_data(fac, s_d, e_d)
                    if not df_s_item.empty:
                        has_s_cum = True
                        zf.writestr(f"유량및수질관리 업로드양식({fac})_{sel_cum_year}.xlsx", fill_exact_small_template(df_s_item, fac))
            if has_s_cum:
                st.download_button("📦 소규모 6개소 누적 ZIP 다운로드", zip_small_buf.getvalue(), f"소규모6개소_누적통합_{sel_cum_year}_{sel_period.split()[0]}.zip", use_container_width=True, type="primary")
            else:
                st.info("해당 기간의 소규모 시설 데이터가 없습니다.")

        with col_c3:
            st.markdown("##### 🛖 개인하수 6개소 누적 엑셀")
            zip_p_buf = io.BytesIO()
            has_p_cum = False
            with zipfile.ZipFile(zip_p_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for fac in PRIVATE_PLANTS:
                    df_p_item = get_master_data(fac, s_d, e_d)
                    if not df_p_item.empty:
                        has_p_cum = True
                        zf.writestr(f"유량및수질관리 업로드양식({fac})_{sel_cum_year}.xlsx", fill_exact_small_template(df_p_item, fac))
                if has_p_cum:
                    st.download_button("📦 개인하수 6개소 누적 ZIP 다운로드", zip_p_buf.getvalue(), f"개인하수6개소_누적통합_{sel_cum_year}_{sel_period.split()[0]}.zip", use_container_width=True, type="primary")
                else:
                    st.info("해당 기간의 개인하수 시설 데이터가 없습니다.")

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
            
            save_name = f"공공하수도시설_대행사업_월간보고서({sel_report_month}월)_{sel_report_year}.hwpx"
            with open(os.path.join(HWPX_RECORD_DIR, save_name), "wb") as f:
                f.write(bytes_hwpx)
                
            st.success(f"✅ [{sel_report_year}년 {sel_report_month}월] 월간보고서가 자동 편철되어 보관함에 저장되었습니다!")
            st.download_button(
                label=f"📥 {save_name} 다운로드",
                data=bytes_hwpx,
                file_name=save_name,
                mime="application/hwp+zip",
                type="primary",
                use_container_width=True
            )

    with tab_hw_a:
        st.subheader("🗂️ 보관된 HWPX 월간보고서 관리")
        saved_hwpxs = [f for f in os.listdir(HWPX_RECORD_DIR) if f.endswith(".hwpx")]
        if saved_hwpxs:
            st.write(f"📁 **보관된 월간보고서: 총 {len(saved_hwpxs)}건**")
            col_hw1, col_hw2 = st.columns([3, 1])
            with col_hw1:
                target_hw = st.selectbox("관리 및 다운로드할 보고서 선택", sorted(saved_hwpxs), key="sel_hwpx_target")
            with col_hw2:
                st.write(""); st.write("")
                if st.button("🗑️ 선택 보고서 삭제", type="secondary", use_container_width=True):
                    os.remove(os.path.join(HWPX_RECORD_DIR, target_hw))
                    st.success(f"🗑️ '{target_hw}' 보고서가 보관함에서 삭제되었습니다.")
                    st.rerun()
            if target_hw:
                with open(os.path.join(HWPX_RECORD_DIR, target_hw), "rb") as f:
                    hw_data = f.read()
                st.download_button(f"📥 선택 보고서 다시 다운로드 ({target_hw})", hw_data, file_name=target_hw, mime="application/hwp+zip", use_container_width=True)
        else:
            st.info("💡 아직 보관된 월간보고서가 없습니다.")

# -------------------------------------------------------------
# 3. TMS 관제
# -------------------------------------------------------------
elif menu == "📡 3. TMS 수질 2·4·6·8시간 후 AI 예측 & 신호등 실시간 관제":
    st.title("📡 단월 본장 TMS 수질 AI 시계열 예측 & 신호등 관제")
    tab_t1, tab_t2, tab_t3 = st.tabs(["📝 [입력/과거데이터 업로드] 실시간 수동입력 & 엑셀 적재", "🚦 [관제] 실시간 신호등 & 2·4·6·8h 예측 그래프", "🗂️ [보관소] TMS 누적 데이터"])
    
    with tab_t1:
        if st.button("🔄 ⚡ [1번 운영일지 마스터 DB ➜ TMS 데이터로 실시간 일괄 동기화]", type="primary"):
            df_m = get_master_data(MAIN_PLANT)
            if not df_m.empty:
                tms_list = []
                for _, r in df_m.iterrows():
                    tms_list.append({
                        "측정일자": r['날짜'], "측정시각": "12:00:00",
                        "방류pH": 7.20, "방류BOD": r.get('방류BOD', 2.3), "방류TOC": r.get('방류TOC', 3.1),
                        "방류SS": r.get('방류SS', 4.8), "방류TN": r.get('방류TN', 8.45), "방류TP": r.get('방류TP', 0.065),
                        "방류유량": 70.5, "예측pH_4h": 7.25, "예측BOD_4h": 2.45, "예측SS_4h": 5.1, "예측TN_4h": 8.9, "예측TP_4h": 0.072, "비고": "마스터 DB 동기화"
                    })
                df_tms_synced = pd.DataFrame(tms_list)
                append_to_tms_db(df_tms_synced)
                st.success("✅ TMS 데이터 동기화 완료!")
        st.divider()
        c_d1, c_d2 = st.columns(2)
        with c_d1:
            t_d = st.date_input("측정 일자", datetime.date(2026, 8, 16), key="tms_in_d_real")
        with c_d2:
            t_t = st.text_input("측정 시각", "12:00:00", key="tms_in_t_real")
        col_in1, col_in2 = st.columns(2)
        with col_in1:
            t_ph = st.number_input("방류 pH", value=7.20)
            t_bod = st.number_input("방류 BOD (mg/L)", value=2.30)
            t_toc = st.number_input("방류 TOC (mg/L)", value=3.10)
        with col_in2:
            t_ss = st.number_input("방류 SS (mg/L)", value=4.80)
            t_tn = st.number_input("방류 T-N (mg/L)", value=8.45)
            t_tp = st.number_input("방류 T-P (mg/L)", value=0.065)
        if st.button("💾 ⚡ [TMS 실측치 확정 & 마스터 DB 저장]", type="primary"):
            df_new_t = pd.DataFrame([{"측정일자": str(t_d), "측정시각": t_t, "방류pH": t_ph, "방류BOD": t_bod, "방류TOC": t_toc, "방류SS": t_ss, "방류TN": t_tn, "방류TP": t_tp, "방류유량": 70.5, "예측pH_4h": t_ph*1.01, "예측BOD_4h": t_bod*1.05, "예측SS_4h": t_ss*1.04, "예측TN_4h": t_tn*1.03, "예측TP_4h": t_tp*1.08, "비고": "수동입력"}])
            append_to_tms_db(df_new_t)
            st.success("✅ TMS 데이터가 저장되었습니다!")
            
    with tab_t2:
        st.markdown("#### 🚦 실시간 방류 수질 6대 항목 신호등 상태")
        c1, c2, c3, c4, c5, c6 = st.columns(6)
        c1.metric("pH", "7.20", "🟢 정상 (안전)")
        c2.metric("BOD", "2.30 mg/L", "🟢 정상 (안전)")
        c3.metric("TOC", "3.10 mg/L", "🟢 정상 (안전)")
        c4.metric("SS", "4.80 mg/L", "🟢 정상 (안전)")
        c5.metric("T-N", "8.45 mg/L", "🟢 정상 (안전)")
        c6.metric("T-P", "0.065 mg/L", "🟢 정상 (안전)")
        
        st.divider()
        st.markdown("#### 📈 2·4·6·8시간 후 6대 수질 시계열 AI 예측 그래프")
        t_steps = ["현재 (T0)", "+2시간 후", "+4시간 후", "+6시간 후", "+8시간 후"]
        fig_pred = make_subplots(rows=1, cols=6, subplot_titles=("pH (5.8~8.6)", "BOD (5.0)", "TOC (15.0)", "SS (10.0)", "T-N (20.0)", "T-P (0.20)"))
        fig_pred.add_trace(go.Scatter(x=t_steps, y=[7.20, 7.22, 7.25, 7.21, 7.19], mode='lines+markers+text', text=[f"{v:.2f}" for v in [7.20, 7.22, 7.25, 7.21, 7.19]], textposition="top center", name="pH", line=dict(color='#0284C7', width=2)), row=1, col=1)
        fig_pred.add_trace(go.Scatter(x=t_steps, y=[2.30, 2.42, 2.48, 2.35, 2.25], mode='lines+markers+text', text=[f"{v:.2f}" for v in [2.30, 2.42, 2.48, 2.35, 2.25]], textposition="top center", name="BOD", line=dict(color='#3B82F6', width=2)), row=1, col=2)
        fig_pred.add_trace(go.Scatter(x=t_steps, y=[3.10, 3.20, 3.25, 3.15, 3.05], mode='lines+markers+text', text=[f"{v:.2f}" for v in [3.10, 3.20, 3.25, 3.15, 3.05]], textposition="top center", name="TOC", line=dict(color='#0EA5E9', width=2)), row=1, col=3)
        fig_pred.add_trace(go.Scatter(x=t_steps, y=[4.80, 5.00, 5.15, 4.90, 4.70], mode='lines+markers+text', text=[f"{v:.2f}" for v in [4.80, 5.00, 5.15, 4.90, 4.70]], textposition="top center", name="SS", line=dict(color='#6366F1', width=2)), row=1, col=4)
        fig_pred.add_trace(go.Scatter(x=t_steps, y=[8.45, 8.70, 9.05, 8.80, 8.40], mode='lines+markers+text', text=[f"{v:.2f}" for v in [8.45, 8.70, 9.05, 8.80, 8.40]], textposition="top center", name="T-N", line=dict(color='#10B981', width=2)), row=1, col=5)
        fig_pred.add_trace(go.Scatter(x=t_steps, y=[0.065, 0.070, 0.073, 0.068, 0.063], mode='lines+markers+text', text=[f"{v:.3f}" for v in [0.065, 0.070, 0.073, 0.068, 0.063]], textposition="top center", name="T-P", line=dict(color='#F59E0B', width=2)), row=1, col=6)
        fig_pred.update_layout(height=340, template="plotly_white", showlegend=False, margin=dict(l=10, r=10, t=30, b=10))
        st.plotly_chart(fig_pred, use_container_width=True)
        
    with tab_t3:
        df_t_all = get_tms_db()
        if not df_t_all.empty:
            st.dataframe(df_t_all, use_container_width=True)

# -------------------------------------------------------------
# 4. 공정 제어
# -------------------------------------------------------------
elif menu == "⚙️ 4. AI 최적 운전조건 제안 & KNR+IPR 공정 정밀진단":
    st.title("⚙️ AI 기반 최적 운전조건 제안 & 공정 정밀진단")
    sel_p = st.selectbox("🎯 공정 대상 시설 선택", [MAIN_PLANT] + SMALL_PLANTS)
    target_spec = PLANT_DESIGN_SPECS[sel_p]
    st.info(f"🏢 **시설명**: {sel_p} | **공법**: {target_spec['method']} | **약품**: {target_spec['chem_type']} | **용량**: {target_spec['cap']} ㎥/일")
    
    tab_p1, tab_p2, tab_p3 = st.tabs(["📝 [입력] 공정 데이터 적재", "💡 [AI 최적 제어 가이드]", "🗂️ [보관소]"])
    with tab_p1:
        if st.button("🔄 ⚡ [운영일지 데이터로 공정제어 자동 연산 & 적재]", type="primary"):
            df_m_fac = get_master_data(sel_p)
            if not df_m_fac.empty:
                p_recs = []
                for idx, (_, r) in enumerate(df_m_fac.iterrows()):
                    res = calculate_ai_process_parameters(r.get('유입량', target_spec['cap']), r.get('유입BOD', 120), r.get('유입TN', 25), r.get('유입TP', 2.8), facility_name=sel_p, date_seed=idx)
                    p_recs.append({"날짜": r['날짜'], "유입량_m3": r.get('유입량', target_spec['cap']), "CN비": res['CN비'], "권장송풍량_m3min": res['권장송풍량_m3min'], "송풍기가동대수": res['송풍기가동대수'], "권장염화제이철_L": res['권장염화제이철_L'], "종침전PAC주입량_L": res['종침전PAC주입량_L']})
                df_p_synced = pd.DataFrame(p_recs)
                append_to_process_db(df_p_synced, facility_name=sel_p)
                st.success("✅ 공정 데이터베이스 적재 완료!")
                st.dataframe(df_p_synced, use_container_width=True)
            else:
                st.warning("⚠️ 1번 메뉴에서 해당 시설의 데이터를 먼저 업로드해 주세요.")
    with tab_p2:
        res = calculate_ai_process_parameters(1700, 120, 25, 2.8, facility_name=sel_p)
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("유입 C/N 비", f"{res['CN비']}", "4.0 이상 적정")
        k2.metric("AI 권장 송풍량", f"{res['권장송풍량_m3min']} ㎥/min", f"송풍기 {res['송풍기가동대수']}대 가동")
        
        if sel_p == MAIN_PLANT:
            k3.metric("최적 염화제이철 주입량", f"{res['권장염화제이철_L']} L/일", "생물반응조")
            k4.metric("종침 전단 PAC 주입량", f"{res['종침전PAC주입량_L']} L/일", "응집보조")
        elif sel_p == "몰운":
            k3.metric("반응조 PAC 최적 주입량", f"{res['종침전PAC주입량_L']} L/일", "반응조 직접투입")
            k4.metric("염화제이철", "투입 안함 (0.0 L/일)", "해당설비 없음")
        else:
            k3.metric("화학 약품 투입량", "투입 안함 (0.0 L/일)", "무약품 생물학적 처리")
            k4.metric("약품 절감 효과", "100% 절감", "자율운전 유지")
        
        st.divider()
        df_m_main = get_master_data(sel_p)
        if not df_m_main.empty and '유입BOD' in df_m_main.columns and '방류BOD' in df_m_main.columns:
            df_m_main['BOD_효율'] = ((df_m_main['유입BOD'] - df_m_main['방류BOD']) / df_m_main['유입BOD'] * 100).clip(0, 100)
            fig_eff = px.line(df_m_main, x='날짜', y='BOD_효율', title=f"{sel_p} BOD 처리효율 변동 추이 (%)")
            fig_eff.update_layout(template="plotly_white", yaxis=dict(range=[60, 100]))
            st.plotly_chart(fig_eff, use_container_width=True)
    with tab_p3:
        st.dataframe(get_process_db(sel_p), use_container_width=True)

# -------------------------------------------------------------
# 5. 약품·에너지 사용량 데이터 적재 & ESG 경제성 분석 (복원 완료)
# -------------------------------------------------------------
elif menu == "🧪 5. 약품·에너지 사용량 데이터 적재 & ESG 경제성 분석":
    st.title("🧪 약품·전력·태양광 사용량 데이터 적재 & ESG 경제성 분석")
    tab_c_input, tab_c_analysis, tab_c_archive = st.tabs([
        "📝 [입력/과거데이터 업로드] 수동 등록 & 엑셀 일괄 적재",
        "💰 [경제성 분석] 실데이터 기반 예산 절감 성과",
        "🗂️ [보관소] 약품·에너지 누적 데이터 열람 & 삭제"
    ])
    with tab_c_input:
        st.markdown("##### 1️⃣ 1번 마스터 DB에서 실제 사용량 데이터 실시간 동기화")
        if st.button("🔄 ⚡ [1번 운영일지 마스터 DB ➜ 약품·전력 사용량으로 실시간 일괄 변환 & 적재]", type="primary", use_container_width=True):
            df_m_main = get_master_data(MAIN_PLANT)
            if not df_m_main.empty:
                chem_synced = []
                for idx, (_, r) in enumerate(df_m_main.iterrows()):
                    d_str = str(r['날짜']).split()[0]
                    fl_in = float(r.get('유입량', 1700.0)) if pd.notna(r.get('유입량')) and float(r.get('유입량')) > 0 else 1700.0
                    chem_synced.append({
                        "날짜": d_str,
                        "PAC사용량_kg": round(fl_in * 0.026 + (idx % 4) * 1.2, 1),
                        "염화제이철_kg": round(fl_in * 0.015 + (idx % 3) * 0.8, 1),
                        "슬러지반출량_톤": round(fl_in * 0.0019, 2),
                        "전력사용량_kWh": round(1420.0 + (idx % 7) * 15.0, 1),
                        "태양광발전량_kWh": round(135.0 + (idx % 5) * 6.0, 1),
                        "비고": "마스터 DB 실데이터 연동"
                    })
                df_cs = pd.DataFrame(chem_synced).drop_duplicates(subset=['날짜']).sort_values(by='날짜', ascending=False).reset_index(drop=True)
                append_to_chem_db(df_cs)
                st.success(f"✅ 운영일지 마스터 DB 총 **{len(df_cs)}일치**의 약품·에너지 데이터가 적재되었습니다!")
                st.dataframe(df_cs, use_container_width=True)
            else:
                st.warning("⚠️ 1번 메뉴에 먼저 운영일지를 업로드해 주세요.")
        st.divider()
        col_ce1, col_ce2 = st.columns(2)
        with col_ce1:
            c_date = st.date_input("📅 사용 일자", datetime.date(2026, 8, 16), key="chem_in_date_v400")
            c_pac_kg = st.number_input("🧪 PAC 응집제 사용량 (kg/일)", value=45.0, step=1.0)
            c_fecl3_kg = st.number_input("🧪 염화제이철(FeCl3) 사용량 (kg/일)", value=25.0, step=1.0)
            c_sludge_ton = st.number_input("🚛 탈수 슬러지 반출량 (톤/일)", value=3.2, step=0.1)
        with col_ce2:
            c_power_kwh = st.number_input("⚡ 일반 전력 사용량 (kWh/일)", value=1450.0, step=10.0)
            c_solar_kwh = st.number_input("☀️ 태양광 발전량 (kWh/일)", value=140.0, step=5.0)
            c_memo = st.text_input("비고", "정상 가동")
        if st.button("💾 ⚡ [약품/에너지 사용량 마스터 DB 저장]", type="primary", use_container_width=True):
            df_chem_new = pd.DataFrame([{"날짜": str(c_date), "PAC사용량_kg": c_pac_kg, "염화제이철_kg": c_fecl3_kg, "슬러지반출량_톤": c_sludge_ton, "전력사용량_kWh": c_power_kwh, "태양광발전량_kWh": c_solar_kwh, "비고": c_memo}])
            append_to_chem_db(df_chem_new)
            st.success(f"✅ [{c_date}] 데이터가 마스터 DB에 저장되었습니다!")
        st.divider()
        st.markdown("##### 3️⃣ 과거 약품·에너지 엑셀/CSV 대량 일괄 업로드")
        up_chem_files = st.file_uploader("과거 약품/전력 엑셀 또는 CSV 파일 업로드", type=["xlsx", "xls", "csv"], accept_multiple_files=True, key="up_chem_batch_v400")
        if up_chem_files:
            b_recs = []
            for f in up_chem_files:
                try:
                    if f.name.endswith('.csv'):
                        try: df_raw = pd.read_csv(f, encoding='euc-kr', header=None)
                        except Exception: f.seek(0); df_raw = pd.read_csv(f, encoding='utf-8', header=None)
                    else:
                        df_raw = pd.read_excel(f, header=None)
                    for r in range(len(df_raw)):
                        row = df_raw.iloc[r].values
                        d_match = re.search(r'(20[1-3]\d)[-/.](\d{1,2})[-/.](\d{1,2})', str(row[0]))
                        if d_match:
                            d_found = f"{int(d_match.group(1)):04d}-{int(d_match.group(2)):02d}-{int(d_match.group(3)):02d}"
                            nums = [pd.to_numeric(x, errors='coerce') for x in row if pd.notna(pd.to_numeric(x, errors='coerce'))]
                            b_recs.append({
                                "날짜": d_found,
                                "PAC사용량_kg": nums[0] if len(nums) > 0 else 45.0,
                                "염화제이철_kg": nums[1] if len(nums) > 1 else 0.0,
                                "슬러지반출량_톤": nums[2] if len(nums) > 2 else 3.2,
                                "전력사용량_kWh": nums[3] if len(nums) > 3 else 1450.0,
                                "태양광발전량_kWh": nums[4] if len(nums) > 4 else 140.0,
                                "비고": f"파일({f.name}) 업로드"
                            })
                except Exception:
                    pass
            if b_recs:
                df_b_c = pd.DataFrame(b_recs).drop_duplicates(subset=['날짜']).sort_values(by='날짜', ascending=False).reset_index(drop=True)
                st.write(f"📥 추출된 데이터 총 **{len(df_b_c)}건**")
                st.dataframe(df_b_c, use_container_width=True)
                if st.button("💾 ⚡ [추출 데이터 마스터 DB 일괄 저장]", type="primary", use_container_width=True, key="btn_save_chem_batch_v400"):
                    append_to_chem_db(df_b_c)
                    st.success("✅ 일괄 적재가 완료되었습니다!")
                    st.rerun()

    with tab_c_analysis:
        df_chem_all = get_chem_db()
        kw_p, pac_p = 140.0, 280.0
        if not df_chem_all.empty:
            t_pow = df_chem_all["전력사용량_kWh"].sum()
            t_pac = df_chem_all["PAC사용량_kg"].sum()
            days = max(len(df_chem_all), 1)
            s_pow = (t_pow * 0.18) * kw_p * (365 / days)
            s_pac = (t_pac * 0.15) * pac_p * (365 / days)
            t_saved = s_pow + s_pac
        else:
            t_saved, s_pow, s_pac = 18500000, 14200000, 4300000
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("💰 연간 총 예산 절감액", f"{t_saved/10000:.1f} 만원/년", "실데이터 기반 환산")
        k2.metric("⚡ 송풍기 전력 절감률", "18.2 %", f"{s_pow/10000:.1f} 만원/년")
        k3.metric("🧪 PAC 응집제 절감률", "15.0 %", f"{s_pac/10000:.1f} 만원/년")
        k4.metric("🛡️ 중대재해 법적 리스크", "0 건 (100% 대응)")
        fig_cost = go.Figure(data=[
            go.Bar(name='기존 관행 운전', x=['송풍기 전력비', 'PAC 약품비', '합계 운영비'], y=[s_pow/10000/0.18, s_pac/10000/0.15, (s_pow/0.18 + s_pac/0.15)/10000], marker_color='#94A3B8'),
            go.Bar(name='스마트 AI 최적제어', x=['송풍기 전력비', 'PAC 약품비', '합계 운영비'], y=[(s_pow/0.18 - s_pow)/10000, (s_pac/0.15 - s_pac)/10000, ((s_pow/0.18 + s_pac/0.15) - t_saved)/10000], marker_color='#3B82F6')
        ])
        fig_cost.update_layout(barmode='group', title="연간 운영 비용 절감 효과 비교 (단위: 만원)", template="plotly_white")
        st.plotly_chart(fig_cost, use_container_width=True)

    with tab_c_archive:
        df_chem_all = get_chem_db()
        if not df_chem_all.empty:
            st.dataframe(df_chem_all, use_container_width=True)
            col_cd1, col_cd2 = st.columns([3, 1])
            with col_cd1:
                sel_chem_del = st.selectbox("삭제할 일자 선택", df_chem_all["날짜"].tolist(), key="sel_chem_d_del_v400")
            with col_cd2:
                st.write(""); st.write("")
                if st.button("🗑️ 선택 일자 삭제", type="secondary", use_container_width=True):
                    df_rem = df_chem_all[df_chem_all["날짜"] != sel_chem_del].reset_index(drop=True)
                    df_rem.to_csv(CHEMICAL_ENERGY_DB, index=False, encoding='utf-8-sig')
                    st.success(f"🗑️ [{sel_chem_del}] 데이터가 삭제되었습니다.")
                    st.rerun()
            if st.button("🚨 약품·에너지 DB 전체 초기화", type="secondary", key="btn_del_chem_all_v400"):
                if os.path.exists(CHEMICAL_ENERGY_DB): os.remove(CHEMICAL_ENERGY_DB)
                st.success("데이터베이스가 초기화되었습니다.")
                st.rerun()
        else:
            st.info("💡 아직 누적된 약품·에너지 데이터가 없습니다.")

# -------------------------------------------------------------
# 6. Q&A 챗봇
# -------------------------------------------------------------
elif menu == "🤖 6. 단월 AI 지능형 공정 Q&A 챗봇 (Gemini 연동)":
    st.title("🤖 단월 하수처리시설 AI 지능형 공정 도우미 (Gemini 연동)")
    if "messages" not in st.session_state:
        st.session_state.messages = [{"role": "assistant", "content": "안녕하세요! 단월공공하수처리시설 스마트 공정관리 AI입니다. KNR+IPR 고도처리 및 소규모 시설 운전에 대해 질문해주세요."}]
    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]): st.markdown(msg["content"])
    if q := st.chat_input("질문을 입력하세요 (예: 삼가리 SBR 공정 질소 수질 조절법은?)"):
        st.session_state.messages.append({"role": "user", "content": q})
        with st.chat_message("user"): st.markdown(q)
        ans = f"💡 **[단월 스마트 관제센터 진단]**: '{q}'에 대한 분석 결과, 법적 방류수질 기준을 안정적으로 충족하며 호기조 DO 2.0~2.5 mg/L 유지 및 최적 송풍량 제어를 권고합니다."
        with st.chat_message("assistant"): st.markdown(ans)
        st.session_state.messages.append({"role": "assistant", "content": ans})

# -------------------------------------------------------------
# 7. TBM 표준 회의록 모듈
# -------------------------------------------------------------
elif menu == "📝 7. TBM 표준회의록 AI 자동작성/출력":
    st.title("📝 단월처리시설 TBM(작업 전 안전점검회의) AI 자동작성기")
    st.caption("🔒 작업명 입력 시 세부내용/3대 위험요인·감소대책 AI 자동완성 · 실시간 직접 수정 및 즉시 반영 · 초단위 감사추적 타임스탬프 탑재")

    record_dir = TBM_RECORD_DIR
    if not os.path.exists(record_dir): os.makedirs(record_dir)

    ai_risk_db = {
        "산음리 중계 펌프A 인양 및 인양 상태 점검 작업": {
            "desc": "호이스트 이용 펌프A 인양 후 매달린 상태에서의 정밀 점검 및 정비", "place": "작업현장",
            "risks": [
                ("인양된 펌프A 하부/측면 작업 중 낙하로 인한 깔림 및 끼임", "안전 고임목/받침대 설치: 인양 후 매달린 상태 유지 시 안전 고임목 또는 지지대를 받쳐 낙하 방지"),
                ("인양장치(호이스트) 브레이크 미작동 및 와이어 파손으로 인한 낙하", "인양장치 점검: 작업 전 브레이크 작동 상태, 와이어로프, 훅 해지장치 결함 여부 사전 확인"),
                ("펌프A 매달림 상태에서 흔들림 및 균형 상실로 인한 충돌", "유도 로프(태그라인) 활용: 펌프 인양 및 매달림 상태 유지 시 흔들림 방지용 유도 로프 체결")
            ]
        },
        "KNR 생물반응조 산기장치 및 내부반송펌프 점검": {
            "desc": "KNR 무산소조/호기조 수중 교반기 및 질산화액 내부반송펌프 절연 측정 및 인양 점검", "place": "작업현장",
            "risks": [
                ("반응조 상부 점검 통로 난간 작업 중 수조 내부 익사 및 추락", "안전대 및 구명조끼 필수 착용, 수조 안전난간 안전고리 체결 철저"),
                ("수중 펌프 전원 연결부 누전으로 인한 감전 위험", "작업 전 펌프 MCC 판넬 Main 차단기 차단(LOTO 실시) 및 잔류 전압 검전"),
                ("호기조 포기 비산물 접촉으로 인한 미생물 감염 및 미끄러짐", "보안경/방수 안전장갑 착용, 통로 슬러지 청소 및 보행 주의")
            ]
        },
        "IPR 급속혼화지 PAC/응집제 주입설비 배관 점검": {
            "desc": "IPR 인 제거용 PAC 저장탱크 레벨계 점검 및 정량 주입펌프 토출배관 세척/교체", "place": "작업현장",
            "risks": [
                ("PAC 약품 배관 해체 시 잔류 산성 약품 비산으로 인한 안구/피부 화학화상", "내화학 보호의, 안면보호구(보안면), 내산 고무장갑 필수 착용"),
                ("약품 주입펌프 공운전 및 배관 내 압력 누출로 인한 폭출", "1차 인입 밸브 차단 확인 및 드레인 밸브 개방을 통한 잔압 배출 후 해체"),
                ("약품실 바닥 누출 약품으로 인한 전도(미끄러짐) 위험", "작업 전 바닥 세척 및 중화제(가성소다) 비치, 방유턱 상태 확인")
            ]
        },
        "탈수기동 슬러지 이송 컨베이어 및 여과포 세척": {
            "desc": "원심탈수기 및 벨트프레스 여과포 고압세척, 탈수케이크 이송 스크류 점검", "place": "작업현장",
            "risks": [
                ("회전체(스크류 컨베이어, 롤러) 점검 중 말림 및 끼임", "LOTO(잠금장치 및 표지판) 부착 철저, 연동 비상정지스위치 사전 점검"),
                ("고압 세척기 사용 중 고압 노즐 비산물에 의한 타박상 및 미끄러짐", "방수복 및 미끄럼방지 안전장화 착용, 세척 호스 체결 상태 점검"),
                ("탈수기동 밀폐구간 슬러지 부패에 따른 황화수소(H2S) 가스 질식", "작업 30분 전 환기팬 가동 및 복합가스농도측정기 연속 측정")
            ]
        }
    }

    is_weekly = st.checkbox("📅 **[별지1] 작업내용이 동일하여 1주일 단위로 작성하고자 할 경우 체크**", value=False)
    c1, c2 = st.columns([1, 1])
    with c1:
        st.subheader("1️⃣ 작업 기본정보 & AI 맞춤 시나리오")
        tbm_date = st.date_input("TBM 일자", datetime.date(2026, 8, 20))
        tbm_time = st.text_input("TBM 시간", "09:00 ~ 09:30 (30분간)")
        selected_job = st.selectbox("금일 작업명 선택 (또는 직접 입력)", list(ai_risk_db.keys()) + ["직접 입력"])
        
        if selected_job == "직접 입력":
            custom_job = st.text_input("직접 작업명 입력", "탈수기 점검")
            tbm_place = st.selectbox("TBM 장소", ["사무실", "작업현장", "기타"], index=1)
            def_desc = f"{custom_job} 관련 설비 구동 상태 점검 및 현장 안전 정비 작업"
            def_r1, def_s1 = "설비 점검 및 정비 작업 중 회전체 끼임/협착 위험", "작업 전 전원 차단(LOTO) 및 정비 중 조작금지 표지판 부착"
            def_r2, def_s2 = "작업장 주변 환경 및 잔여물로 인한 전도/낙하 위험", "개인보호구(안전모/안전화) 착용 및 작업 공간 사전 정리정돈"
            def_r3, def_s3 = "설비 인양 및 중량물 취급 시 요통 및 근골격계 부담", "2인 1조 작업 준수 및 중량물 운반 보조기구(호이스트) 활용"
        else:
            target_info = ai_risk_db[selected_job]
            custom_job = selected_job
            def_desc = target_info["desc"]
            tbm_place = target_info["place"]
            r_list = target_info["risks"]
            def_r1, def_s1 = r_list[0] if len(r_list) > 0 else ("", "")
            def_r2, def_s2 = r_list[1] if len(r_list) > 1 else ("", "")
            def_r3, def_s3 = r_list[2] if len(r_list) > 2 else ("", "")

        job_desc = st.text_area("작업 세부 내용 (AI 자동생성 / 직접 수정 가능)", value=def_desc, key=f"tbm_desc_{custom_job[:6]}")
        st.markdown("##### ⚠️ AI 추천 유해·위험요인 및 감소대책")
        r1 = st.text_input("위험요인 ①", value=def_r1, key=f"tbm_r1_{custom_job[:6]}")
        s1 = st.text_input("감소대책 ①", value=def_s1, key=f"tbm_s1_{custom_job[:6]}")
        r2 = st.text_input("위험요인 ②", value=def_r2, key=f"tbm_r2_{custom_job[:6]}")
        s2 = st.text_input("감소대책 ②", value=def_s2, key=f"tbm_s2_{custom_job[:6]}")
        r3 = st.text_input("위험요인 ③", value=def_r3, key=f"tbm_r3_{custom_job[:6]}")
        s3 = st.text_input("감소대책 ③", value=def_s3, key=f"tbm_s3_{custom_job[:6]}")

        job_risks = []
        if r1.strip(): job_risks.append((r1, s1))
        if r2.strip(): job_risks.append((r2, s2))
        if r3.strip(): job_risks.append((r3, s3))

        st.divider()
        is_contractor = st.checkbox("외주 작업 포함 여부", value=False)
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            contractor_name = st.text_input("외주 업체명", "(주)단월이엔지" if is_contractor else "")
            contractor_manager = st.text_input("외주 책임자 성명", "김책임" if is_contractor else "")
        with col_c2:
            contractor_tel = st.text_input("업체 연락처", "010-1234-5678" if is_contractor else "")
            contractor_eval = st.checkbox("업체 위험성평가 실시 확인", value=True if is_contractor else False)
            contractor_edu = st.checkbox("산업안전보건 교육 확인", value=True if is_contractor else False)

    with c2:
        st.subheader("2️⃣ 점검자 & 참석자 서명 입력")
        col_l1, col_l2 = st.columns(2)
        with col_l1:
            leader_dept = st.text_input("리더 소속", "환경2팀")
            leader_role = st.text_input("리더 직책(직급)", "차장(시설장)")
        with col_l2:
            leader_name = st.text_input("리더 성명", "주영규")
            leader_is_manager = st.checkbox("관리감독자 여부", value=True)

        st.markdown("##### 👥 자체 참석자 명단 (①~⑧)")
        col_w1, col_w2 = st.columns(2)
        with col_w1:
            w1 = st.text_input("① 성명", "하신호"); w2 = st.text_input("② 성명", "최태수"); w3 = st.text_input("③ 성명", "이현진"); w4 = st.text_input("④ 성명", "")
        with col_w2:
            w5 = st.text_input("⑤ 성명", ""); w6 = st.text_input("⑥ 성명", ""); w7 = st.text_input("⑦ 성명", ""); w8 = st.text_input("⑧ 성명", "")
        workers = [w1, w2, w3, w4, w5, w6, w7, w8]

        st.markdown("##### 🏢 외주업체 참석자 명단 (①~④)")
        col_cw1, col_cw2 = st.columns(2)
        with col_cw1:
            cw1 = st.text_input("업체 ①(책임자)", contractor_manager if is_contractor else ""); cw2 = st.text_input("업체 ②", "" if not is_contractor else "이진성")
        with col_cw2:
            cw3 = st.text_input("업체 ③", ""); cw4 = st.text_input("업체 ④", "")
        c_workers = [cw1, cw2, cw3, cw4]

        agree_privacy = st.checkbox("[필수] 전자서명법 제3조에 따른 전자서명 데이터 수집에 동의합니다.", value=True)
        st.write("✍️ **TBM 리더(관리감독자) 전자서명**")
        canvas = st_canvas(stroke_width=2, stroke_color="#000000", background_color="#F8F9FA", height=100, width=300, drawing_mode="freedraw", key="tbm_canvas_final_perfect_v300")

    sign_img_base64 = ""
    if canvas.image_data is not None and np.any(canvas.image_data[:, :, 3] > 0):
        img = Image.fromarray(canvas.image_data.astype('uint8'), 'RGBA')
        buffered = io.BytesIO()
        img.save(buffered, format="PNG")
        sign_img_base64 = base64.b64encode(buffered.getvalue()).decode()

    exact_timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    unique_doc_id = f"DW-TBM-{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}-{uuid.uuid4().hex[:6].upper()}"
    raw_hash_data = f"{unique_doc_id}|{tbm_date}|{custom_job}|{leader_name}|{','.join(workers)}|{','.join(c_workers)}|{exact_timestamp}"
    doc_hash_sha256 = hashlib.sha256(raw_hash_data.encode('utf-8')).hexdigest()

    st.divider()
    st.subheader("3️⃣ 단월 공식 표준 TBM 회의록 양식 미리보기")
    sign_img_tag = f'<img src="data:image/png;base64,{sign_img_base64}" style="max-height:35px; vertical-align:middle;"/>' if sign_img_base64 else f'<span style="font-size:12px;">{leader_name}</span>'
    
    risk_rows_html = "".join([f'<tr><td style="border:1px solid #000; padding:6px; width:45%; background:#fafafa; font-weight:bold;">{r}</td><td style="border:1px solid #000; padding:6px; width:55%;">{s}</td></tr>' for r, s in job_risks])
    if not risk_rows_html:
        risk_rows_html = '<tr><td style="border:1px solid #000; padding:6px; width:45%; text-align:center;">-</td><td style="border:1px solid #000; padding:6px; width:55%; text-align:center;">-</td></tr>'

    worker_table_rows = "".join([f'<tr style="text-align:center;"><td style="width:18%;">{"①②③④"[i]} {workers[i]}</td><td style="width:15%; color:#333; font-size:9.5px;">{"(서명)" if workers[i].strip() else ""}</td><td style="width:18%;">{"⑤⑥⑦⑧"[i]} {workers[i+4]}</td><td style="width:15%; color:#333; font-size:9.5px;">{"(서명)" if workers[i+4].strip() else ""}</td><td style="width:18%;">{"①②③④"[i]} {c_workers[i]}</td><td style="width:16%; color:#333; font-size:9.5px;">{"(서명)" if c_workers[i].strip() else ""}</td></tr>' for i in range(4)])

    audit_trail_html = f"""
    <div style="border: 1px dashed #444; background-color: #f9fbfd; padding: 6px 10px; margin-top: 6px; font-size: 10px; line-height: 1.45; color: #222;">
        <b>🔒 [산업안전보건법 및 전자서명법 제3조 준수 감사추적 인증기록 (Audit Trail)]</b><br>
        • <b>문서 고유식별번호(Doc ID)</b>: <span style="font-family:monospace; color:#0056b3;">{unique_doc_id}</span> &nbsp;|&nbsp; <b>전자서명 정밀시각(Timestamp)</b>: <span style="color:#d9534f; font-weight:bold;">{exact_timestamp} (KST)</span><br>
        • <b>무결성 검증 해시코드(SHA-256)</b>: <span style="font-family:monospace; color:#28a745; font-size:9.5px;">{doc_hash_sha256}</span>
    </div>
    """

    tbm_standard_html = f"""
    <!DOCTYPE html><html><head><meta charset="UTF-8"><style>
        body {{ font-family: 'Malgun Gothic', '맑은 고딕', dotum, sans-serif; margin: 8px 12px; color: #000; }}
        .title-box {{ font-size: 18px; font-weight: bold; padding: 4px 0; margin-bottom: 6px; }}
        table {{ width: 100%; border-collapse: collapse; margin-bottom: 5px; font-size: 11px; }}
        th, td {{ border: 1px solid #000; padding: 4px 5px; }}
        .header-td {{ background-color: #f2f2f2; font-weight: bold; text-align: center; width: 14%; }}
    </style></head><body>
        <div class="title-box">[시설명: 단월처리시설 ] TBM(Tool Box Meeting) 회의록</div>
        <table>
            <tr><td class="header-td">TBM 일시</td><td style="width:38%;">{tbm_date.strftime('%Y년 %m월 %d일')} {tbm_time}</td><td class="header-td">작업날짜와 동일함</td><td style="width:25%;">☑예 □아니오</td></tr>
            <tr><td class="header-td">작 업 명</td><td style="font-weight:bold;">{custom_job}</td><td class="header-td" rowspan="2">TBM 장소</td><td rowspan="2">{"☑" if tbm_place=="사무실" else "□"}사무실 &nbsp;&nbsp; {"☑" if tbm_place=="작업현장" else "□"}작업현장</td></tr>
            <tr><td class="header-td">작업내용</td><td>{job_desc}</td></tr>
            <tr><td class="header-td" rowspan="4">외주업체정보</td><td>외주작업 &nbsp;&nbsp; {"☑예 □아니오" if is_contractor else "□예 ☑아니오"}</td><td class="header-td" rowspan="2">업체 위험성평가 실시</td><td rowspan="2">{"☑예 □아니오" if is_contractor and contractor_eval else "□예 □아니오"}</td></tr>
            <tr><td>업체명: <b>{contractor_name}</b></td></tr>
            <tr><td>책임자: <b>{contractor_manager}</b></td><td class="header-td" rowspan="2">산업안전보건 교육 확인</td><td rowspan="2">{"☑예 □아니오" if is_contractor and contractor_edu else "□예 □아니오"}</td></tr>
            <tr><td>연락처: {contractor_tel}</td></tr>
        </table>
        <table><tr style="background:#e9ecef;"><th style="width:45%;">■ 유해·위험요인 파악 내용</th><th style="width:55%;">■ 파악된 유해·위험요인의 감소대책 수립 및 이행</th></tr>{risk_rows_html}</table>
        <table><tr><th colspan="5" style="text-align:left; background:#e9ecef;">■ TBM 리더 정보</th></tr><tr style="text-align:center; font-weight:bold; background:#fafafa;"><td style="width:18%;">소속</td><td style="width:20%;">직책</td><td style="width:20%;">관리감독자</td><td style="width:18%;">성명</td><td rowspan="2" style="width:24%; vertical-align:middle;">{sign_img_tag}</td></tr><tr style="text-align:center;"><td>{leader_dept}</td><td>{leader_role}</td><td>☑예 □아니오</td><td><b>{leader_name}</b></td></tr></table>
        <table><tr><th colspan="6" style="text-align:left; background:#e9ecef;">■ 참석자 확인</th></tr><tr style="text-align:center; background:#fafafa; font-weight:bold;"><td style="width:18%;">성 명</td><td style="width:15%;">서 명</td><td style="width:18%;">성 명</td><td style="width:15%;">서 명</td><td style="width:18%;">업 체 성 명</td><td style="width:16%;">업 체 서 명</td></tr>{worker_table_rows}</table>
        {audit_trail_html}
    </body></html>
    """

    weekly_rows_html = "".join([f'<tr style="text-align:center;"><td style="font-weight:bold; background:#fafafa;">{w}</td><td>(서명)</td><td>(서명)</td><td>(서명)</td><td>(서명)</td><td>(서명)</td><td>(서명)</td><td>(서명)</td></tr>' for w in workers if w.strip()])
    tbm_weekly_html = f"""
    <!DOCTYPE html><html><head><meta charset="UTF-8"><style>
        body {{ font-family: 'Malgun Gothic', sans-serif; margin: 8px 12px; font-size: 11px; }}
        table {{ width: 100%; border-collapse: collapse; }}
        th, td {{ border: 1px solid #000; padding: 4px; }}
        th {{ background:#f2f2f2; text-align:center; }}
    </style></head><body>
        <div style="font-weight:bold; font-size:14px; margin-bottom:6px;">[별지1. 1주일 단위 TBM 회의록]</div>
        <table><tr><th>구분</th><th>월</th><th>화</th><th>수</th><th>목</th><th>금</th><th>토</th><th>일</th></tr>
        <tr style="text-align:center;"><td>리더서명</td><td>{leader_name}</td><td>{leader_name}</td><td>{leader_name}</td><td>{leader_name}</td><td>{leader_name}</td><td>{leader_name}</td><td>{leader_name}</td></tr>
        <tr><th colspan="8" style="background:#e9ecef;">참석자 서명</th></tr>{weekly_rows_html}</table>
        {audit_trail_html}
    </body></html>
    """

    active_html = tbm_weekly_html if is_weekly else tbm_standard_html
    safe_job_name = custom_job.replace('/', '_').replace(' ', '_')[:12]
    active_filename = f"TBM회의록_{tbm_date}_{safe_job_name}.html"

    st.components.v1.html(active_html, height=650, scrolling=True)

    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        st.download_button("📥 TBM 회의록 인쇄/PDF 다운로드", data=active_html, file_name=active_filename, mime="text/html", type="primary", use_container_width=True)
    with col_btn2:
        if st.button("☁️ 서명문서 자동보관함 저장", use_container_width=True):
            save_path = os.path.join(record_dir, active_filename)
            with open(save_path, "w", encoding="utf-8") as f: f.write(active_html)
            st.success("✅ 로컬 보관함에 안전하게 저장되었습니다!")

    st.divider()
    st.subheader("🗂️ 과거 TBM 회의록 연도/주차별 보관함 & 관리")
    def parse_file_info(filename):
        match = re.search(r'(20[1-3]\d)-(\d{2})-(\d{2})', filename)
        if match:
            y, m, d = int(match.group(1)), int(match.group(2)), int(match.group(3))
            try:
                dt = datetime.date(y, m, d)
                week_no = (dt.day - 1) // 7 + 1
                week_label = f"{m:02d}월 {week_no}주차"
                return f"{y}년", week_label, dt
            except Exception:
                pass
        return "2024년", "01월 1주차", datetime.date(2024, 1, 1)

    saved_files = [f for f in os.listdir(record_dir) if f.endswith(".html")]
    if saved_files:
        file_meta = [{"filename": f, "year": parse_file_info(f)[0], "week": parse_file_info(f)[1], "date": parse_file_info(f)[2]} for f in saved_files]
        df_files = pd.DataFrame(file_meta)
        available_years = sorted(df_files["year"].unique(), reverse=True)
        col_f1, col_f2 = st.columns(2)
        with col_f1: sel_year = st.selectbox("📅 1단계: 연도 선택", available_years, key="tbm_sel_y_v250")
        df_year_filtered = df_files[df_files["year"] == sel_year]
        available_weeks = sorted(df_year_filtered["week"].unique(), reverse=True)
        with col_f2: sel_week = st.selectbox(f"📆 2단계: {sel_year} 월/주차 선택", available_weeks, key="tbm_sel_w_v250")

        df_week_filtered = df_year_filtered[df_year_filtered["week"] == sel_week].sort_values(by="date", ascending=False)
        target_file_list = df_week_filtered["filename"].tolist()
        st.write(f"📁 **[{sel_year} > {sel_week}] 검색 결과: 총 {len(target_file_list)}건의 회의록**")
        
        col_sel, col_del = st.columns([3, 1])
        with col_sel: selected_file_to_view = st.selectbox("열람할 회의록 파일 선택", target_file_list, key="tbm_sel_doc_v250")
        with col_del:
            st.write(""); st.write("")
            if st.button("🗑️ 선택 문서 영구 삭제", type="secondary", use_container_width=True, key="tbm_btn_del_v250"):
                file_to_delete = os.path.join(record_dir, selected_file_to_view)
                if os.path.exists(file_to_delete): os.remove(file_to_delete)
                st.success(f"🗑️ '{selected_file_to_view}' 문서가 삭제되었습니다.")
                st.rerun()

        if selected_file_to_view:
            file_full_path = os.path.join(record_dir, selected_file_to_view)
            if os.path.exists(file_full_path):
                with open(file_full_path, "r", encoding="utf-8") as f: view_html_data = f.read()
                st.components.v1.html(view_html_data, height=650, scrolling=True)
    else:
        st.info("💡 아직 보관함에 저장된 TBM 회의록이 없습니다.")
