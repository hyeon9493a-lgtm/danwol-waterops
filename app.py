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

# 4. DB 입출력 핸들러 (데이터 영구 보존)
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

# 6. 연도/월 다중 파일 파서 (단월 본장)
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
            
            xl = pd.ExcelFile(f)
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
                    if v in ['1', '1.0', 1]:
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

# [소규모 6개소 24열 서식 및 운영일지 파서]
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
                            
                            in_bod = pd.to_numeric(ws.cell(r, 9).value, errors='coerce')
                            in_toc = pd.to_numeric(ws.cell(r, 10).value, errors='coerce')
                            in_ss = pd.to_numeric(ws.cell(r, 11).value, errors='coerce')
                            in_tn = pd.to_numeric(ws.cell(r, 12).value, errors='coerce')
                            in_tp = pd.to_numeric(ws.cell(r, 13).value, errors='coerce')
                            in_coli = pd.to_numeric(ws.cell(r, 14).value, errors='coerce')
                            
                            out_bod = pd.to_numeric(ws.cell(r, 17).value, errors='coerce')
                            out_toc = pd.to_numeric(ws.cell(r, 18).value, errors='coerce')
                            out_ss = pd.to_numeric(ws.cell(r, 19).value, errors='coerce')
                            out_tn = pd.to_numeric(ws.cell(r, 20).value, errors='coerce')
                            out_tp = pd.to_numeric(ws.cell(r, 21).value, errors='coerce')
                            out_coli = pd.to_numeric(ws.cell(r, 22).value, errors='coerce')
                            
                            if pd.notna(in_bod): rec['유입BOD'] = float(in_bod)
                            if pd.notna(in_toc): rec['유입TOC'] = float(in_toc)
                            if pd.notna(in_ss): rec['유입SS'] = float(in_ss)
                            if pd.notna(in_tn): rec['유입TN'] = float(in_tn)
                            if pd.notna(in_tp): rec['유입TP'] = float(in_tp)
                            if pd.notna(in_coli): rec['유입대장균'] = float(in_coli)
                            
                            if pd.notna(out_bod): rec['방류BOD'] = float(out_bod)
                            if pd.notna(out_toc): rec['방류TOC'] = float(out_toc)
                            if pd.notna(out_ss): rec['방류SS'] = float(out_ss)
                            if pd.notna(out_tn): rec['방류TN'] = float(out_tn)
                            if pd.notna(out_tp): rec['방류TP'] = float(out_tp)
                            if pd.notna(out_coli): rec['방류대장균'] = float(out_coli)
                else:
                    sheet_m = None
                    sm_match = re.search(r'(\d{1,2})월', sname)
                    if sm_match: sheet_m = int(sm_match.group(1))

                    for r in range(1, min(ws.max_row + 1, 600)):
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
                                    dt_val = datetime.date(file_year_anchor, v.month, v.day); break
                                elif isinstance(v, str) and re.search(r'(20[1-3]\d)[-/.](\d{1,2})[-/.](\d{1,2})', v):
                                    m = re.search(r'(20[1-3]\d)[-/.](\d{1,2})[-/.](\d{1,2})', v)
                                    dt_val = datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3))); break
                                elif isinstance(v, str) and re.search(r'(\d{1,2})[-/.](\d{1,2})', v):
                                    m = re.search(r'(\d{1,2})[-/.](\d{1,2})', v)
                                    dt_val = datetime.date(file_year_anchor, int(m.group(1)), int(m.group(2))); break
                                elif isinstance(v, (int, float)) and sheet_m and (1 <= int(v) <= 31):
                                    try:
                                        dt_val = datetime.date(file_year_anchor, sheet_m, int(v)); break
                                    except Exception: pass

                        if cur_fac and dt_val:
                            d_str = dt_val.strftime('%Y-%m-%d')
                            if d_str not in accumulated_data[cur_fac]:
                                accumulated_data[cur_fac][d_str] = {'날짜': d_str}
                            rec = accumulated_data[cur_fac][d_str]
                            
                            nums = [pd.to_numeric(val, errors='coerce') for val in row if pd.notna(pd.to_numeric(val, errors='coerce'))]
                            if len(nums) >= 6:
                                rec.update({
                                    '유입BOD': nums[0], '유입TOC': nums[1], '유입SS': nums[2],
                                    '유입TN': nums[3], '유입TP': nums[4],
                                    '방류BOD': nums[5], '방류TOC': nums[6] if len(nums)>6 else np.nan,
                                    '방류SS': nums[7] if len(nums)>7 else np.nan, '방류TN': nums[8] if len(nums)>8 else np.nan,
                                    '방류TP': nums[9] if len(nums)>9 else np.nan
                                })
                            if len(nums) >= 1 and ('유입량' not in rec or pd.isna(rec['유입량'])):
                                if len(nums) >= 14 and pd.notna(nums[13]):
                                    rec['유입량'] = float(nums[13])
                                    rec['방류량'] = float(nums[13])
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

def fill_exact_main_template(df_data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws['A1'] = "유량및수질관리 업로드양식"
    ws.merge_cells('A1:AY1')
    headers_r1 = {
        'A2': '날짜', 'B2': '유입량(반류수 포함)(㎥/일)', 'C2': '반류수 유량(㎥/일)',
        'D2': '실제 유입량(㎥/일)', 'E2': '처리량', 'H2': '방류량(㎥)/일',
        'I2': '처리시설 유입전 우수토실 방류량(㎥)/일', 'J2': '수온(℃)',
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

def fill_exact_small_template(df_data, fac_name, start_date=None, end_date=None, year=2026):
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

    if start_date and end_date:
        d_range = pd.date_range(start_date, end_date)
    elif df_data is not None and not df_data.empty and '날짜' in df_data.columns:
        min_d = df_data['날짜'].min()
        max_d = df_data['날짜'].max()
        d_range = pd.date_range(min_d, max_d)
    else:
        d_range = pd.date_range(f"{year}-01-01", f"{year}-12-31")

    lookup = {}
    if df_data is not None and not df_data.empty and '날짜' in df_data.columns:
        for _, r in df_data.iterrows():
            d_key = str(r['날짜']).split()[0]
            lookup[d_key] = r

    last_known_flow_in = default_f
    last_known_flow_out = default_f

    for r_idx, dt in enumerate(d_range, start=4):
        d_str = dt.strftime('%Y-%m-%d')
        c1 = ws.cell(r_idx, 1, dt.date())
        c1.number_format = 'yyyy-mm-dd'

        r_match = lookup.get(d_str, None)
        if r_match is not None:
            raw_in = r_match.get('유입량', np.nan)
            raw_out = r_match.get('방류량', np.nan)
            if pd.notna(raw_in) and 0.001 <= float(raw_in) <= 2000:
                last_known_flow_in = float(raw_in)
            if pd.notna(raw_out) and 0.001 <= float(raw_out) <= 2000:
                last_known_flow_out = float(raw_out)

            flow_in = last_known_flow_in
            flow_out = last_known_flow_out

            ws.cell(r_idx, 2, flow_in)
            ws.cell(r_idx, 5, flow_in)
            ws.cell(r_idx, 6, flow_out)

            raw_temp = r_match.get('수온', np.nan)
            if pd.notna(raw_temp): ws.cell(r_idx, 7, float(raw_temp))

            if pd.notna(r_match.get('유입BOD')): ws.cell(r_idx, 9, float(r_match.get('유입BOD')))
            if pd.notna(r_match.get('유입TOC')): ws.cell(r_idx, 10, float(r_match.get('유입TOC')))
            if pd.notna(r_match.get('유입SS')): ws.cell(r_idx, 11, float(r_match.get('유입SS')))
            if pd.notna(r_match.get('유입TN')): ws.cell(r_idx, 12, float(r_match.get('유입TN')))
            if pd.notna(r_match.get('유입TP')): ws.cell(r_idx, 13, float(r_match.get('유입TP')))
            if pd.notna(r_match.get('유입대장균')): ws.cell(r_idx, 14, float(r_match.get('유입대장균')))

            if pd.notna(r_match.get('방류BOD')): ws.cell(r_idx, 17, float(r_match.get('방류BOD')))
            if pd.notna(r_match.get('방류TOC')): ws.cell(r_idx, 18, float(r_match.get('방류TOC')))
            if pd.notna(r_match.get('방류SS')): ws.cell(r_idx, 19, float(r_match.get('방류SS')))
            if pd.notna(r_match.get('방류TN')): ws.cell(r_idx, 20, float(r_match.get('방류TN')))
            if pd.notna(r_match.get('방류TP')): ws.cell(r_idx, 21, float(r_match.get('방류TP')))
            if pd.notna(r_match.get('방류대장균')): ws.cell(r_idx, 22, float(r_match.get('방류대장균')))
        else:
            ws.cell(r_idx, 2, last_known_flow_in)
            ws.cell(r_idx, 5, last_known_flow_in)
            ws.cell(r_idx, 6, last_known_flow_out)

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
        "body { font-family: 'Malgun Gothic', 'Pretendard', sans-serif; margin: 8px 12px; color: #000; font-size: 11px; }",
        ".title-box { font-size: 17px; font-weight: bold; padding: 4px 0; margin-bottom: 6px; border-bottom: 2px solid #000; }",
        "table { width: 100%; border-collapse: collapse; margin-bottom: 5px; font-size: 11px; }",
        "th, td { border: 1px solid #000; padding: 4px 5px; }",
        ".header-td { background-color: #f2f2f2; font-weight: bold; text-align: center; width: 14%; }",
        "</style></head><body>",
        "<div class='title-box'>[시설명: 단월처리시설 ] TBM(Tool Box Meeting) 회의록</div>",
        "<table>",
        f"<tr><td class='header-td'>TBM 일시</td><td style='width:38%;'>{tbm_date.strftime('%Y년 %m월 %d일')} {tbm_time}</td><td class='header-td'>작업날짜와 동일함</td><td style='width:25%;'>☑예 □아니오</td></tr>",
        f"<tr><td class='header-td'>작 업 명</td><td style='font-weight:bold;'>{custom_job}</td><td class='header-td' rowspan='2'>TBM 장소</td><td rowspan='2'>{'☑' if tbm_place=='사무실' else '□'}사무실 &nbsp;&nbsp; {'☑' if tbm_place=='작업현장' else '□'}작업현장</td></tr>",
        f"<tr><td class='header-td'>작업내용</td><td>{job_desc}</td></tr>",
        f"<tr><td class='header-td' rowspan='4'>외주업체정보</td><td>외주작업 &nbsp;&nbsp; {'☑예 □아니오' if is_contractor else '□예 ☑아니오'}</td><td class='header-td' rowspan='2'>업체 위험성평가 실시</td><td rowspan='2'>{'☑예 □아니오' if is_contractor and contractor_eval else '□예 □아니오'}
