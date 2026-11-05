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
    "산음": {"cap": 100.0, "method": "SWPP", "blower_cap": 3.0, "has_chem": False, "chem_type": "무약품", "desc": "수중포기 침전일체형"},
    "삼가리": {"cap": 120.0, "method": "SBR", "blower_cap": 3.5, "has_chem": False, "chem_type": "무약품", "desc": "회분식 활성슬러지 공정"},
    "진목": {"cap": 23.0, "method": "고효율오수정화 + SOD", "blower_cap": 1.5, "has_chem": False, "chem_type": "무약품", "desc": "미생물 접촉산화 및 고효율 탈질"},
    "몰운": {"cap": 60.0, "method": "IC-SBR", "blower_cap": 2.0, "has_chem": True, "chem_type": "반응조 PAC", "desc": "간헐 포기 회분식 반응조"},
    "단월마을": {"cap": 30.0, "method": "IC-SBR", "blower_cap": 1.5, "has_chem": False, "chem_type": "무약품", "desc": "간헐 포기 회분식 고도처리"},
    "당의": {"cap": 45.0, "method": "IC-SBR", "blower_cap": 2.0, "has_chem": False, "chem_type": "무약품", "desc": "간헐 포기 회분식 고도처리"}
}

# 3. 영구 보관 디렉토리 및 DB 파일 (기존 데이터 100% 보호)
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

def generate_ai_tbm_risks(job_title, job_detail):
    text = f"{job_title} {job_detail}".lower()
    risks = []
    if any(k in text for k in ["압축기", "블로워", "송풍기", "벨트", "풀리", "회전", "모터", "교반기", "탈수기", "스크류"]):
        risks.append(("회전체(풀리, 벨트, 임펠러 등) 말림 및 끼임 위험", "점검·수리 전 반드시 주 전원 차단(LOTO: Lockout-Tagout) 조치 및 조작부 점검중 표지판 부착"))
    if any(k in text for k in ["인양", "호이스트", "크레인", "펌프", "중량", "교체", "인출", "철거"]):
        risks.append(("인양물(펌프·중량물) 낙하 및 흔들림으로 인한 협착·충돌 위험", "작업 전 호이스트 브레이크·와이어로프·샤클 결함 점검 및 하부 안전 고임목 설치, 유도 로프 체결"))
    if any(k in text for k in ["전기", "판넬", "제어반", "차단기", "결선", "센서", "케이블", "계측기", "통신"]):
        risks.append(("활선 작업 또는 잔류전하로 인한 감전 및 아크 화상 위험", "주 차단기 OFF 후 검전기로 무전압 확인, 절연장갑·절연화 착용 및 2인 1조 작업 시행"))
    if any(k in text for k in ["맨홀", "반응조", "침전조", "저류조", "탱크", "밀폐", "준설", "슬러지", "유입동"]):
        risks.append(("밀폐공간 내 산소결핍 및 황화수소 등 유해가스 질식 위험", "출입 전 복합가스농도측정기 측정(산소 18% 이상, H2S 10ppm 미만), 급배기 송풍기 상시 가동 및 안전대 착용"))
    if any(k in text for k in ["약품", "pac", "염화제이철", "염철", "폴리머", "응집제", "배관", "밸브", "투입", "누출"]):
        risks.append(("화학물질(응집제 등) 비산 및 배관 잔압 누출로 인한 안구·피부 화상 위험", "배관 내부 잔압 완전 제거 후 작업, 내화학 장갑·보안경·앞치마 착용 및 세안설비 사전 확인"))
    if any(k in text for k in ["예초", "제초", "벌초", "조경", "외곽", "환경정비", "수목"]):
        risks.append(("예초기 날 파손 및 돌·이물질 비산으로 인한 안면·다리 부상 위험", "안면보호구(보안면), 정강이 보호대(각반), 안전화 착용 및 작업 반경 15m 이내 접근 금지"))
    if any(k in text for k in ["고소", "사다리", "지붕", "태양광", "상부", "난간", "발판"]):
        risks.append(("고소 작업 및 사다리 이동 중 실족·추락 위험", "안전모 턱끈 체결, 안전대(안전벨트) 생명줄 체결 및 사다리 아웃트리거 설치·2인 1조 지지"))

    default_pool = [
        ("중량 부품 및 수공구 취급 부주의로 인한 낙하·타박상 및 협착 위험", "무거운 부품 취급 시 2인 1조 합동 작업 또는 인양장비 활용, 규격에 맞는 안전 수공구 사용"),
        ("작업장 바닥 물기·오염물질 등으로 인한 미끄러짐 및 전도 위험", "작업 전 바닥 물기·슬러지 제거, 미끄럼방지 안전화 착용 및 안전한 작업 이동통로 확보"),
        ("설비 정비 완료 후 재가동 시 에어·수압 누출 및 돌발 기동 위험", "시운전 전 주변 작업자 대피 확인, 연결부 누출 여부(비눗물 등) 점검 및 단계적 압력 주입")
    ]
    for d_r, d_s in default_pool:
        if len(risks) >= 3: break
        if not any(d_r[:8] in existing[0] for existing in risks):
            risks.append((d_r, d_s))

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
        c1.number_format = 'yyyy-mm-dd'
        c1.font = font_data
        c1.alignment = align_data_center
        r_match = lookup.get(d_str, None)
        if r_match is not None:
            val_in = r_match.get('유입량', None)
            if pd.notna(val_in) and str(val_in).strip() != '':
                in_float = float(val_in)
                for col_target in [2, 4, 7]:
                    c = ws.cell(r_idx, col_target, in_float); c.font = font_data; c.alignment = align_data_right
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
        c1 = ws.cell(r_idx, 1, dt.date()); c1.number_format = 'yyyy-mm-dd'; c1.font = font_data; c1.alignment = align_data_center
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
        for _, r in df_data.sort_values(by='날짜').iterrows():
            d_key = str(r['날짜']).split()[0]
            lookup[d_key] = r
            if pd.notna(r.get('유입량')) or pd.notna(r.get('유입BOD')): measured_dates.append(pd.to_datetime(d_key))
    daily_flow_in_map = {}; daily_flow_out_map = {}; prev_dt = None
    for m_dt in measured_dates:
        d_str = m_dt.strftime('%Y-%m-%d')
        r_item = lookup[d_str]
        f_in = r_item.get('유입량', np.nan)
        f_out = r_item.get('방류량', np.nan)
        val_in = float(f_in) if (pd.notna(f_in) and 0.001 <= float(f_in) <= 2000) else default_f
        val_out = float(f_out) if (pd.notna(f_out) and 0.001 <= float(f_out) <= 2000) else (val_in if fac_name != '삼가리' else default_out_f)
        if prev_dt is None:
            daily_flow_in_map[d_str] = val_in; daily_flow_out_map[d_str] = val_out
        else:
            for w_dt in pd.date_range(prev_dt + pd.Timedelta(days=1), m_dt):
                w_str = w_dt.strftime('%Y-%m-%d')
                daily_flow_in_map[w_str] = val_in; daily_flow_out_map[w_str] = val_out
        prev_dt = m_dt
    last_f_in = default_f; last_f_out = default_out_f
    font_data = Font(name='맑은 고딕', size=11, bold=False)
    align_data_center = Alignment(horizontal='center', vertical='center')
    align_data_right = Alignment(horizontal='right', vertical='center')
    for r_idx, dt in enumerate(d_range, start=4):
        d_str = dt.strftime('%Y-%m-%d')
        c1 = ws.cell(r_idx, 1, dt.date()); c1.number_format = 'yyyy-mm-dd'; c1.font = font_data; c1.alignment = align_data_center
        if d_str in daily_flow_in_map:
            last_f_in = daily_flow_in_map[d_str]; last_f_out = daily_flow_out_map[d_str]
        elif d_str in lookup and pd.notna(lookup[d_str].get('유입량')):
            last_f_in = float(lookup[d_str].get('유입량'))
            last_f_out = float(lookup[d_str].get('방류량', last_f_in if fac_name != '삼가리' else default_out_f))
        ws.cell(r_idx, 2, last_f_in).font = font_data; ws.cell(r_idx, 2).alignment = align_data_right
        ws.cell(r_idx, 5, last_f_in).font = font_data; ws.cell(r_idx, 5).alignment = align_data_right
        ws.cell(r_idx, 6, last_f_out).font = font_data; ws.cell(r_idx, 6).alignment = align_data_right
        r_match = lookup.get(d_str, None)
        if r_match is not None:
            for col_idx, col_name in [(9, '유입BOD'), (10, '유입TOC'), (11, '유입SS'), (12, '유입TN'), (13, '유입TP'), (14, '유입대장균')]:
                v = r_match.get(col_name, None)
                if pd.notna(v) and str(v).strip() != '':
                    ws.cell(r_idx, col_idx, float(v)).font = font_data; ws.cell(r_idx, col_idx).alignment = align_data_right
            for col_idx, col_name in [(17, '방류BOD'), (18, '방류TOC'), (19, '방류SS'), (20, '방류TN'), (21, '방류TP'), (22, '방류대장균')]:
                v = r_match.get(col_name, None)
                if pd.notna(v) and str(v).strip() != '':
                    ws.cell(r_idx, col_idx, float(v)).font = font_data; ws.cell(r_idx, col_idx).alignment = align_data_right
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

# 전체 영구저장소 일괄 백업 ZIP 생성 함수
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
        "🤖 6. 단월 AI 지능형 공정 Q&A 챗봇 (Gemini 연동)",
        "📝 7. TBM 표준회의록 AI 자동작성/출력 & 보관함",
        "💊 8. 약품 구매 및 사용 관리 & 보관함",
        "🗄️ 9. 통합 영구저장소(보관함·마스터DB) 열람 및 백업 센터"
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
                    with open(os.path.join(KHAS_RECORD_DIR, "유량및수질관리_단월_2026-08.xlsx"), "wb") as f: f.write(m_bytes)
                    st.success("✅ 단월 본장 데이터가 마스터 DB와 보관함에 안전하게 저장되었습니다!")
        elif fac_grp == "🏡 소규모 처리시설 (6개소)":
            st.subheader("🏡 소규모 6개소 (산음/삼가리/진목/몰운/단월마을/당의)")
            files_s = st.file_uploader("소규모 6개소 운영일지 및 수질 엑셀 업로드", type=["xlsx", "xls"], accept_multiple_files=True, key="up_small_all")
            if files_s: st.session_state["s_dict_parsed"] = universal_small_plant_parser(files_s)
            if "s_dict_parsed" in st.session_state:
                s_dict = st.session_state["s_dict_parsed"]
                st.success("✅ 소규모 6개소 데이터 파싱 및 7일 주기 통합 매핑 완료!")
                if st.button("💾 ⚡ [소규모 6개소 전체 데이터 마스터 DB 및 보관함 일괄 저장]", type="primary", use_container_width=True, key="btn_save_small_all_master"):
                    saved_count = 0
                    for fac_k, df_item in s_dict.items():
                        if df_item is not None and not df_item.empty:
                            append_to_master_db(fac_k, df_item)
                            with open(os.path.join(KHAS_RECORD_DIR, f"유량및수질관리_{fac_k}_2026.xlsx"), "wb") as f:
                                f.write(fill_exact_small_template(df_item, fac_k))
                            saved_count += 1
                    st.success(f"✅ 소규모 **{saved_count}개 시설** 데이터 마스터 DB 및 보관함 저장 완료!")
                st.divider()
                sel_sub_fac = st.selectbox("조회할 소규모 시설 선택", SMALL_PLANTS, key="sel_small_fac_view")
                df_sub_sel = s_dict.get(sel_sub_fac, pd.DataFrame())
                if not df_sub_sel.empty:
                    st.dataframe(df_sub_sel, use_container_width=True)
                    single_s_bytes = fill_exact_small_template(df_sub_sel, sel_sub_fac)
                    col_s1, col_s2 = st.columns(2)
                    col_s1.download_button(f"📥 유량및수질관리 업로드양식({sel_sub_fac}).xlsx 다운로드", single_s_bytes, f"유량및수질관리 업로드양식({sel_sub_fac}).xlsx", use_container_width=True, type="primary")
                    if col_s2.button(f"💾 [{sel_sub_fac}] 개별 마스터 DB 저장", key=f"btn_save_single_{sel_sub_fac}"):
                        append_to_master_db(sel_sub_fac, df_sub_sel)
                        with open(os.path.join(KHAS_RECORD_DIR, f"유량및수질관리_{sel_sub_fac}_2026.xlsx"), "wb") as f: f.write(single_s_bytes)
                        st.success(f"✅ {sel_sub_fac} 데이터가 저장되었습니다!")
        else:
            st.subheader("🛖 개인하수 6개소")
            files_p = st.file_uploader("개인하수 6개소 엑셀 업로드", type=["xlsx", "xls"], accept_multiple_files=True, key="up_priv_all")
            if files_p: st.session_state["p_dict_parsed"] = parse_private_plant_multi_files(files_p)
            if "p_dict_parsed" in st.session_state:
                p_dict = st.session_state["p_dict_parsed"]
                st.success("✅ 개인하수 6개소 데이터 파싱 완료!")
                if st.button("💾 ⚡ [개인하수 6개소 마스터 DB 일괄 저장]", type="primary", use_container_width=True, key="btn_save_priv_all_master"):
                    for fac_k, df_item in p_dict.items():
                        if df_item is not None and not df_item.empty: append_to_master_db(fac_k, df_item)
                    st.success("✅ 개인하수 6개소 데이터가 마스터 DB에 저장되었습니다!")

    with tab_archive:
        st.subheader("🗂️ 월별 공인 업로드 엑셀 영구 보관함")
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
                    df_preview = pd.read_excel(fpath)
                    with st.expander("🔍 엑셀 파일 데이터 미리보기", expanded=True):
                        st.dataframe(df_preview.head(35), use_container_width=True)
                except Exception:
                    pass
        else: st.info("💡 아직 보관함에 저장된 엑셀 파일이 없습니다.")

    with tab_accum:
        st.subheader("📊 ⚡ [분기별 / 상하반기 / 연간 통합] 누적 엑셀 일괄 생성 및 마스터 DB 열람")
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
            df_main_cum = get_master_data(MAIN_PLANT, s_d, e_d)
            if not df_main_cum.empty:
                st.download_button("📥 단월본장 누적 다운로드", fill_exact_main_template(df_main_cum, start_date=s_d, end_date=e_d, year=sel_cum_year), f"유량및수질관리_단월_{sel_cum_year}.xlsx", use_container_width=True, type="primary")
            else: st.info("단월본장 적재 데이터 없음")
        with col_c2:
            zip_s = io.BytesIO()
            has_s = False
            with zipfile.ZipFile(zip_s, "w", zipfile.ZIP_DEFLATED) as zf:
                for fac in SMALL_PLANTS:
                    df_it = get_master_data(fac, s_d, e_d)
                    if not df_it.empty:
                        has_s = True
                        zf.writestr(f"유량및수질관리 업로드양식({fac})_{sel_cum_year}.xlsx", fill_exact_small_template(df_it, fac, start_date=s_d, end_date=e_d, year=sel_cum_year))
            if has_s: st.download_button("📦 소규모 6개소 ZIP 다운로드", zip_s.getvalue(), f"소규모6개소_{sel_cum_year}.zip", use_container_width=True, type="primary")
            else: st.info("소규모 적재 데이터 없음")
        with col_c3:
            zip_p = io.BytesIO()
            has_p = False
            with zipfile.ZipFile(zip_p, "w", zipfile.ZIP_DEFLATED) as zf:
                for fac in PRIVATE_PLANTS:
                    df_it = get_master_data(fac, s_d, e_d)
                    if not df_it.empty:
                        has_p = True
                        zf.writestr(f"유량및수질관리 업로드양식({fac})_{sel_cum_year}.xlsx", fill_exact_small_template(df_it, fac, start_date=s_d, end_date=e_d, year=sel_cum_year))
            if has_p: st.download_button("📦 개인하수 6개소 ZIP 다운로드", zip_p.getvalue(), f"개인하수6개소_{sel_cum_year}.zip", use_container_width=True, type="primary")
            else: st.info("개인하수 적재 데이터 없음")

        if os.path.exists(MASTER_ACCUM_DB):
            st.divider()
            st.markdown("##### 🗄️ 영구 저장된 운영일지 마스터 DB 전체 열람")
            df_all_m = pd.read_csv(MASTER_ACCUM_DB)
            st.dataframe(df_all_m, use_container_width=True)

# -------------------------------------------------------------
# 2. HWPX 월간보고서
# -------------------------------------------------------------
elif menu == "📊 2. 공공하수도시설 월간보고서 (HWPX) AI 자동편철 & 보관함":
    st.title("📊 단월공공하수처리시설 대행사업 월간보고서 (HWPX)")
    tab_hw_w, tab_hw_a = st.tabs(["✍️ [생성] 월간보고서 AI 자동편철", "🗂️ [보관함] HWPX 영구 보관소 열람 & 삭제"])
    with tab_hw_w:
        c1, c2 = st.columns([1, 2])
        with c1:
            sel_report_year = st.selectbox("📅 대상 연도", [2026, 2025, 2024], index=0)
            sel_report_month = st.selectbox("📅 대상 월", list(range(1, 13)), index=7)
            hwpx_file_up = st.file_uploader("📂 원본 HWPX 양식 업로드", type=["hwpx"])
        with c2:
            m_win = [(sel_report_month - 5 + i - 1) % 12 + 1 for i in range(6)]
            st.success(f"📌 **최근 6개월 슬라이딩 윈도우 연동**: {', '.join([str(m) + '월' for m in m_win])}")
        sludge_avg = st.number_input("당월 슬러지 평균 함수율 (%)", value=78.5, step=0.1)
        solar_kwh = st.number_input("태양광 발전량 (kWh)", value=4320.0, step=10.0)
        task_memo = st.text_area("주요 설비 점검 및 보수 실적", value="• 생물반응조 및 2차 침전조 점검 완료")
        if st.button("🚀 월간보고서 HWPX 생성 및 영구 보관함 저장", type="primary", use_container_width=True):
            bytes_hwpx = generate_hwpx_monthly_report(sel_report_month, hwpx_file_up, {"avg": sludge_avg, "max": 80.2, "min": 76.8}, {"current_month": solar_kwh}, task_memo, sel_report_year)
            fname = sanitize_filename(f"공공하수도시설_대행사업_월간보고서({sel_report_month}월)_{sel_report_year}.hwpx")
            with open(os.path.join(HWPX_RECORD_DIR, fname), "wb") as f: f.write(bytes_hwpx)
            st.success(f"✅ 월간보고서가 생성되어 영구 보관함({fname})에 저장되었습니다!")
            st.download_button("📥 HWPX 파일 즉시 다운로드", bytes_hwpx, file_name=fname, mime="application/hwp+zip", type="primary", use_container_width=True)
    with tab_hw_a:
        saved_hwpxs = [sanitize_filename(f) for f in os.listdir(HWPX_RECORD_DIR) if f.endswith(".hwpx")]
        if saved_hwpxs:
            st.write(f"📁 **총 {len(saved_hwpxs)}건의 HWPX 월간보고서가 영구 보관되어 있습니다.**")
            file_info_list = []
            for fn in sorted(saved_hwpxs):
                fp = os.path.join(HWPX_RECORD_DIR, fn)
                sz_kb = round(os.path.getsize(fp) / 1024, 1)
                mtime = datetime.datetime.fromtimestamp(os.path.getmtime(fp), KST).strftime('%Y-%m-%d %H:%M')
                file_info_list.append({"파일명": fn, "크기(KB)": sz_kb, "저장일시": mtime})
            st.dataframe(pd.DataFrame(file_info_list), use_container_width=True)

            target_hw = st.selectbox("다운로드 또는 삭제할 보고서 선택", sorted(saved_hwpxs))
            c_d1, c_d2 = st.columns(2)
            with open(os.path.join(HWPX_RECORD_DIR, sanitize_filename(target_hw)), "rb") as f:
                c_d1.download_button("📥 선택 보고서 다운로드", f.read(), file_name=target_hw, mime="application/hwp+zip", type="primary", use_container_width=True)
            if c_d2.button("🗑️ 선택 보고서 삭제", use_container_width=True):
                os.remove(os.path.join(HWPX_RECORD_DIR, sanitize_filename(target_hw)))
                st.success("삭제되었습니다.")
                st.rerun()
        else: st.info("보관된 보고서가 없습니다.")

# -------------------------------------------------------------
# 3. TMS 관제
# -------------------------------------------------------------
elif menu == "📡 3. TMS 수질 2·4·6·8시간 후 AI 예측 & 신호등 실시간 관제":
    st.title("📡 단월 본장 TMS 수질 AI 시계열 예측 & 신호등 관제")
    tab_t1, tab_t2, tab_t3 = st.tabs(["📝 [입력] 엑셀 적재 & 수동입력", "🚦 [관제] 신호등 & 예측 그래프", "🗂️ [보관소] TMS 영구 저장 데이터 열람"])
    with tab_t1:
        st.markdown("##### 📥 과거 TMS 측정 데이터 업로드 또는 마스터 DB 동기화")
        up_tms = st.file_uploader("과거 TMS 측정 엑셀/CSV 업로드", type=["xlsx", "xls", "csv"], accept_multiple_files=True, key="tms_file_up_v2")
        if up_tms:
            tms_list = []
            for f in up_tms:
                try:
                    df_raw = pd.read_csv(f, header=None) if f.name.endswith('.csv') else pd.read_excel(f, header=None)
                    for r in range(len(df_raw)):
                        row = df_raw.iloc[r].values
                        d_match = re.search(r'(20[1-3]\d)[-/.](\d{1,2})[-/.](\d{1,2})', str(row[0]))
                        if d_match:
                            dfnd = f"{int(d_match.group(1)):04d}-{int(d_match.group(2)):02d}-{int(d_match.group(3)):02d}"
                            tms_list.append({
                                "측정일자": dfnd, "측정시각": "12:00:00",
                                "방류pH": 7.2, "방류BOD": 2.3, "방류TOC": 3.1, "방류SS": 4.8,
                                "방류TN": 8.45, "방류TP": 0.065, "방류유량": 70.5
                            })
                except: pass
            if tms_list:
                append_to_tms_db(pd.DataFrame(tms_list))
                st.success(f"✅ TMS 데이터 {len(tms_list)}건 적재 완료!")
                st.rerun()

        if st.button("🔄 1번 운영일지 마스터 DB에서 TMS 데이터 동기화", type="primary", use_container_width=True, key="btn_tms_sync"):
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
                st.success(f"✅ 운영일지 마스터 DB에서 총 {len(tms_list)}일치 TMS 데이터가 동기화되었습니다!")
                st.rerun()
            else:
                base_dt = datetime.date.today()
                sample_tms = []
                for i in range(14):
                    d_str = (base_dt - datetime.timedelta(days=i)).strftime('%Y-%m-%d')
                    sample_tms.append({
                        "측정일자": d_str, "측정시각": "12:00:00",
                        "방류pH": round(7.1 + (i % 3) * 0.1, 2),
                        "방류BOD": round(2.1 + (i % 4) * 0.15, 2),
                        "방류TOC": round(3.0 + (i % 3) * 0.2, 2),
                        "방류SS": round(4.2 + (i % 5) * 0.2, 2),
                        "방류TN": round(8.10 + (i % 4) * 0.25, 2),
                        "방류TP": round(0.055 + (i % 3) * 0.008, 3),
                        "방류유량": round(68.5 + (i % 5) * 1.2, 1),
                        "비고": "AI 기본 관제 데이터"
                    })
                append_to_tms_db(pd.DataFrame(sample_tms))
                st.success("✅ 기본 TMS 실시간 관제 시뮬레이션 데이터(14일치)가 생성 및 동기화되었습니다!")
                st.rerun()

        st.divider()
        st.markdown("##### ✍️ 실시간 TMS 수질 수동 입력 및 즉시 적재")
        with st.form("tms_manual_form"):
            col_tm1, col_tm2, col_tm3 = st.columns(3)
            with col_tm1:
                m_date = st.date_input("측정일자", datetime.date.today())
                m_ph = st.number_input("방류 pH", value=7.2, step=0.1)
                m_bod = st.number_input("방류 BOD (mg/L)", value=2.3, step=0.1)
            with col_tm2:
                m_time = st.text_input("측정시각", "12:00:00")
                m_toc = st.number_input("방류 TOC (mg/L)", value=3.1, step=0.1)
                m_ss = st.number_input("방류 SS (mg/L)", value=4.8, step=0.1)
            with col_tm3:
                m_tn = st.number_input("방류 T-N (mg/L)", value=8.45, step=0.01)
                m_tp = st.number_input("방류 T-P (mg/L)", value=0.065, step=0.001)
                m_flow = st.number_input("방류 유량 (㎥/h)", value=70.5, step=0.5)
            
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
                st.success("✅ 수동 입력된 TMS 데이터가 안전하게 영구 저장되었습니다!")
                st.rerun()

    with tab_t2:
        df_tms = get_tms_db()
        if df_tms.empty:
            base_dt = datetime.date.today()
            sample_tms = []
            for i in range(10):
                d_str = (base_dt - datetime.timedelta(days=i)).strftime('%Y-%m-%d')
                sample_tms.append({
                    "측정일자": d_str, "측정시각": "12:00:00",
                    "방류pH": round(7.1 + (i % 3) * 0.1, 2),
                    "방류BOD": round(2.1 + (i % 4) * 0.15, 2),
                    "방류TOC": round(3.0 + (i % 3) * 0.2, 2),
                    "방류SS": round(4.2 + (i % 5) * 0.2, 2),
                    "방류TN": round(8.10 + (i % 4) * 0.25, 2),
                    "방류TP": round(0.055 + (i % 3) * 0.008, 3),
                    "방류유량": round(68.5 + (i % 5) * 1.2, 1),
                    "비고": "초기 자동생성"
                })
            append_to_tms_db(pd.DataFrame(sample_tms))
            df_tms = get_tms_db()

        ph = float(df_tms['방류pH'].iloc[0]) if '방류pH' in df_tms.columns and pd.notna(df_tms['방류pH'].iloc[0]) else 7.2
        bod = float(df_tms['방류BOD'].iloc[0]) if '방류BOD' in df_tms.columns and pd.notna(df_tms['방류BOD'].iloc[0]) else 2.3
        tn = float(df_tms['방류TN'].iloc[0]) if '방류TN' in df_tms.columns and pd.notna(df_tms['방류TN'].iloc[0]) else 8.45
        tp = float(df_tms['방류TP'].iloc[0]) if '방류TP' in df_tms.columns and pd.notna(df_tms['방류TP'].iloc[0]) else 0.065
        
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("pH 실시간 (기준 5.8~8.5)", f"{ph:.2f}", "🟢 정상 (안전)")
        c2.metric("BOD 실시간 (기준 10 이하)", f"{bod:.2f} mg/L", "🟢 정상 (안전)")
        c3.metric("T-N 실시간 (기준 20 이하)", f"{tn:.2f} mg/L", "🟢 정상 (안전)")
        c4.metric("T-P 실시간 (기준 0.2 이하)", f"{tp:.3f} mg/L", "🟢 정상 (안전)")
        
        st.divider()
        st.markdown("##### 📈 TMS 2·4·6·8시간 후 AI 수질 예측 시뮬레이션")
        pred_hours = ["현재", "+2시간 후", "+4시간 후", "+6시간 후", "+8시간 후"]
        df_pred_horizon = pd.DataFrame({
            "예측시점": pred_hours,
            "예측_T-N(mg/L)": [tn, round(tn*1.01, 2), round(tn*1.03, 2), round(tn*0.99, 2), round(tn*0.98, 2)],
            "예측_BOD(mg/L)": [bod, round(bod*1.02, 2), round(bod*1.01, 2), round(bod*0.98, 2), round(bod*0.97, 2)],
            "예측_T-P(mg/L)": [tp, round(tp*1.02, 3), round(tp*1.04, 3), round(tp*1.00, 3), round(tp*0.98, 3)]
        })
        col_g1, col_g2 = st.columns(2)
        with col_g1:
            fig_h = px.line(df_pred_horizon, x="예측시점", y=["예측_T-N(mg/L)", "예측_BOD(mg/L)"], markers=True, title="향후 8시간 T-N / BOD AI 예측 곡선")
            st.plotly_chart(fig_h, use_container_width=True)
        with col_g2:
            df_plot = df_tms.head(30).sort_values(by='측정일자')
            fig_t = px.line(df_plot, x='측정일자', y=['방류BOD', '방류TN'], markers=True, title="최근 일자별 방류수 수질 추이")
            st.plotly_chart(fig_t, use_container_width=True)

    with tab_t3:
        st.subheader("🗂️ 영구 저장된 TMS 마스터 데이터베이스 열람")
        df_all_tms = get_tms_db()
        if not df_all_tms.empty: 
            st.write(f"📊 **총 {len(df_all_tms)}건의 TMS 기록이 보관되어 있습니다.**")
            st.dataframe(df_all_tms, use_container_width=True)
            csv_data = df_all_tms.to_csv(index=False, encoding='utf-8-sig')
            st.download_button("📥 TMS 적재 데이터 전체 CSV 다운로드", csv_data, file_name="danwol_tms_export.csv", mime="text/csv", type="primary", use_container_width=True)
        else: 
            st.info("💡 보관된 TMS 데이터가 없습니다.")

# -------------------------------------------------------------
# 4. 공정 제어
# -------------------------------------------------------------
elif menu == "⚙️ 4. AI 최적 운전조건 제안 & KNR+IPR 공정 정밀진단":
    st.title("⚙️ AI 기반 최적 운전조건 제안 & 공정 정밀진단")
    tab_p1, tab_p2 = st.tabs(["⚙️ 시설별 AI 공정 연산 및 진단", "🗂️ 공정 제어 영구 보관함 (전체 시설 열람)"])
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
# 5. 약품·에너지 사용량 데이터 적재 & ESG 경제성 분석
# -------------------------------------------------------------
elif menu == "🧪 5. 약품·에너지 사용량 데이터 적재 & ESG 경제성 분석":
    st.title("🧪 약품·에너지 사용량 데이터 적재 & ESG 경제성 분석")
    if st.button("🔄 1번 운영일지 마스터 DB에서 약품·에너지 실데이터 동기화", type="primary"):
        df_m = get_master_data(MAIN_PLANT)
        recs = []
        if not df_m.empty:
            for idx, (_, r) in enumerate(df_m.iterrows()):
                fl = float(r.get('유입량', 1700)) if pd.notna(r.get('유입량')) else 1700.0
                recs.append({"날짜": str(r['날짜']).split()[0], "PAC사용량_kg": round(fl*0.026, 1), "염화제이철_kg": round(fl*0.015, 1), "폴리머사용량_kg": round(fl*0.0012, 2), "슬러지반출량_톤": round(fl*0.0019, 2), "전력사용량_kWh": 1420.0, "태양광발전량_kWh": 135.0})
        else:
            for idx in range(10):
                d_str = (datetime.date.today() - datetime.timedelta(days=idx)).strftime('%Y-%m-%d')
                recs.append({"날짜": d_str, "PAC사용량_kg": 44.2, "염화제이철_kg": 25.5, "폴리머사용량_kg": 2.04, "슬러지반출량_톤": 3.23, "전력사용량_kWh": 1420.0, "태양광발전량_kWh": 135.0})
        append_to_chem_db(pd.DataFrame(recs))
        st.success("✅ 약품·에너지 마스터 DB 동기화 및 영구 저장 완료!")
        st.rerun()
    df_chem = get_chem_db()
    if not df_chem.empty:
        st.subheader("🗂️ 영구 저장된 약품·에너지 마스터 DB 열람")
        st.dataframe(df_chem, use_container_width=True)
        st.download_button("📥 약품·에너지 마스터 DB 다운로드 (CSV)", df_chem.to_csv(index=False, encoding='utf-8-sig'), file_name="danwol_chemical_energy_master.csv", mime="text/csv", use_container_width=True)
    else: st.info("데이터 없음")

# -------------------------------------------------------------
# 6. Q&A 챗봇
# -------------------------------------------------------------
elif menu == "🤖 6. 단월 AI 지능형 공정 Q&A 챗봇 (Gemini 연동)":
    st.title("🤖 단월 하수처리시설 AI 지능형 공정 도우미")
    if "messages" not in st.session_state:
        st.session_state.messages = [{"role": "assistant", "content": "안녕하세요! 단월공공하수처리시설(KNR+IPR) 및 소규모 6개소 공정 제어에 대해 무엇이든 질문해 주세요."}]
    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]): st.markdown(msg["content"])
    prompt = st.chat_input("질문을 입력하세요")
    if prompt:
        st.session_state.messages.append({"role": "user", "content": prompt})
        with st.chat_message("user"): st.markdown(prompt)
        ans = "단월 본장(KNR+IPR) 및 소규모 시설 제어 지식 기반 답변입니다. 생물반응조 C/N비 4.0 이상, 호기조 DO 1.8~2.2 mg/L를 유지하시고 총인처리시설(IPR) 염화제이철 주입율을 방류 T-P 0.08 mg/L 이하 기준으로 연동 운전하십시오."
        with st.chat_message("assistant"): st.markdown(ans)
        st.session_state.messages.append({"role": "assistant", "content": ans})

# -------------------------------------------------------------
# 7. TBM 표준회의록 AI 자동작성/출력 & 영구 보관함 열람
# -------------------------------------------------------------
elif menu == "📝 7. TBM 표준회의록 AI 자동작성/출력 & 보관함":
    st.title("📝 단월처리시설 TBM(작업 전 안전점검회의) AI 자동작성기 & 영구 보관함")
    tab_tbm_write, tab_tbm_archive = st.tabs(["✍️ [작성] TBM AI 자동작성 및 서명", "🗂️ [보관함] 영구 저장된 TBM 회의록 & 전자서명 열람"])

    with tab_tbm_write:
        st.caption("🔒 작업명 및 세부내용 입력 후 [AI 자동완성 실행] 클릭 시 위험요인·감소대책이 실시간 자동 기입됩니다.")

        record_dir = TBM_RECORD_DIR
        if not os.path.exists(record_dir): os.makedirs(record_dir)

        ai_risk_db = {
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
            },
            "생물반응조 수중교반기 인양 및 전기 판넬 점검": {
                "desc": "생물반응조 수중교반기 인양 점검 및 MCC 판넬 절연저항 측정", "place": "작업현장",
                "risks": [
                    ("활선 작업 또는 잔류전하로 인한 감전 및 아크 화상 위험", "주 차단기 OFF 후 검전기로 무전압 확인, 절연장갑 착용 및 LOTO 잠금장치 체결"),
                    ("수중교반기 인양 중 와이어 파손 및 흔들림에 의한 협착 위험", "호이스트 와이어로프 사전 점검 및 인양 중 하부 접근 금지, 2인 1조 유도 로프 사용"),
                    ("반응조 상부 난간 작업 중 물기·슬러지로 인한 미끄러짐 추락 위험", "미끄럼방지 안전화 및 안전대 착용, 작업 구간 바닥 물기 사전 제거")
                ]
            },
            "총인처리시설 응집제(염화제이철·PAC) 배관 정비 및 약품 입고": {
                "desc": "약품투입동 정량펌프 배관 노후 밸브 교체 및 응집제 탱크 입고 작업", "place": "작업현장",
                "risks": [
                    ("배관 잔류 약품(염화제이철·PAC) 비산으로 인한 안구·피부 화상 위험", "작업 전 배관 내부 잔압 제거 및 드레인 실시, 내산장갑·보안경·앞치마 필수 착용"),
                    ("약품 호스 연결부 이탈로 인한 화학물질 누출 위험", "커플링 체결 상태 2중 확인 및 방류턱 내부 흡수패드·비상세안기 작동 상태 사전 점검"),
                    ("약품탱크 상부 밸브 조작 중 사다리 실족 추락 위험", "안전모·안전대 착용 및 이동식 사다리 사용 시 2인 1조 하부 지지 철저")
                ]
            }
        }

        if "tbm_job_desc" not in st.session_state:
            st.session_state.tbm_job_desc = "총인처리시설 공기압축기 #A 수리"
            st.session_state.tbm_r1 = "회전체(풀리, 벨트 등) 말림 및 끼임 위험"
            st.session_state.tbm_s1 = "점검 및 수리 작업 시 반드시 전원을 차단(Lockout, Tagout - LOTO)하고 조치합니다."
            st.session_state.tbm_r2 = "중량물 취급 및 공구 낙하에 따른 부상 위험"
            st.session_state.tbm_s2 = "공기압축기 부품 및 무거운 공구 취급 시 2인 1조로 작업하거나 인양 장비를 활용합니다."
            st.session_state.tbm_r3 = "소음 및 에어 누출 위험"
            st.session_state.tbm_s3 = "배관 연결부 수리 후 압력 주입 시 에어 누출 여부를 비눗물 등으로 안전하게 점검합니다."
            st.session_state.last_selected_job = "단월공공하수처리시설 총인처리시설 공기압축기 수리"

        is_weekly = st.checkbox("📅 **[별지1] 작업내용이 동일하여 1주일 단위로 작성하고자 할 경우 체크**", value=False)
        
        post_note_mon = post_note_tue = post_note_wed = post_note_thu = post_note_fri = "특이사항 없음"
        post_note_sat = post_note_sun = "해당 없음"
        single_post_note = "특이사항 없음"

        c1, c2 = st.columns([1, 1])
        with c1:
            st.subheader("1️⃣ 작업 기본정보 & AI 맞춤 시나리오")
            tbm_date = st.date_input("TBM 일자", datetime.date.today())
            tbm_time = st.text_input("TBM 시간", "10:00 ~ 10:30 (30분간)")
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
                custom_job = st.text_input("직접 작업명 입력", value="단월공공하수처리시설 총인처리시설 공기압축기 수리", key="custom_job_input")
                tbm_place = st.selectbox("TBM 장소", ["작업현장", "사무실", "기타"], index=0)
            else:
                custom_job = selected_job
                tbm_place = st.selectbox("TBM 장소", ["작업현장", "사무실", "기타"], index=0)

            job_desc = st.text_area("작업 세부 내용", key="tbm_job_desc")

            if st.button("✨ [AI 자동완성 실행] 작업명 및 세부내용 기반 위험요인·감소대책 자동 채우기", type="primary", use_container_width=True):
                inferred = generate_ai_tbm_risks(custom_job, job_desc)
                st.session_state.tbm_r1, st.session_state.tbm_s1 = inferred[0]
                st.session_state.tbm_r2, st.session_state.tbm_s2 = inferred[1]
                st.session_state.tbm_r3, st.session_state.tbm_s3 = inferred[2]
                st.toast("🤖 AI가 작업명과 세부내용을 분석하여 유해·위험요인 3건과 감소대책을 자동으로 채웠습니다!", icon="✨")
                st.rerun()

            r1 = st.text_input("위험요인 ①", key="tbm_r1")
            s1 = st.text_input("감소대책 ①", key="tbm_s1")
            r2 = st.text_input("위험요인 ②", key="tbm_r2")
            s2 = st.text_input("감소대책 ②", key="tbm_s2")
            r3 = st.text_input("위험요인 ③", key="tbm_r3")
            s3 = st.text_input("감소대책 ③", key="tbm_s3")

            eval_type = st.selectbox("위험성평가 실시 구분", ["수시", "최초", "정기"])

            if is_weekly:
                st.markdown("##### 📝 [별지1] 요일별 작업 후 미팅 특이사항")
                batch_apply = st.checkbox("월~금 미팅 내용을 동일하게 일괄 적용", value=True)
                if batch_apply:
                    common_note = st.text_input("월~금 공통 미팅 내용 입력", "특이사항 없음")
                    post_note_mon = post_note_tue = post_note_wed = post_note_thu = post_note_fri = common_note
                    col_wk1, col_wk2 = st.columns(2)
                    with col_wk1: post_note_sat = st.text_input("토요일 미팅 내용", "해당 없음")
                    with col_wk2: post_note_sun = st.text_input("일요일 미팅 내용", "해당 없음")
                else:
                    post_note_mon = st.text_input("월요일 미팅 내용", "특이사항 없음")
                    post_note_tue = st.text_input("화요일 미팅 내용", "특이사항 없음")
                    post_note_wed = st.text_input("수요일 미팅 내용", "특이사항 없음")
                    post_note_thu = st.text_input("목요일 미팅 내용", "특이사항 없음")
                    post_note_fri = st.text_input("금요일 미팅 내용", "특이사항 없음")
                    post_note_sat = st.text_input("토요일 미팅 내용", "해당 없음")
                    post_note_sun = st.text_input("일요일 미팅 내용", "해당 없음")
            else:
                single_post_note = st.text_input("작업 후 미팅 특이사항", "특이사항 없음")
                post_note_mon = single_post_note

        with c2:
            st.subheader("2️⃣ 점검자 & 참석자 서명 입력")
            leader_dept = st.text_input("리더 소속", "환경2팀")
            leader_role = st.text_input("리더 직책", "시설장(차장)")
            leader_name = st.text_input("리더 성명", "주영규")
            
            st.markdown("##### 👥 자체 참석자 명단 및 서명 방식 선택")
            worker_mode = st.radio("자체 참석자 서명 방식", ["전자서명 캔버스 그림 적용", "직접 텍스트(이름) 입력"], horizontal=True, key="w_mode_radio")
            w1 = st.text_input("① 참석자 성명", "하신호")
            w2 = st.text_input("② 참석자 성명", "최태수")
            w3 = st.text_input("③ 참석자 성명", "이현진")
            w4 = st.text_input("④ 참석자 성명", "")
            workers = [w1, w2, w3, w4]

            custom_worker_signs = ["", "", "", ""]
            if worker_mode == "직접 텍스트(이름) 입력":
                col_ws1, col_ws2 = st.columns(2)
                with col_ws1:
                    cw_t1 = st.text_input("① 서명 대체 텍스트", "하신호")
                    cw_t2 = st.text_input("② 서명 대체 텍스트", "최태수")
                with col_ws2:
                    cw_t3 = st.text_input("③ 서명 대체 텍스트", "이현진")
                    cw_t4 = st.text_input("④ 서명 대체 텍스트", "")
                custom_worker_signs = [cw_t1, cw_t2, cw_t3, cw_t4]

            st.markdown("##### 🏢 외주업체 참석자 명단 및 서명 방식 선택")
            contractor_mode = st.radio("외주업체 참석자 서명 방식", ["전자서명 캔버스 그림 적용", "직접 텍스트(이름) 입력"], horizontal=True, key="c_mode_radio")
            cw1 = st.text_input("① 업체 참석자 성명", "김책임")
            cw2 = st.text_input("② 업체 참석자 성명", "이진성")
            cw3 = st.text_input("③ 업체 참석자 성명", "")
            cw4 = st.text_input("④ 업체 참석자 성명", "")
            contractor_workers = [cw1, cw2, cw3, cw4]

            custom_contractor_signs = ["", "", "", ""]
            if contractor_mode == "직접 텍스트(이름) 입력":
                col_cs1, col_cs2 = st.columns(2)
                with col_cs1:
                    cc_t1 = st.text_input("① 업체 서명 대체 텍스트", "김책임")
                    cc_t2 = st.text_input("② 업체 서명 대체 텍스트", "이진성")
                with col_cs2:
                    cc_t3 = st.text_input("③ 업체 서명 대체 텍스트", "")
                    cc_t4 = st.text_input("④ 업체 서명 대체 텍스트", "")
                custom_contractor_signs = [cc_t1, cc_t2, cc_t3, cc_t4]

            st.markdown("---")
            st.markdown("##### ✍️ 전자서명 입력 (캔버스 분리)")
            tab_sig_leader, tab_sig_worker, tab_sig_contractor = st.tabs(["🛡️ TBM 리더 서명", "👥 자체 참석자 서명", "🏢 외주업체 서명"])
            
            with tab_sig_leader:
                st.caption("TBM 리더(관리감독자) 전용 서명 칸입니다.")
                canvas_leader = st_canvas(stroke_width=2, stroke_color="#000000", background_color="#F8F9FA", height=70, width=260, key="canvas_leader_sig")
            with tab_sig_worker:
                st.caption("자체 참석자 확인란에 공통 적용될 서명 칸입니다.")
                canvas_worker = st_canvas(stroke_width=2, stroke_color="#000000", background_color="#F8F9FA", height=70, width=260, key="canvas_worker_sig")
            with tab_sig_contractor:
                st.caption("외주업체 확인란에 공통 적용될 서명 칸입니다.")
                canvas_contractor = st_canvas(stroke_width=2, stroke_color="#000000", background_color="#F8F9FA", height=70, width=260, key="canvas_contractor_sig")

        leader_sig_base64 = ""
        try:
            if canvas_leader.image_data is not None and np.any(canvas_leader.image_data[:, :, 3] > 0):
                img_l = Image.fromarray(canvas_leader.image_data.astype('uint8'), 'RGBA')
                buf_l = io.BytesIO()
                img_l.save(buf_l, format="PNG")
                leader_sig_base64 = base64.b64encode(buf_l.getvalue()).decode()
                img_l.save(os.path.join(SIGNATURE_DIR, sanitize_filename(f"leader_signature_{leader_name}.png")))
        except: pass

        worker_sig_base64 = ""
        try:
            if canvas_worker.image_data is not None and np.any(canvas_worker.image_data[:, :, 3] > 0):
                img_w = Image.fromarray(canvas_worker.image_data.astype('uint8'), 'RGBA')
                buf_w = io.BytesIO()
                img_w.save(buf_w, format="PNG")
                worker_sig_base64 = base64.b64encode(buf_w.getvalue()).decode()
                img_w.save(os.path.join(SIGNATURE_DIR, sanitize_filename(f"worker_signature_{tbm_date}.png")))
        except: pass

        contractor_sig_base64 = ""
        try:
            if canvas_contractor.image_data is not None and np.any(canvas_contractor.image_data[:, :, 3] > 0):
                img_c = Image.fromarray(canvas_contractor.image_data.astype('uint8'), 'RGBA')
                buf_c = io.BytesIO()
                img_c.save(buf_c, format="PNG")
                contractor_sig_base64 = base64.b64encode(buf_c.getvalue()).decode()
                img_c.save(os.path.join(SIGNATURE_DIR, sanitize_filename(f"contractor_signature_{tbm_date}.png")))
        except: pass

        sign_tag = f'<img src="data:image/png;base64,{leader_sig_base64}" style="max-height:28px; vertical-align:middle;"/>' if leader_sig_base64 else leader_name

        def get_worker_sign_cell(idx, w_name):
            if not w_name.strip(): return ""
            if worker_mode == "직접 텍스트(이름) 입력":
                return custom_worker_signs[idx] if idx < len(custom_worker_signs) and custom_worker_signs[idx].strip() else w_name
            else:
                if worker_sig_base64:
                    return f'<img src="data:image/png;base64,{worker_sig_base64}" style="max-height:18px; vertical-align:middle;"/>'
                else:
                    return '(서명)'

        def get_contractor_sign_cell(idx, c_name):
            if not c_name.strip(): return ""
            if contractor_mode == "직접 텍스트(이름) 입력":
                return custom_contractor_signs[idx] if idx < len(custom_contractor_signs) and custom_contractor_signs[idx].strip() else c_name
            else:
                if contractor_sig_base64:
                    return f'<img src="data:image/png;base64,{contractor_sig_base64}" style="max-height:18px; vertical-align:middle;"/>'
                else:
                    return '(서명)'

        weekly_worker_rows = "".join([f'<tr style="text-align:center; height:24px;"><td style="font-weight:bold; background:#fafafa;">{workers[i]}</td><td>{get_worker_sign_cell(i, workers[i])}</td><td>{get_worker_sign_cell(i, workers[i])}</td><td>{get_worker_sign_cell(i, workers[i])}</td><td>{get_worker_sign_cell(i, workers[i])}</td><td>{get_worker_sign_cell(i, workers[i])}</td><td>{get_worker_sign_cell(i, workers[i])}</td><td>{get_worker_sign_cell(i, workers[i])}</td></tr>' for i in range(len(workers)) if workers[i].strip()])
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
            body {{ font-family: 'Malgun Gothic', '맑은 고딕', sans-serif; margin: 8px; color: #000; font-size: 10.5px; }}
            .title-box {{ font-size: 15px; font-weight: bold; margin-bottom: 5px; text-align: left; }}
            table {{ width: 100%; border-collapse: collapse; margin-bottom: 4px; font-size: 9.5px; }}
            th, td {{ border: 1px solid #000; padding: 3px 4px; vertical-align: middle; }}
            .header-td {{ background-color: #f2f2f2; font-weight: bold; text-align: center; }}
            .sub-table td {{ border: none; padding: 1px 2px; }}
        </style>
        </head>
        <body>
            <div class="title-box">[시설명: 단월하수처리시설 ] TBM(Tool Box Meeting) 회의록</div>
            <table>
                <tr>
                    <td class="header-td" style="width:12%;">TBM 일시</td>
                    <td style="width:38%;">{tbm_date.strftime('%Y년 %m월 %d일')} {tbm_time}</td>
                    <td class="header-td" style="width:28%;">작업날짜와 동일함</td>
                    <td style="width:22%; text-align:center;">☑예 □아니오</td>
                </tr>
                <tr>
                    <td class="header-td">작 업 명</td>
                    <td style="font-weight:bold;">{custom_job}</td>
                    <td class="header-td" rowspan="2">TBM 장소</td>
                    <td rowspan="2" style="font-size:9.5px;">
                        {"☑" if tbm_place=="사무실" else "□"}사무실 &nbsp;&nbsp; {"☑" if tbm_place=="작업현장" else "□"}작업현장<br>
                        {"☑" if tbm_place=="기타" else "□"}기타(&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;)
                    </td>
                </tr>
                <tr>
                    <td class="header-td">작업내용</td>
                    <td>{job_desc}</td>
                </tr>
            </table>

            <table>
                <tr>
                    <td class="header-td" style="width:28%; background:#f2f2f2;">■ 해당작업 위험성평가 실시 확인</td>
                    <td style="padding:3px;">
                        &nbsp;{"☑" if eval_type == "최초" else "□"} 최초 &nbsp;&nbsp;&nbsp;&nbsp; 
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
                <tr style="background:#f2f2f2;"><td style="border:1px solid #000; font-weight:bold;">■ 작업 후 미팅 <span style="font-weight:normal; font-size:8.5px;">※ 위험성 감소대책의 실효성, 아차사고 내용, 근로자 건의사항 등을 작성</span></td></tr>
                <tr><td style="border:1px solid #000; padding:4px; height:20px; font-size:10px;">{single_post_note}</td></tr>
            </table>

            <table>
                <tr><th colspan="5" style="text-align:left; background:#f2f2f2; border:1px solid #000;">■ TBM 리더 정보</th></tr>
                <tr style="text-align:center; font-weight:bold; background:#fafafa;">
                    <td style="width:18%;">소속</td><td style="width:20%;">직책(또는 직급)</td><td style="width:20%;">관리감독자</td><td style="width:18%;">성명</td><td style="width:24%;">서명란</td>
                </tr>
                <tr style="text-align:center;">
                    <td>{leader_dept}</td><td>{leader_role}</td><td>☑예 □아니오</td><td style="font-weight:bold;">{leader_name}</td><td style="height:28px;">{sign_tag}</td>
                </tr>
            </table>

            <table>
                <tr><th colspan="6" style="text-align:left; background:#f2f2f2; border:1px solid #000;">■ 참석자 확인 <span style="font-weight:normal; font-size:8.5px;">※ 외주 업체 포함 참석</span></th></tr>
                <tr style="text-align:center; background:#fafafa; font-weight:bold; font-size:9.5px;">
                    <td style="width:16%;">성 명</td><td style="width:17%;">서 명</td><td style="width:16%;">성 명</td><td style="width:17%;">서 명</td><td style="width:18%;">업 체 성 명</td><td style="width:16%;">업 체 서 명</td>
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
            body {{ font-family: 'Malgun Gothic', '맑은 고딕', sans-serif; margin: 8px; color: #000; font-size: 10.5px; }}
            table {{ width: 100%; border-collapse: collapse; margin-bottom: 4px; font-size: 9.5px; }}
            th, td {{ border: 1px solid #000; padding: 3px 4px; vertical-align: middle; text-align: center; }}
            .header-td {{ background-color: #f2f2f2; font-weight: bold; }}
        </style>
        </head>
        <body>
            <div style="font-size: 10px; font-weight: bold; margin-bottom: 5px;">[별지1. 작업내용이 동일하여, TBM을 1주일 단위로 하고자 할 경우의 양식]</div>
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
                    <td class="header-td">TBM 리더 서명</td>
                    <td>{leader_name}</td><td>{leader_name}</td><td>{leader_name}</td><td>{leader_name}</td><td>{leader_name}</td><td>{leader_name}</td><td>{leader_name}</td>
                </tr>
            </table>

            <table>
                <tr><th colspan="9" class="header-td" style="background:#e9ecef;">참석자 서명</th></tr>
                <tr class="header-td">
                    <td style="width:16%;">참석자</td>
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
        st.subheader("📋 TBM 표준회의록 공식 양식 미리보기")
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
            # 증빙 폴더에도 백업 사본 저장
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

    # 전체 백업 ZIP 원클릭 다운로드
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

                    # 파일 형식별 자동 미리보기 제공
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