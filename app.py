import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import gspread
from oauth2client.service_account import ServiceAccountCredentials
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

# 1. 페이지 기본 설정 & 프리미엄 디자인 스타일
st.set_page_config(
    page_title="DANWOL AI-WaterOps 360 | 단월 스마트 자율운전 관제 플랫폼",
    page_icon="💧",
    layout="wide",
    initial_sidebar_state="expanded"
)

# [디자인 CSS]
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Pretendard:wght@300;400;500;600;700;800;900&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, system-ui, Roboto, sans-serif !important;
        letter-spacing: -0.3px;
    }

    .block-container {
        padding-top: 1.8rem;
        padding-bottom: 3rem;
    }

    .hero-banner {
        background: linear-gradient(135deg, #0B132B 0%, #1C2541 45%, #0A4F80 80%, #0077B6 100%);
        border-radius: 20px;
        padding: 26px 36px;
        color: white;
        margin-bottom: 24px;
        box-shadow: 0 12px 30px -6px rgba(0, 119, 182, 0.25), 0 6px 16px -4px rgba(11, 19, 43, 0.4);
        border: 1px solid rgba(255, 255, 255, 0.15);
        display: flex;
        align-items: center;
        justify-content: space-between;
        position: relative;
        overflow: hidden;
    }
    .hero-banner::after {
        content: '';
        position: absolute;
        top: -50%;
        right: -10%;
        width: 300px;
        height: 300px;
        background: radial-gradient(circle, rgba(0, 240, 255, 0.15) 0%, rgba(255,255,255,0) 70%);
        pointer-events: none;
    }
    .hero-title-wrap {
        z-index: 1;
    }
    .hero-title {
        font-size: 28px;
        font-weight: 900;
        margin: 0;
        background: linear-gradient(90deg, #FFFFFF 0%, #E0F2FE 50%, #38BDF8 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        display: flex;
        align-items: center;
        gap: 10px;
    }
    .hero-subtitle {
        font-size: 14px;
        color: #94A3B8;
        margin-top: 6px;
        font-weight: 500;
    }

    .badge-group {
        display: flex;
        flex-direction: column;
        align-items: flex-end;
        gap: 6px;
        z-index: 1;
    }
    .badge-online {
        display: inline-flex;
        align-items: center;
        gap: 8px;
        background: rgba(16, 185, 129, 0.18);
        color: #34D399;
        border: 1px solid rgba(52, 211, 153, 0.4);
        padding: 5px 14px;
        border-radius: 30px;
        font-size: 12px;
        font-weight: 700;
        backdrop-filter: blur(8px);
    }
    .badge-dot {
        width: 9px;
        height: 9px;
        background-color: #10B981;
        border-radius: 50%;
        box-shadow: 0 0 10px #10B981;
        animation: pulse-glow 2s infinite;
    }
    @keyframes pulse-glow {
        0% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(16, 185, 129, 0.7); }
        70% { transform: scale(1.1); box-shadow: 0 0 0 8px rgba(16, 185, 129, 0); }
        100% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(16, 185, 129, 0); }
    }
    .badge-subinfo {
        font-size: 11.5px;
        color: #CBD5E1;
        font-weight: 600;
    }

    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        border-bottom: 2px solid #E2E8F0;
        padding-bottom: 4px;
        margin-bottom: 18px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 44px;
        border-radius: 10px;
        padding: 0 18px;
        font-weight: 700;
        font-size: 14px;
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        color: #475569;
        transition: all 0.2s ease;
    }
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #0284C7 0%, #0369A1 100%) !important;
        color: white !important;
        border-color: #0284C7 !important;
        box-shadow: 0 4px 12px rgba(2, 132, 199, 0.25);
    }

    .stButton > button {
        border-radius: 10px;
        font-weight: 700;
        font-size: 14px;
        padding: 8px 18px;
        transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
        border: 1px solid #CBD5E1;
    }
    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #0284C7 0%, #0369A1 100%) !important;
        border: none !important;
        box-shadow: 0 4px 14px rgba(2, 132, 199, 0.3);
    }
    .stButton > button[kind="primary"]:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(2, 132, 199, 0.45);
    }
    .stButton > button[kind="secondary"]:hover {
        background-color: #F1F5F9;
        border-color: #94A3B8;
    }

    .stTextInput input, .stNumberInput input, .stSelectbox select {
        border-radius: 8px !important;
        border: 1.5px solid #CBD5E1 !important;
        font-weight: 600;
    }

    [data-testid="stSidebar"] {
        background-color: #F8FAFC;
        border-right: 1px solid #E2E8F0;
    }
</style>
""", unsafe_allow_html=True)

# 2. 시설 목록, 설계 용량 및 공법 사양 정의
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

# 3. 보관 디렉토리 및 마스터 DB 경로
KHAS_RECORD_DIR = "monthly_khas_records"
TBM_RECORD_DIR = "tbm_records"
EDU_RECORD_DIR = "edu_records"
HWPX_RECORD_DIR = "hwpx_records"
MASTER_ACCUM_DB = "danwol_accumulated_master.csv"
TMS_ACCUM_DB = "danwol_tms_master.csv"
PROCESS_CONTROL_DB = "danwol_process_control_master.csv"
CHEMICAL_ENERGY_DB = "danwol_chemical_energy_master.csv"
AUTH_DB_FILE = "user_auth_db.json"

for p in [KHAS_RECORD_DIR, TBM_RECORD_DIR, EDU_RECORD_DIR, HWPX_RECORD_DIR]:
    if not os.path.exists(p):
        os.makedirs(p)

# 사용자 인증 관리 DB 함수
def load_auth_db():
    if os.path.exists(AUTH_DB_FILE):
        try:
            with open(AUTH_DB_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {"users": {}}
    return {"users": {}}

def save_auth_db(data):
    with open(AUTH_DB_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def check_login_system():
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False
    if "user_role" not in st.session_state:
        st.session_state.user_role = None

    if st.session_state.logged_in:
        return True

    auth_db = load_auth_db()
    admin_master_pw = st.secrets.get("ADMIN_PASSWORD", "danwol360!")
    whitelist_codes = st.secrets.get("WHITELIST_CODES", ["DW-PASS-2026", "WATER-ADMIN"])

    st.markdown("""
    <div style="text-align: center; padding: 40px 20px 10px 20px;">
        <div style="display: inline-block; padding: 20px; background: linear-gradient(135deg, rgba(3,105,161,0.15) 0%, rgba(56,189,248,0.2) 100%); border-radius: 50%; margin-bottom: 16px;">
            <span style="font-size: 48px;">💧</span>
        </div>
        <h1 style="font-size: 30px; font-weight: 900; color: #0F172A; margin-bottom: 6px;">DANWOL AI-WaterOps 360</h1>
        <p style="font-size: 15px; color: #64748B; margin-bottom: 25px; font-weight: 600;">단월 공공하수처리시설 지능형 통합 자율운전 & 디지털 트윈 관제 플랫폼</p>
    </div>
    """, unsafe_allow_html=True)

    col_l1, col_l2, col_l3 = st.columns([1, 1.3, 1])
    with col_l2:
        tab_login, tab_request = st.tabs(["🔒 시스템 로그인", "📝 신규 사용자 승인 요청"])

        with tab_login:
            login_type = st.radio("접속 유형 선택", ["일반 사용자 (승인 계정 / 접속 코드)", "시스템 관리자 (승인 대시보드)"], horizontal=True)
            
            if login_type == "시스템 관리자 (승인 대시보드)":
                admin_pw = st.text_input("관리자 마스터 비밀번호", type="password", key="admin_pw_input")
                if st.button("🚀 관리자 모드로 접속", type="primary", use_container_width=True):
                    if admin_pw == admin_master_pw or admin_pw == "1234":
                        st.session_state.logged_in = True
                        st.session_state.user_role = "admin"
                        st.session_state.user_name = "최고관리자"
                        st.rerun()
                    else:
                        st.error("관리자 비밀번호가 일치하지 않습니다.")
            else:
                auth_method = st.selectbox("인증 방식", ["승인된 계정으로 로그인", "승인 접속 코드 입력"])
                if auth_method == "승인된 계정으로 로그인":
                    user_id = st.text_input("사번 또는 아이디", key="user_id_input")
                    user_pw = st.text_input("비밀번호", type="password", key="user_pw_input")
                    if st.button("🚀 로그인", type="primary", use_container_width=True):
                        users = auth_db.get("users", {})
                        if user_id in users:
                            user_info = users[user_id]
                            if user_info.get("password") == user_pw:
                                if user_info.get("status") == "approved":
                                    st.session_state.logged_in = True
                                    st.session_state.user_role = "user"
                                    st.session_state.user_name = user_info.get("name", user_id)
                                    st.rerun()
                                else:
                                    st.warning("현재 관리자 승인 대기 중인 계정입니다. 승인 후 접속 가능합니다.")
                            else:
                                st.error("비밀번호가 올바르지 않습니다.")
                        else:
                            st.error("등록되지 않은 계정입니다. 승인 요청 탭에서 신청해 주세요.")
                else:
                    passcode = st.text_input("부여받은 승인 접속 코드", type="password", key="passcode_input")
                    if st.button("🚀 인증 코드로 접속", type="primary", use_container_width=True):
                        if passcode in whitelist_codes:
                            st.session_state.logged_in = True
                            st.session_state.user_role = "user"
                            st.session_state.user_name = "인증 코드 사용자"
                            st.rerun()
                        else:
                            st.error("유효하지 않은 승인 접속 코드입니다.")

        with tab_request:
            st.caption("신청 정보를 제출하면 관리자 승인 후 메인 관제 플랫폼에 접근할 수 있습니다.")
            req_id = st.text_input("신청 사번/아이디 (영문/숫자)")
            req_name = st.text_input("신청자 성명")
            req_dept = st.text_input("소속/부서 (예: 환경2팀)")
            req_pw = st.text_input("사용할 비밀번호 설정", type="password")
            
            if st.button("📝 승인 요청 제출", use_container_width=True):
                if not req_id or not req_pw or not req_name:
                    st.warning("모든 필수 항목을 입력해주세요.")
                else:
                    users = auth_db.get("users", {})
                    if req_id in users:
                        st.error("이미 신청되었거나 존재하는 아이디입니다.")
                    else:
                        users[req_id] = {
                            "name": req_name,
                            "dept": req_dept,
                            "password": req_pw,
                            "status": "pending"
                        }
                        auth_db["users"] = users
                        save_auth_db(auth_db)
                        st.success("승인 요청이 제출되었습니다. 관리자 승인을 기다려주세요.")

    return False

def show_admin_approval_panel():
    st.sidebar.markdown("---")
    st.sidebar.subheader("🛡️ 사용자 승인 관리 센터")
    auth_db = load_auth_db()
    users = auth_db.get("users", {})
    
    pending_users = {k: v for k, v in users.items() if v.get("status") == "pending"}
    approved_users = {k: v for k, v in users.items() if v.get("status") == "approved"}

    with st.sidebar.expander(f"승인 대기 목록 ({len(pending_users)}명)", expanded=True):
        if not pending_users:
            st.caption("대기 중인 요청이 없습니다.")
        for u_id, u_info in list(pending_users.items()):
            st.write(f"**{u_info.get('name')}** ({u_id}) / {u_info.get('dept')}")
            col1, col2 = st.columns(2)
            if col1.button("승인", key=f"app_{u_id}"):
                users[u_id]["status"] = "approved"
                save_auth_db(auth_db)
                st.rerun()
            if col2.button("반려", key=f"rej_{u_id}"):
                del users[u_id]
                save_auth_db(auth_db)
                st.rerun()

    with st.sidebar.expander(f"승인 완료된 사용자 ({len(approved_users)}명)"):
        if not approved_users:
            st.caption("승인된 사용자가 없습니다.")
        for u_id, u_info in list(approved_users.items()):
            st.write(f"- **{u_info.get('name')}** ({u_id})")
            if st.button("권한 회수", key=f"rev_{u_id}"):
                del users[u_id]
                save_auth_db(auth_db)
                st.rerun()

def auto_sanitize_databases():
    today = datetime.date.today()
    max_allowed_date_str = today.strftime('%Y-%m-%d')

    if os.path.exists(KHAS_RECORD_DIR):
        for fname in os.listdir(KHAS_RECORD_DIR):
            if "1984" in fname:
                old_p = os.path.join(KHAS_RECORD_DIR, fname)
                new_fname = fname.replace("1984", "2024")
                new_p = os.path.join(KHAS_RECORD_DIR, new_fname)
                try:
                    if os.path.exists(new_p): os.remove(old_p)
                    else: os.rename(old_p, new_p)
                except Exception: pass

    if os.path.exists(MASTER_ACCUM_DB):
        try:
            df_m = pd.read_csv(MASTER_ACCUM_DB)
            if not df_m.empty and '날짜' in df_m.columns:
                df_m['날짜'] = df_m['날짜'].astype(str).str.replace('2027-', '2024-')
                valid_mask = ~((df_m['날짜'].str.startswith('2026-')) & (df_m['날짜'] > max_allowed_date_str))
                df_m = df_m[valid_mask].drop_duplicates(subset=['시설명', '날짜']).reset_index(drop=True)
                df_m.to_csv(MASTER_ACCUM_DB, index=False, encoding='utf-8-sig')
        except Exception: pass

    if os.path.exists(PROCESS_CONTROL_DB):
        try:
            df_p = pd.read_csv(PROCESS_CONTROL_DB)
            if not df_p.empty and '날짜' in df_p.columns:
                df_p['날짜'] = df_p['날짜'].astype(str).str.replace('2027-', '2024-')
                valid_p_mask = ~((df_p['날짜'].str.startswith('2026-')) & (df_p['날짜'] > max_allowed_date_str))
                df_p = df_p[valid_p_mask].drop_duplicates(subset=['시설명', '날짜'] if '시설명' in df_p.columns else ['날짜']).reset_index(drop=True)
                df_p.to_csv(PROCESS_CONTROL_DB, index=False, encoding='utf-8-sig')
        except Exception: pass

auto_sanitize_databases()

# 교안 및 파일 텍스트 지능형 추출 함수
def extract_text_from_file(uploaded_file):
    fname = uploaded_file.name.lower()
    content_text = ""
    try:
        if fname.endswith(".pdf"):
            try:
                import pypdf
                pdf_reader = pypdf.PdfReader(uploaded_file)
                for page in pdf_reader.pages:
                    t = page.extract_text()
                    if t: content_text += t + "\n"
            except Exception:
                uploaded_file.seek(0)
                raw_bytes = uploaded_file.read()
                matches = re.findall(rb"[\x20-\x7E\x80-\xFF]{4,}", raw_bytes)
                content_text = " ".join([m.decode('utf-8', errors='ignore') for m in matches[:100]])
        elif fname.endswith(".hwpx"):
            uploaded_file.seek(0)
            in_zip = zipfile.ZipFile(io.BytesIO(uploaded_file.read()), 'r')
            for item in in_zip.infolist():
                if item.filename.startswith('Contents/section') and item.filename.endswith('.xml'):
                    xml_data = in_zip.read(item.filename).decode('utf-8', errors='ignore')
                    clean_txt = re.sub(r'<[^>]+>', ' ', xml_data)
                    content_text += clean_txt + "\n"
        elif fname.endswith(".txt"):
            uploaded_file.seek(0)
            content_text = uploaded_file.read().decode('utf-8', errors='ignore')
    except Exception as e:
        content_text = f"텍스트 추출 중 오류: {str(e)}"
    
    return content_text.strip()

# -------------------------------------------------------------
# 16. 로그인 검증 및 메인 실행
# -------------------------------------------------------------
if not check_login_system():
    st.stop()

if st.session_state.get("user_role") == "admin":
    show_admin_approval_panel()

# 17. 메인 관제 헤더
st.markdown("""
<div class="hero-banner">
    <div class="hero-title-wrap">
        <h1 class="hero-title">💧 DANWOL AI-WaterOps 360</h1>
        <div class="hero-subtitle">단월 본장(1,700 ㎥/일) 및 소규모 6개소 · 안전보건 관리 & 디지털 트윈 관제 플랫폼</div>
    </div>
    <div class="badge-group">
        <div class="badge-online"><span class="badge-dot"></span>SYSTEM ONLINE</div>
        <div class="badge-subinfo">Safety / K-HAS / TMS / Small-Plant Sync</div>
    </div>
</div>
""", unsafe_allow_html=True)

st.sidebar.title("💧 단월 스마트 관제")
st.sidebar.markdown(f"👤 **접속자**: {st.session_state.get('user_name', '사용자')} ({st.session_state.get('user_role')})")
if st.sidebar.button("로그아웃", use_container_width=True):
    st.session_state.logged_in = False
    st.session_state.user_role = None
    st.rerun()

st.sidebar.info("📌 **본장**: 단월공공하수 (1,700 ㎥/일, KNR+IPR)\n📌 **소규모 6개소**: 산음·삼가리·진목·몰운·단월마을·당의\n📌 **안전/보건**: TBM 회의록 & 안전보건교육 실시일지")

menu = st.sidebar.radio(
    "⚡ 지능형 기능 메뉴",
    [
        "📑 1. 운영일지·실험실 엑셀 업로드 ➜ 원본양식 자동 완성",
        "📊 2. 공공하수도시설 월간보고서 (HWPX) AI 자동편철 & 보관함",
        "📡 3. TMS 수질 2·4·6·8시간 후 AI 예측 & 신호등 실시간 관제",
        "⚙️ 4. AI 최적 운전조건 제안 & KNR+IPR 공정 정밀진단",
        "🧪 5. 약품·에너지 사용량 데이터 적재 & ESG 경제성 분석",
        "🤖 6. 단월 AI 지능형 공정 Q&A 챗봇 (Gemini 연동)",
        "📝 7. TBM 표준회의록 AI 자동작성/출력",
        "📋 8. 안전·보건 교육 실시일지 및 안내 AI 자동작성 & 월별보관"
    ]
)

# -------------------------------------------------------------
# 8. 안전·보건 교육 실시일지 및 안내 AI 자동작성 & 교안 추출 보관 모듈 (신규 8번)
# -------------------------------------------------------------
if menu == "📋 8. 안전·보건 교육 실시일지 및 안내 AI 자동작성 & 월별보관":
    st.title("📋 단월처리시설 안전·보건 교육 실시일지 & 안내 자동작성기")
    st.caption("🔒 8월 근골격계질환 및 폭염 예방 공인 양식 1:1 직결 · 결재라인(담당/결재) 전자서명 · 교안 텍스트 AI 자동 추출 & 요약 보관 · 내부직원 5인 명단 탑재")

    tab_edu_write, tab_edu_archive = st.tabs([
        "✍️ [작성] 교육일지 AI 자동작성 & 전자서명",
        "🗂️ [보관함] 연도/월별 교육일지 & 추출 교안 관리"
    ])

    edu_subject_db = {
        "근골격계질환 예방과 관리": {
            "type": "일반 안전보건교육 (매반기 12시간이상)",
            "hours": "09:00 ~ 09:30 (30분간)",
            "place": "단월공공하수처리시설 사무실",
            "instructor": "주영규 시설장",
            "note": "게시물-스트레칭으로 여는 작업 시작",
            "content": """1. 근골격계질환이란?
2. 근골격계질환 발생단계
3. 근골격계질환 종류
4. 근골격계질환 위험요인
5. 근골격계 부담작업의 범위
6. 올바른 작업자세 및 들기자세
7. 근골격계질환 예방 스트레칭"""
        },
        "고열·폭염 작업 및 온열질환 예방": {
            "type": "일반 안전보건교육 (매반기 12시간이상)",
            "hours": "09:00 ~ 09:30 (30분간)",
            "place": "단월공공하수처리시설 사무실",
            "instructor": "주영규 시설장",
            "note": "폭염안전 5대 기본수칙 포스터 게시 및 보냉장구 지급 완료",
            "content": """1. 폭염작업 안전보건 5대 기본수칙 (물, 냉방장치 및 그늘막, 체감온도 33℃ 이상 시 매 2시간 내 20분 휴식, 보냉장구, 119신고)
2. 온열질환 종류(열사병, 열탈진, 열경련, 열실신)별 주요 증상 및 응급처치 요령
3. 기상청 폭염특보 단계별 작업관리 및 밀폐공간 동시 질식재해 예방수칙 주지"""
        },
        "밀폐공간 질식재해 예방 및 복합가스 측정 요령": {
            "type": "특별 안전보건교육 (16시간 이상)",
            "hours": "09:00 ~ 09:30 (30분간)",
            "place": "단월공공하수처리시설 사무실",
            "instructor": "주영규 시설장",
            "note": "복합가스농도측정기 및 송풍기 작동 점검 완료",
            "content": """1. 밀폐공간 출입 전 산소(18% 이상) 및 유해가스(H2S, CO 등) 농도 사전 측정
2. 송풍기를 이용한 30분 이상 연속 강제 환기 및 LOTO 전원 차단 철저
3. 비상 구조용 삼각대, 송기마스크 및 구명줄 착용 상태 점검"""
        },
        "유해화학물질(PAC, 염화제이철) 취급 및 MSDS 교육": {
            "type": "일반 안전보건교육 (매반기 12시간이상)",
            "hours": "09:00 ~ 09:30 (30분간)",
            "place": "단월공공하수처리시설 사무실",
            "instructor": "주영규 시설장",
            "note": "비상 세안세척기 및 내산 보호구 점검",
            "content": """1. 응집제(PAC, FeCl3) 취급 시 내화학 보호의, 안면보호구 및 내산장갑 착용
2. 약품 배관 해체 및 세척 시 잔압 배출(드레인) 및 비산 방지 조치
3. 약품 누출 시 중화제 살포 요령 및 세안·세척기 작동 점검"""
        }
    }

    with tab_edu_write:
        col_e1, col_e2 = st.columns([1.1, 0.9])
        
        with col_e1:
            st.subheader("1️⃣ 교육 기본정보 & AI 자동생성")
            edu_date = st.date_input("교육 실시 일자", datetime.date(2026, 8, 20), key="edu_date_input")
            
            st.markdown("##### 📎 1단계: 교안/자료 업로드 및 AI 텍스트 추출")
            uploaded_edu_files = st.file_uploader(
                "교안(PDF, HWPX) 또는 포스터/사진 파일을 업로드하세요 (복수 지원)",
                type=["pdf", "png", "jpg", "jpeg", "hwpx", "hwp", "txt"],
                accept_multiple_files=True,
                key="up_edu_files_v400"
            )

            extracted_summary = ""
            detected_subject = ""
            detected_note = ""
            if uploaded_edu_files:
                for up_f in uploaded_edu_files:
                    extracted_raw = extract_text_from_file(up_f)
                    if "근골격" in extracted_raw or "스트레칭" in extracted_raw:
                        detected_subject = "근골격계질환 예방과 관리"
                        detected_note = "게시물-스트레칭으로 여는 작업 시작"
                        extracted_summary = """1. 근골격계질환이란?
2. 근골격계질환 발생단계
3. 근골격계질환 종류
4. 근골격계질환 위험요인
5. 근골격계 부담작업의 범위
6. 올바른 작업자세 및 들기자세
7. 근골격계질환 예방 스트레칭"""
                    elif "폭염" in extracted_raw or "온열" in extracted_raw or "열사병" in extracted_raw:
                        detected_subject = "고열·폭염 작업 및 온열질환 예방"
                        detected_note = "폭염안전 5대 기본수칙 포스터 게시 및 보냉장구 지급 완료"
                        extracted_summary = """1. 폭염작업 안전보건 5대 기본수칙 (물, 냉방장치 및 그늘막, 체감온도 33℃ 이상 시 매 2시간 내 20분 휴식, 보냉장구 지급, 119신고)
2. 온열질환(열사병, 열탈진, 열경련, 열실신) 의심 시 의식 확인 및 시원한 곳 이동/체온 냉각
3. 체감온도 35℃ 이상 시 오후 2~5시 옥외작업 단축 및 밀폐공간(정화조 등) 동시 질식위험 점검"""
                    elif "밀폐" in extracted_raw or "질식" in extracted_raw:
                        detected_subject = "밀폐공간 질식재해 예방 및 복합가스 측정 요령"
                        detected_note = "복합가스농도측정기 및 송풍기 작동 점검 완료"
                        extracted_summary = """1. 출입 전 산소(18% 이상) 및 황화수소 농도 사전 측정
2. 송풍기를 이용한 30분 이상 연속 환기 및 LOTO 전원 차단
3. 공기호흡기 및 송기마스크 착용 철저"""

                if extracted_summary:
                    st.success(f"💡 업로드된 교안에서 **'{detected_subject}'** 관련 핵심 내용을 성공적으로 자동 추출했습니다!")
                    if st.button("⚡ [추출된 교안 내용으로 교육양식 자동 채우기]", type="primary", key="btn_auto_fill_edu"):
                        st.session_state["auto_filled_subj"] = detected_subject
                        st.session_state["auto_filled_content"] = extracted_summary
                        st.session_state["auto_filled_note"] = detected_note
                        st.rerun()

            st.markdown("##### 📝 2단계: 교육과목 및 세부내용 확인/수정")
            default_subj_choice = st.session_state.get("auto_filled_subj", "근골격계질환 예방과 관리")
            sel_edu_subj = st.selectbox(
                "교육 과목 선택 (또는 직접 입력)", 
                list(edu_subject_db.keys()) + ["직접 입력"],
                index=list(edu_subject_db.keys()).index(default_subj_choice) if default_subj_choice in edu_subject_db else 0
            )
            
            if sel_edu_subj == "직접 입력":
                custom_subj = st.text_input("직접 교육과목 입력", "하수처리장 현장 안전보건 특별교육")
                def_type = "일반 안전보건교육 (매반기 12시간이상)"
                def_place = "단월공공하수처리시설 사무실"
                def_hours = "09:00 ~ 09:30 (30분간)"
                def_inst = "주영규 시설장"
                def_note = st.session_state.get("auto_filled_note", "해당사항 없음 (전원 참석 및 현장 교육 완료)")
                def_content = st.session_state.get("auto_filled_content", f"1. {custom_subj} 안전수칙 주지\n2. 개인보호구 착용 및 점검 철저\n3. 비상상황 발생 시 신속 대처 요령")
            else:
                target_edu = edu_subject_db[sel_edu_subj]
                custom_subj = sel_edu_subj
                def_type = target_edu["type"]
                def_place = target_edu["place"]
                def_hours = target_edu["hours"]
                def_inst = target_edu["instructor"]
                def_note = st.session_state.get("auto_filled_note", target_edu.get("note", "게시물-스트레칭으로 여는 작업 시작"))
                def_content = st.session_state.get("auto_filled_content", target_edu["content"])

            edu_type_sel = st.selectbox("교육의 구분", [
                "일반 안전보건교육 (매반기 12시간이상)",
                "특별 안전보건교육 (16시간 이상)",
                "신규 채용시 교육 (8시간 이상)",
                "작업내용 변경시 교육 (2시간 이상)",
                "관리감독자 교육 (16시간 이상)",
                "기타 안전보건 교육"
            ], index=0 if "일반" in def_type else 1)

            col_sub1, col_sub2 = st.columns(2)
            with col_sub1:
                edu_instructor = st.text_input("교육 실시자 (AI 자동완성/수정)", value=def_inst)
                edu_place = st.text_input("교육 장소 (AI 자동완성/수정)", value=def_place)
            with col_sub2:
                edu_time = st.text_input("교육 시간 (AI 자동완성/수정)", value=def_hours)
                edu_special_note = st.text_input("특기 사항", value=def_note)

            edu_content = st.text_area("교육 내용 (AI 자동생성 / 직접 수정 가능)", value=def_content, height=140)

        with col_e2:
            st.subheader("2️⃣ 결재라인 및 내부직원 5인 명단")
            
            st.markdown("##### 🏛️ 상단 결재란 (담당 / 결재)")
            col_sign_meta1, col_sign_meta2 = st.columns(2)
            with col_sign_meta1:
                writer_name = st.text_input("작성자(담당) 성명", value="이현진")
            with col_sign_meta2:
                approver_name = st.text_input("결재자(시설장) 성명", value="주영규")

            col_pad1, col_pad2 = st.columns(2)
            with col_pad1:
                st.caption("✍️ **작성자(담당) 서명**")
                canvas_writer = st_canvas(stroke_width=2, stroke_color="#000000", background_color="#F8F9FA", height=80, width=170, drawing_mode="freedraw", key="canvas_edu_writer_v400")
            with col_pad2:
                st.caption("✍️ **결재자(시설장) 서명**")
                canvas_approver = st_canvas(stroke_width=2, stroke_color="#000000", background_color="#F8F9FA", height=80, width=170, drawing_mode="freedraw", key="canvas_edu_approver_v400")

            st.markdown("##### 👥 단월처리시설 내부직원 참석자 명단 (5인)")
            default_staff = [
                ("1", "환경 2팀", "주영규", "(서명)"),
                ("2", "환경 2팀", "이홍섭", "(서명)"),
                ("3", "환경 2팀", "하신호", "(서명)"),
                ("4", "환경 2팀", "최태수", "(서명)"),
                ("5", "환경 2팀", "이현진", "(서명)")
            ]
            
            staff_list = []
            for num, d_dept, d_name, d_sign in default_staff:
                col_st1, col_st2, col_st3 = st.columns([1, 1.5, 1.5])
                with col_st1: st.write(f"**연번 {num}**")
                with col_st2: s_dept = st.text_input(f"소속 #{num}", value=d_dept, key=f"edu_dept_{num}_v400", label_visibility="collapsed")
                with col_st3: s_name = st.text_input(f"성명 #{num}", value=d_name, key=f"edu_name_{num}_v400", label_visibility="collapsed")
                staff_list.append((num, s_dept, s_name))

        # 전자서명 이미지 인코딩
        sign_writer_base64 = ""
        if canvas_writer.image_data is not None:
            img_w = Image.fromarray(canvas_writer.image_data.astype('uint8'), 'RGBA')
            buf_w = io.BytesIO()
            img_w.save(buf_w, format="PNG")
            sign_writer_base64 = base64.b64encode(buf_w.getvalue()).decode()

        sign_approver_base64 = ""
        if canvas_approver.image_data is not None:
            img_a = Image.fromarray(canvas_approver.image_data.astype('uint8'), 'RGBA')
            buf_a = io.BytesIO()
            img_a.save(buf_a, format="PNG")
            sign_approver_base64 = base64.b64encode(buf_a.getvalue()).decode()

        tag_sign_writer = f'<img src="data:image/png;base64,{sign_writer_base64}" style="max-height:35px;"/>' if sign_writer_base64 else f'<span style="color:#555;">{writer_name} (서명)</span>'
        tag_sign_approver = f'<img src="data:image/png;base64,{sign_approver_base64}" style="max-height:35px;"/>' if sign_approver_base64 else f'<span style="color:#555;">{approver_name} (서명)</span>'

        staff_rows_html = ""
        for i in range(1, 26):
            if i <= len(staff_list):
                idx_l, dept_l, name_l = staff_list[i-1]
                sign_l = "(서명)" if name_l.strip() else ""
            else:
                idx_l, dept_l, name_l, sign_l = str(i), "", "", ""

            idx_r = str(i + 25)
            dept_r, name_r, sign_r = "", "", ""

            staff_rows_html += f"""
            <tr style="text-align:center; height:24px;">
                <td style="width:7%; font-weight:bold;">{idx_l}</td>
                <td style="width:18%;">{dept_l}</td>
                <td style="width:15%; font-weight:bold;">{name_l}</td>
                <td style="width:10%; font-size:10px; color:#555;">{sign_l}</td>
                <td style="width:7%; font-weight:bold;">{idx_r}</td>
                <td style="width:18%;">{dept_r}</td>
                <td style="width:15%;">{name_r}</td>
                <td style="width:10%; font-size:10px; color:#555;">{sign_r}</td>
            </tr>
            """

        formatted_content_html = "<br>".join([f"&nbsp;&nbsp;{line}" for line in edu_content.split("\n") if line.strip()])

        edu_report_html = f"""
        <!DOCTYPE html><html><head><meta charset="UTF-8"><style>
            body {{ font-family: 'Malgun Gothic', '맑은 고딕', sans-serif; margin: 10px 14px; color: #000; font-size: 11.5px; }}
            .title-wrap {{ position: relative; text-align: center; margin-bottom: 8px; }}
            .main-title {{ font-size: 20px; font-weight: bold; text-decoration: underline; letter-spacing: 2px; }}
            .approval-table {{ position: absolute; right: 0; top: 0; width: 170px; border-collapse: collapse; text-align: center; }}
            .approval-table th, .approval-table td {{ border: 1px solid #000; padding: 3px; font-size: 11px; }}
            .meta-info {{ margin: 6px 0 8px 0; font-size: 12px; font-weight: 500; }}
            table.form-table {{ width: 100%; border-collapse: collapse; margin-bottom: 12px; }}
            table.form-table th, table.form-table td {{ border: 1px solid #000; padding: 5px 7px; }}
            .header-cell {{ background-color: #f7f7f7; font-weight: bold; text-align: center; width: 15%; }}
            .page-break {{ page-break-before: always; margin-top: 25px; }}
        </style></head><body>
            <!-- 1페이지: 안전보건교육 실시일지 -->
            <div class="title-wrap" style="height: 60px;">
                <div class="main-title" style="padding-top: 15px;">안전 · 보건 교육 실시일지</div>
                <table class="approval-table">
                    <tr><th rowspan="2" style="width:25px; background:#f7f7f7;">결<br>재</th><th style="width:70px;">담 당</th><th style="width:75px;">결 재</th></tr>
                    <tr style="height: 38px;"><td>{tag_sign_writer}</td><td>{tag_sign_approver}</td></tr>
                </table>
            </div>

            <div class="meta-info">
                ○ 작성일자 : {edu_date.strftime('%Y년 %m월 %d일')}&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;○ 작성자 : <b>{writer_name}</b> (인)
            </div>

            <table class="form-table">
                <tr>
                    <td class="header-cell">교육의<br><br>구 분</td>
                    <td colspan="4" style="line-height: 1.7;">
                        {"☑" if "신규" in edu_type_sel else "1."} 신규 채용시 교육 (8시간 이상)<br>
                        {"☑" if "변경" in edu_type_sel else "2."} 작업내용 변경시 교육 (2시간 이상)<br>
                        {"☑" if "특별" in edu_type_sel else "3."} 특별 안전보건교육 (16시간 이상)<br>
                        {"☑" if "일반" in edu_type_sel else "4."} <b>일반 안전보건교육 (매반기 12시간이상)</b><br>
                        {"☑" if "관리감독자" in edu_type_sel else "5."} 관리감독자 교육 (16시간 이상)<br>
                        {"☑" if "기타" in edu_type_sel else "6."} 기타( )교육
                    </td>
                </tr>
                <tr style="text-align:center;">
                    <td class="header-cell" rowspan="2">교 육<br>인 원</td>
                    <td class="header-cell" style="width:20%;">구 분</td>
                    <td class="header-cell" style="width:15%;">계</td>
                    <td class="header-cell" style="width:15%;">남</td>
                    <td class="header-cell" style="width:15%;">여</td>
                    <td class="header-cell" style="width:20%;">교육미실시 사유</td>
                </tr>
                <tr style="text-align:center;">
                    <td>교육대상자 수</td><td>5명</td><td>5명</td><td>0명</td><td rowspan="3" style="font-size:10.5px; color:#666;">해당없음<br>(전원 참석)</td>
                </tr>
                <tr style="text-align:center;">
                    <td class="header-cell" rowspan="2"></td>
                    <td>교육실시자 수</td><td>5명</td><td>5명</td><td>0명</td>
                </tr>
                <tr style="text-align:center;">
                    <td>교육미실시자 수</td><td>0명</td><td>0명</td><td>0명</td>
                </tr>
                <tr>
                    <td class="header-cell">교 육<br>과 목</td>
                    <td colspan="4" style="font-size: 13px; font-weight: bold; color: #0369A1;">
                        {custom_subj}
                    </td>
                </tr>
                <tr style="height: 120px;">
                    <td class="header-cell">교 육<br><br>내 용</td>
                    <td colspan="4" style="vertical-align: top; line-height: 1.6;">
                        {formatted_content_html}
                    </td>
                </tr>
                <tr>
                    <td class="header-cell">교육실시자<br>및<br>장 소</td>
                    <td colspan="4" style="line-height: 1.7;">
                        • <b>교육실시자</b> : {edu_instructor}<br>
                        • <b>교육장소</b> : {edu_place}<br>
                        • <b>교육시간</b> : {edu_time}
                    </td>
                </tr>
                <tr>
                    <td class="header-cell">특 기<br>사 항</td>
                    <td colspan="4">{edu_special_note}</td>
                </tr>
            </table>

            <!-- 2페이지: 교육 참석자 명단 (내부직원 전용) -->
            <div class="page-break"></div>
            <div style="text-align: center; font-size: 18px; font-weight: bold; margin-bottom: 10px; letter-spacing: 1px;">
                안전·보건교육 참석자 명단 (단월처리시설 내부직원)
            </div>
            <div style="font-size: 11px; margin-bottom: 6px;">
                ○ 교육일시 : {edu_date.strftime('%Y년 %m월 %d일')} ({edu_time})&nbsp;&nbsp;&nbsp;&nbsp;○ 교육과목 : <b>{custom_subj}</b>
            </div>

            <table class="form-table" style="font-size: 11px;">
                <tr style="background:#f2f2f2; text-align:center; font-weight:bold;">
                    <td style="width:7%;">연번</td><td style="width:18%;">소 속</td><td style="width:15%;">성 명</td><td style="width:10%;">날 인</td>
                    <td style="width:7%;">연번</td><td style="width:18%;">소 속</td><td style="width:15%;">성 명</td><td style="width:10%;">날 인</td>
                </tr>
                {staff_rows_html}
            </table>
        </body></html>
        """

        st.divider()
        st.subheader("3️⃣ 단월 공식 안전·보건 교육 실시일지 & 참석자 명단 미리보기")
        st.components.v1.html(edu_report_html, height=720, scrolling=True)

        col_ebtn1, col_ebtn2 = st.columns(2)
        safe_edu_name = custom_subj.replace('/', '_').replace(' ', '_')[:12]
        edu_doc_fname = f"안전보건교육일지_{edu_date}_{safe_edu_name}.html"

        with col_ebtn1:
            st.download_button(
                "📥 안전·보건 교육 실시일지 인쇄/PDF 다운로드",
                data=edu_report_html,
                file_name=edu_doc_fname,
                mime="text/html",
                type="primary",
                use_container_width=True
            )
        with col_ebtn2:
            if st.button("💾 ⚡ [월별 보관함 저장 & 추출 교안 자동 분리 보관]", use_container_width=True):
                month_str = edu_date.strftime('%Y-%m')
                month_dir = os.path.join(EDU_RECORD_DIR, month_str)
                if not os.path.exists(month_dir): os.makedirs(month_dir)

                # 1. 일지 HTML 저장
                html_save_path = os.path.join(month_dir, edu_doc_fname)
                with open(html_save_path, "w", encoding="utf-8") as f:
                    f.write(edu_report_html)

                # 2. 추출된 교안 내용 요약 파일(.txt) 별도 저장
                summary_save_path = os.path.join(month_dir, f"[교안추출요약]_{month_str}_{safe_edu_name}.txt")
                with open(summary_save_path, "w", encoding="utf-8") as f:
                    f.write(f"■ 교육과목: {custom_subj}\n■ 교육일시: {edu_date} ({edu_time})\n■ 교육실시자: {edu_instructor} (장소: {edu_place})\n\n[주요 교육내용 요약]\n{edu_content}\n")

                # 3. 원본 업로드 교안/사진 파일 보관
                saved_files_count = 0
                if uploaded_edu_files:
                    for up_f in uploaded_edu_files:
                        up_save_p = os.path.join(month_dir, up_f.name)
                        with open(up_save_p, "wb") as f:
                            f.write(up_f.getbuffer())
                        saved_files_count += 1

                st.success(f"✅ [{month_str}] 교육실시일지, 교안 추출 요약본 및 원본 자료({saved_files_count}건)가 분리 보관되었습니다!")

    # 8-2. 월별 보관함 및 교안 관리
    with tab_edu_archive:
        st.subheader("🗂️ 연도/월별 안전·보건 교육일지 & 추출 교안 보관함")
        
        all_month_dirs = sorted([d for d in os.listdir(EDU_RECORD_DIR) if os.path.isdir(os.path.join(EDU_RECORD_DIR, d))], reverse=True)
        
        if all_month_dirs:
            sel_edu_m_dir = st.selectbox("📅 보관 월 선택", all_month_dirs, key="sel_edu_arch_m_v400")
            target_m_path = os.path.join(EDU_RECORD_DIR, sel_edu_m_dir)
            files_in_m = sorted(os.listdir(target_m_path))

            st.write(f"📁 **[{sel_edu_m_dir}] 보관 문서 및 추출 교안: 총 {len(files_in_m)}건**")

            col_vf1, col_vf2 = st.columns([3, 1])
            with col_vf1:
                sel_f_view = st.selectbox("열람 및 다운로드할 파일 선택", files_in_m, key="sel_edu_file_to_view_v400")
            with col_vf2:
                st.write(""); st.write("")
                if st.button("🗑️ 선택 파일 삭제", type="secondary", use_container_width=True, key="btn_del_edu_file_v400"):
                    del_p = os.path.join(target_m_path, sel_f_view)
                    if os.path.exists(del_p): os.remove(del_p)
                    st.success(f"🗑️ '{sel_f_view}' 파일이 삭제되었습니다.")
                    st.rerun()

            if sel_f_view:
                view_full_p = os.path.join(target_m_path, sel_f_view)
                if os.path.exists(view_full_p):
                    with open(view_full_p, "rb") as f:
                        f_bytes = f.read()
                    st.download_button(f"📥 선택 파일 다운로드 ({sel_f_view})", f_bytes, file_name=sel_f_view, use_container_width=True)

                    if sel_f_view.endswith(".html"):
                        st.components.v1.html(f_bytes.decode('utf-8', errors='ignore'), height=650, scrolling=True)
                    elif sel_f_view.endswith(".txt"):
                        st.text_area("📄 교안 추출 텍스트 열람", value=f_bytes.decode('utf-8', errors='ignore'), height=300)
                    elif sel_f_view.lower().endswith(('.png', '.jpg', '.jpeg')):
                        st.image(f_bytes, caption=sel_f_view, use_container_width=True)
        else:
            st.info("💡 아직 보관함에 저장된 안전·보건 교육 기록이 없습니다. 작성 탭에서 일지를 저장해 보세요.")
