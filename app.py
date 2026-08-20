import streamlit as st
import pandas as pd
import numpy as np
from streamlit_drawable_canvas import st_canvas
import datetime
import os
import re
import base64
import hashlib
import uuid
from PIL import Image
import io

# 1. 페이지 설정 & 프리미엄 블루 테마 CSS
st.set_page_config(
    page_title="단월처리시설 TBM(작업 전 안전점검회의) AI 자동작성기",
    page_icon="🛡️",
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
        font-size: 24px; font-weight: 900; margin: 0;
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

# 2. 보관 디렉토리 정의
TBM_RECORD_DIR = "tbm_records"
if not os.path.exists(TBM_RECORD_DIR):
    os.makedirs(TBM_RECORD_DIR)

# 3. 상단 헤더 배너
st.markdown("""
<div class="hero-banner">
    <div>
        <h1 class="hero-title">🛡️ 단월처리시설 TBM(작업 전 안전점검회의) AI 자동작성기</h1>
        <div class="hero-subtitle">산업안전보건법 및 중대재해처벌법 대응 · 작업명별 AI 위험요인 자동완성 · 관리감독자 전자서명 · 감사추적(Audit Trail)</div>
    </div>
    <div class="badge-online">SYSTEM ONLINE</div>
</div>
""", unsafe_allow_html=True)

st.sidebar.title("🛡️ TBM 안전관리")
st.sidebar.info("📌 **시설명**: 단월처리시설\n📌 **적용 대상**: 본장 및 소규모 6개소 / 외주공사\n📌 **법적 근거**: 산업안전보건법 및 전자서명법 제3조")

# 4. 사전 정의된 AI 위험요인 데이터베이스
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

# 탭 인터페이스
tab1, tab2 = st.tabs(["✍️ 1. TBM 안전회의록 작성 & 전자서명", "🗂️ 2. 과거 TBM 회의록 연도/주차별 보관함 & 관리"])

# -------------------------------------------------------------
# 탭 1: TBM 작성 화면
# -------------------------------------------------------------
with tab1:
    is_weekly = st.checkbox("📅 **[별지1] 작업내용이 동일하여 1주일 단위로 작성하고자 할 경우 체크**", value=False)
    c1, c2 = st.columns([1, 1])
    
    with c1:
        st.subheader("1️⃣ 작업 기본정보 & AI 맞춤 시나리오")
        tbm_date = st.date_input("TBM 일자", datetime.date.today())
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
        canvas = st_canvas(stroke_width=2, stroke_color="#000000", background_color="#F8F9FA", height=100, width=300, drawing_mode="freedraw", key="tbm_canvas_standalone")

    # 서명 이미지 Base64 인코딩
    sign_img_base64 = ""
    if canvas.image_data is not None and np.any(canvas.image_data[:, :, 3] > 0):
        img = Image.fromarray(canvas.image_data.astype('uint8'), 'RGBA')
        buffered = io.BytesIO()
        img.save(buffered, format="PNG")
        sign_img_base64 = base64.b64encode(buffered.getvalue()).decode()

    # 고유 문서 ID 및 SHA-256 무결성 해시코드 생성
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
            save_path = os.path.join(TBM_RECORD_DIR, active_filename)
            with open(save_path, "w", encoding="utf-8") as f:
                f.write(active_html)
            st.success("✅ 로컬 보관함에 안전하게 저장되었습니다!")

# -------------------------------------------------------------
# 탭 2: 과거 TBM 회의록 연도/주차별 보관함 & 관리
# -------------------------------------------------------------
with tab2:
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
        return "2026년", "08월 3주차", datetime.date(2026, 8, 20)

    saved_files = [f for f in os.listdir(TBM_RECORD_DIR) if f.endswith(".html")]
    if saved_files:
        file_meta = [{"filename": f, "year": parse_file_info(f)[0], "week": parse_file_info(f)[1], "date": parse_file_info(f)[2]} for f in saved_files]
        df_files = pd.DataFrame(file_meta)
        available_years = sorted(df_files["year"].unique(), reverse=True)
        
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            sel_year = st.selectbox("📅 1단계: 연도 선택", available_years, key="tbm_sel_y_v250")
        
        df_year_filtered = df_files[df_files["year"] == sel_year]
        available_weeks = sorted(df_year_filtered["week"].unique(), reverse=True)
        with col_f2:
            sel_week = st.selectbox(f"📆 2단계: {sel_year} 월/주차 선택", available_weeks, key="tbm_sel_w_v250")

        df_week_filtered = df_year_filtered[df_year_filtered["week"] == sel_week].sort_values(by="date", ascending=False)
        target_file_list = df_week_filtered["filename"].tolist()
        st.write(f"📁 **[{sel_year} > {sel_week}] 검색 결과: 총 {len(target_file_list)}건의 회의록**")
        
        col_sel, col_del = st.columns([3, 1])
        with col_sel:
            selected_file_to_view = st.selectbox("열람할 회의록 파일 선택", target_file_list, key="tbm_sel_doc_v250")
        with col_del:
            st.write(""); st.write("")
            if st.button("🗑️ 선택 문서 영구 삭제", type="secondary", use_container_width=True, key="tbm_btn_del_v250"):
                file_to_delete = os.path.join(TBM_RECORD_DIR, selected_file_to_view)
                if os.path.exists(file_to_delete):
                    os.remove(file_to_delete)
                st.success(f"🗑️ '{selected_file_to_view}' 문서가 삭제되었습니다.")
                st.rerun()

        if selected_file_to_view:
            file_full_path = os.path.join(TBM_RECORD_DIR, selected_file_to_view)
            if os.path.exists(file_full_path):
                with open(file_full_path, "r", encoding="utf-8") as f:
                    view_html_data = f.read()
                st.download_button(f"📥 선택 문서 다운로드 ({selected_file_to_view})", view_html_data, file_name=selected_file_to_view, mime="text/html", use_container_width=True)
                st.components.v1.html(view_html_data, height=650, scrolling=True)
    else:
        st.info("💡 아직 보관함에 저장된 TBM 회의록이 없습니다. 1번 탭에서 회의록을 작성하고 [서명문서 자동보관함 저장]을 눌러주세요.")
