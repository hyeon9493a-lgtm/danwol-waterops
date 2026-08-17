import streamlit as st
import json
import os

# 승인 유저 저장 파일 경로
AUTH_DB_FILE = "user_auth_db.json"

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
    
    # Streamlit Secrets 또는 기본 마스터 관리자 비밀번호
    admin_master_pw = st.secrets.get("ADMIN_PASSWORD", "danwol360!")
    # 허용된 공용 화이트리스트 패스코드 (Secrets 등록 가능)
    whitelist_codes = st.secrets.get("WHITELIST_CODES", ["DW-PASS-2026", "WATER-ADMIN"])

    st.markdown("<h2 style='text-align: center;'>💧 DANWOL AI-WaterOps 360 보안 접속</h2>", unsafe_allow_html=True)
    tab_login, tab_request = st.tabs(["🔒 시스템 로그인", "📝 신규 사용자 승인 요청"])

    # 탭 1: 로그인
    with tab_login:
        login_type = st.radio("접속 유형 선택", ["일반 사용자 (승인 계정 / 접속 코드)", "시스템 관리자 (승인 대시보드)"], horizontal=True)
        
        if login_type == "시스템 관리자 (승인 대시보드)":
            admin_pw = st.text_input("관리자 마스터 비밀번호", type="password", key="admin_pw_input")
            if st.button("관리자 모드로 접속", use_container_width=True):
                if admin_pw == admin_master_pw:
                    st.session_state.logged_in = True
                    st.session_state.user_role = "admin"
                    st.rerun()
                else:
                    st.error("관리자 비밀번호가 일치하지 않습니다.")
        else:
            auth_method = st.selectbox("인증 방식", ["승인된 계정으로 로그인", "승인 접속 코드 입력"])
            if auth_method == "승인된 계정으로 로그인":
                user_id = st.text_input("사번 또는 아이디", key="user_id_input")
                user_pw = st.text_input("비밀번호", type="password", key="user_pw_input")
                if st.button("로그인", use_container_width=True):
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
                                st.warning("현재 관리자 승인 대기 중인 계정입니다. 관리자 승인 후 접속 가능합니다.")
                        else:
                            st.error("비밀번호가 올바르지 않습니다.")
                    else:
                        st.error("등록되지 않은 사용자입니다. 승인 요청 탭에서 계정을 신청하세요.")
            else:
                passcode = st.text_input("부여받은 승인 접속 코드", type="password", key="passcode_input")
                if st.button("인증 코드로 접속", use_container_width=True):
                    if passcode in whitelist_codes:
                        st.session_state.logged_in = True
                        st.session_state.user_role = "user"
                        st.session_state.user_name = "임시 인증 사용자"
                        st.rerun()
                    else:
                        st.error("유효하지 않은 승인 접속 코드입니다.")

    # 탭 2: 신규 사용자 승인 요청
    with tab_request:
        st.caption("신청 정보를 제출하면 관리자 승인 후 메인 관제 페이지에 접근할 수 있습니다.")
        req_id = st.text_input("신청 사번/아이디 (영문/숫자)")
        req_name = st.text_input("신청자 성명")
        req_dept = st.text_input("소속/부서")
        req_pw = st.text_input("사용할 비밀번호 설정", type="password")
        
        if st.button("승인 요청 제출", use_container_width=True):
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
                        "status": "pending"  # 대기 상태
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
        for u_id, u_info in list(approved_users.items()):
            st.write(f"- {u_info.get('name')} ({u_id})")
            if st.button("권한 회수", key=f"rev_{u_id}"):
                del users[u_id]
                save_auth_db(auth_db)
                st.rerun()
