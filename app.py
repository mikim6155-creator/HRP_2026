import streamlit as st
import pandas as pd
import os
from datetime import datetime

# [1] 페이지 설정
st.set_page_config(page_title="2026년 희망리턴패키지 온라인교육", page_icon="🎓", layout="wide")

# [2] 상태 초기화
if 'submitted' not in st.session_state:
    st.session_state.submitted = False
if 'trigger_popup' not in st.session_state:
    st.session_state.trigger_popup = False

# [3] CSS 설정 (버튼 색상 및 체크박스 스타일 고정)
st.markdown("""
    <style>
    div.stSidebar div.stButton > button[kind="primary"] {
        background-color: #007BFF !important;
        color: white !important;
        font-size: 22px !important;
        font-weight: bold !important;
        height: 4rem !important;
        width: 100% !important;
        border-radius: 10px !important;
    }
    [data-testid="stLinkButton"] a {
        background-color: #87CEEB !important;
        color: white !important;
        font-weight: bold !important;
        border: none !important;
        height: 3.5rem !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        border-radius: 8px !important;
        text-decoration: none !important;
    }
    [data-testid="stDownloadButton"] button {
        background-color: #FFD1DC !important;
        color: #555555 !important;
        font-weight: bold !important;
        border: none !important;
        height: 3.5rem !important;
        width: 100% !important;
        border-radius: 8px !important;
    }
    div[data-testid="stCheckbox"] label[data-disabled="true"] p {
        color: #31333F !important;
        opacity: 1 !important;
    }
    </style>
    """, unsafe_allow_html=True)

# [4] 경로 및 기준 시간 설정
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, "강의목록.xlsx")
PDF_PATH = os.path.join(BASE_DIR, "guide.pdf")
SAVE_PATH = os.path.join(BASE_DIR, "수강신청현황.csv")

TARGET_TOTAL_HOURS, TARGET_MANDATORY_HOURS = 17.0, 3.5
TARGET_COMMON_HOURS, TARGET_SPECIAL_HOURS = 6.5, 7.0

# [5] 최종 확인 팝업창 함수
@st.dialog("📝 최종 수강 신청 완료", width="large")
def show_confirmation_dialog(data, user_info):
    st.balloons()
    st.warning("📸 **[필독] 아래 수강목록을 반드시 촬영하거나 캡쳐해 주세요!**")
    st.write(f"### {user_info['name']}님, 최종 선택하신 수강 신청 현황입니다.")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.subheader("📂 필수교육")
        st.table(data[data['표준대분류'] == '필수교육'][['강의명']])
    with col2:
        st.subheader("📂 업종공통")
        st.table(data[data['표준대분류'] == '업종공통'][['강의명']])
    with col3:
        st.subheader("📂 업종특화")
        st.table(data[data['표준대분류'] == '업종특화'][['강의명']])

    st.divider()
    st.info("✅ 제출완료. 소상공인지식배움터에 접속하여 강의명 검색 후 수강하시기 바랍니다.")
    st.error("⚠️ 촬영을 완료하셨다면 우측 상단의 'X'를 눌러 팝업을 닫아주세요.")

# [6] 데이터 로드 및 메인 로직
if os.path.exists(EXCEL_PATH):
    @st.cache_data
    def load_data(path):
        df = pd.read_excel(path)
        return df.rename(columns={"과정명": "강의명"}) if "과정명" in df.columns else df
    df_lectures = load_data(EXCEL_PATH)

    # 사이드바: 교육생 정보
    st.sidebar.header("👤 교육생 정보")
    u_info = {
        'name': st.sidebar.text_input("성함", placeholder="이름 입력"),
        'biz': st.sidebar.text_input("업체명", placeholder="사업장명 입력"),
        'category': st.sidebar.selectbox("구분", ["경영개선", "재창업"]),
        'phone': st.sidebar.text_input("전화번호", placeholder="010-0000-0000"),
        'email': st.sidebar.text_input("이메일", placeholder="example@mail.com")
    }

    # 수강 현황 계산을 위한 데이터 처리
    selected_data = []
    cur_common_h, cur_special_h = 0.0, 0.0
    sel_special_ind = None
    
    # 시간 계산 루프
    for i, row in df_lectures.iterrows():
        std_cat = "필수교육" if "필수" in str(row['대분류']) else ("업종공통" if "업종공통" in str(row['대분류']) else "업종특화")
        key = f"{std_cat}_{row['중분류']}_{row['강의명']}_{i}"
        
        if st.session_state.get(key) or ("필수" in str(row['대분류'])):
            if "업종공통" in std_cat: cur_common_h += float(row['시간'])
            if "업종특화" in std_cat: 
                cur_special_h += float(row['시간'])
                sel_special_ind = row['중분류']

    # 메인 화면: 안내사항
    st.title("🎓 2026년 희망리턴패키지 실전 온라인교육 수강목록")
    with st.container(border=True):
        st.subheader("📢 필독! 안내사항")
        st.markdown(f"1. 왼쪽 화면의 **교육생 정보**를 입력해 주세요. 모바일로 볼 경우 왼쪽 상단 '** >> **' 눌러주세요.")
        st.link_button("📖 소상공인지식배움터 바로가기", "https://edu.sbiz.or.kr/", use_container_width=True)

    # 강의 리스트 (체크박스 영역)
    cats = {"필수교육": "필수", "업종공통": "업종공통", "업종특화": "업종특화"}
    for label, kw in cats.items():
        st.header(f"📂 {label}")
        m_df = df_lectures[df_lectures["대분류"].str.contains(kw, na=False)]
        for s_label in m_df["중분류"].unique():
            with st.expander(f"➕ {s_label}", expanded=True):
                s_df = m_df[m_df["중분류"] == s_label]
                for i, row in s_df.iterrows():
                    key = f"{label}_{s_label}_{row['강의명']}_{i}"
                    is_m = (label == "필수교육")
                    is_d = st.session_state.submitted
                    
                    if not is_m and not is_d:
                        if label == "업종공통" and not st.session_state.get(key) and cur_common_h >= TARGET_COMMON_HOURS: is_d = True
                        if label == "업종특화" and not st.session_state.get(key):
                            if cur_special_h >= TARGET_SPECIAL_HOURS or (sel_special_ind and sel_special_ind != s_label): is_d = True
                    
                    chk = st.checkbox(f"{row['강의명']} ({row['시간']}H)", key=key, value=is_m, disabled=is_d)
                    if chk:
                        temp = row.to_dict(); temp['표준대분류'] = label
                        selected_data.append(temp)

    # [💡 핵심 복구] 사이드바 실시간 수강 현황 (4단 지표 고정)
    r_df = pd.DataFrame(selected_data)
    t_h = r_df["시간"].sum() if not r_df.empty else 0.0
    m_h = r_df[r_df["표준대분류"] == "필수교육"]["시간"].sum() if not r_df.empty else 0.0
    c_h = r_df[r_df["표준대분류"] == "업종공통"]["시간"].sum() if not r_df.empty else 0.0
    s_h = r_df[r_df["표준대분류"] == "업종특화"]["시간"].sum() if not r_df.empty else 0.0

    st.sidebar.subheader("📊 실시간 수강 현황")
    st.sidebar.metric("총 수강 시간", f"{t_h:.1f} / {TARGET_TOTAL_HOURS}H")
    st.sidebar.metric("필수과목 시간", f"{m_h:.1f} / {TARGET_MANDATORY_HOURS}H")
    st.sidebar.metric("업종공통 시간", f"{c_h:.1f} / {TARGET_COMMON_HOURS}H")
    st.sidebar.metric("업종특화 시간", f"{s_h:.1f} / {TARGET_SPECIAL_HOURS}H")
    st.sidebar.markdown("---")

    valid = (t_h >= TARGET_TOTAL_HOURS and m_h >= TARGET_MANDATORY_HOURS and c_h >= TARGET_COMMON_HOURS and s_h >= TARGET_SPECIAL_HOURS)

    # 제출 로직 및 사이드바 상태 반영
    if st.session_state.submitted:
        st.sidebar.info("✅ 신청이 완료되었습니다.")
        st.sidebar.button("전송 완료", type="secondary", disabled=True)
        # 만약 제출 직후라면 팝업을 띄움
        if st.session_state.trigger_popup:
            st.session_state.trigger_popup = False
            show_confirmation_dialog(r_df, u_info)
    elif valid:
        st.sidebar.success("✅ 모든 이수 요건 충족!")
        if st.sidebar.button("최종 수강목록 확인 및 제출", type="primary"):
            if u_info['name'] and u_info['biz'] and u_info['phone']:
                # 데이터 저장
                save_df = r_df.copy()
                save_df['제출시간'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                save_df['성함'], save_df['업체명'], save_df['연락처'] = u_info['name'], u_info['biz'], u_info['phone']
                save_df['이메일'], save_df['구분'] = u_info['email'], u_info['category']
                
                if not os.path.exists(SAVE_PATH): save_df.to_csv(SAVE_PATH, index=False, encoding='utf-8-sig')
                else: save_df.to_csv(SAVE_PATH, mode='a', header=False, index=False, encoding='utf-8-sig')
                
                # 상태 변경 후 '새로고침'하여 사이드바를 먼저 업데이트하고, 그 다음 팝업을 띄움
                st.session_state.submitted = True
                st.session_state.trigger_popup = True
                st.rerun()
            else: st.sidebar.error("기본 정보를 입력해 주세요.")
    else: st.sidebar.warning("⚠️ 이수 요건 부족")
else:
    st.error("강의목록.xlsx 파일을 찾을 수 없습니다.")