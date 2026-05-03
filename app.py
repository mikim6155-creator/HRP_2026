import streamlit as st
import pandas as pd
import os
from datetime import datetime

# [1] 페이지 설정
st.set_page_config(page_title="2026년 희망리턴패키지 온라인교육", page_icon="🎓", layout="wide")

# [2] 상태 초기화 (중복 저장 방지용)
if 'submitted' not in st.session_state:
    st.session_state.submitted = False

# [3] CSS 고정: 버튼 스타일 및 팝업 'X' 버튼 차단
st.markdown("""
    <style>
    /* 사이드바 제출 버튼 (파란색) */
    div.stSidebar div.stButton > button[kind="primary"] {
        background-color: #007BFF !important;
        color: white !important;
        font-size: 22px !important;
        font-weight: bold !important;
        height: 4rem !important;
        width: 100% !important;
        border-radius: 10px !important;
    }

    /* 소상공인지식배움터 바로가기 (하늘색) */
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

    /* 가이드 다운로드 (연한 딸기우유색) */
    [data-testid="stDownloadButton"] button {
        background-color: #FFD1DC !important;
        color: #555555 !important;
        font-weight: bold !important;
        border: none !important;
        height: 3.5rem !important;
        width: 100% !important;
        border-radius: 8px !important;
    }

    /* 팝업창 'X' 버튼 숨기기 (강제 제출 유도) */
    button[aria-label="Close"] {
        display: none !important;
    }

    /* 비활성화 체크박스 글자색 검정 고정 */
    div[data-testid="stCheckbox"] label[data-disabled="true"] p {
        color: #31333F !important;
        opacity: 1 !important;
    }
    </style>
    """, unsafe_allow_html=True)

# [4] 이수 기준 및 경로 설정
TARGET_TOTAL_HOURS, TARGET_MANDATORY_HOURS = 17.0, 3.5
TARGET_COMMON_HOURS, TARGET_SPECIAL_HOURS = 6.5, 7.0

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, "강의목록.xlsx")
PDF_PATH = os.path.join(BASE_DIR, "guide.pdf")
SAVE_PATH = os.path.join(BASE_DIR, "수강신청현황.csv")

# [5] 최종 확인 팝업창 (자동 저장 로직)
@st.dialog("📝 최종 수강 신청 완료", width="large")
def show_confirmation_dialog(data, user_info):
    if not st.session_state.submitted:
        save_df = data.copy()
        save_df['제출시간'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        save_df['성함'], save_df['업체명'] = user_info['name'], user_info['biz']
        save_df['연락처'], save_df['이메일'] = user_info['phone'], user_info['email']
        save_df['구분'] = user_info['category']
        
        save_cols = ['제출시간', '성함', '업체명', '연락처', '이메일', '구분', '표준대분류', '중분류', '강의명', '시간']
        save_df = save_df[save_cols]

        if not os.path.exists(SAVE_PATH):
            save_df.to_csv(SAVE_PATH, index=False, encoding='utf-8-sig')
        else:
            save_df.to_csv(SAVE_PATH, mode='a', header=False, index=False, encoding='utf-8-sig')
        
        st.session_state.submitted = True
        st.balloons()

    st.success("🎉 신청 정보가 서버에 자동으로 기록되었습니다.")
    st.warning("📸 **[필독] 아래 수강목록을 반드시 촬영하거나 캡쳐해 주세요!**")
    
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
    st.error("⚠️ 촬영을 완료하셨다면 브라우저 창(탭)을 닫아주세요.")

# [6] 메인 로직
if os.path.exists(EXCEL_PATH):
    @st.cache_data
    def load_data(path):
        df = pd.read_excel(path)
        return df.rename(columns={"과정명": "강의명"}) if "과정명" in df.columns else df

    df_lectures = load_data(EXCEL_PATH)

    # 사이드바
    st.sidebar.header("👤 교육생 정보")
    u_info = {
        'name': st.sidebar.text_input("성함", placeholder="이름 입력"),
        'biz': st.sidebar.text_input("업체명", placeholder="사업장명 입력"),
        'category': st.sidebar.selectbox("구분", ["경영개선", "재창업"]),
        'phone': st.sidebar.text_input("전화번호", placeholder="010-0000-0000"),
        'email': st.sidebar.text_input("이메일", placeholder="example@mail.com")
    }

    # 시간 계산
    cur_common_h, cur_special_h = 0.0, 0.0
    sel_special_ind = None
    for i, row in df_lectures.iterrows():
        if "업종공통" in str(row['대분류']) and st.session_state.get(f"업종공통_{row['중분류']}_{row['강의명']}_{i}"):
            cur_common_h += float(row['시간'])
        if "업종특화" in str(row['대분류']) and st.session_state.get(f"업종특화_{row['중분류']}_{row['강의명']}_{i}"):
            cur_special_h += float(row['시간'])
            sel_special_ind = row['중분류']

    # 메인 안내사항 (연구원님 요청 문구 수정 완료)
    st.title("🎓 2026년 희망리턴패키지 실전 온라인교육 수강목록")
    with st.container(border=True):
        st.subheader("📢 필독! 안내사항")
        st.markdown(f"""
        먼저 **2026년 희망리턴패키지 재기사업화 최종 선정**되신 것을 진심으로 축하드립니다.  
        선정 되신 이후에는 반드시 **실전교육(24H)**을 수료 하셔야 합니다. (대면7H+온라인17H= 총 24H)  
        실전교육 미수료 시 **선정취소 처리**되는 점 유의하시기 바랍니다.

        1. 왼쪽 화면의 **교육생 정보**를 입력해 주세요. 모바일로 볼 경우 왼쪽 상단 "** >> **" 눌러주세요.
        2. 아래 **필수교육+업종공통+업종특화** 강의목록을 보고 수강 할 과목을 선택 해 주세요.
        3. **필수교육**은 무조건 수강을 하셔야 합니다. (4개 中 1개라도 미 수강시 수료처리 안됨)
        4. **업종공통(6.5H)** 및 **업종특화(7H)**는 기준 시간 충족 시 추가 선택이 제한됩니다.
        5. 왼쪽 총 수강 시간이 **"{TARGET_TOTAL_HOURS}H"**가 되어야 만 **제출버튼**이 생성됩니다.
        6. 제출 후 **"내 수강신청 현황"** 팝업창이 뜨면 **화면캡쳐**를 한 뒤 강의를 수강해주세요.
        """)
        
        c1, c2 = st.columns(2)
        with c1: st.link_button("📖 소상공인지식배움터 바로가기", "https://edu.sbiz.or.kr/", use_container_width=True)
        with c2:
            if os.path.exists(PDF_PATH):
                with open(PDF_PATH, "rb") as f:
                    st.download_button("📄 수강방법 PDF 가이드 다운로드", f, "guide.pdf", "application/pdf", use_container_width=True)

    # 강의 선택 리스트
    selected_data = []
    cats = {f"필수교육({TARGET_MANDATORY_HOURS}H)": "필수", f"업종공통({TARGET_COMMON_HOURS}H)": "업종공통", f"업종특화({TARGET_SPECIAL_HOURS}H)": "업종특화"}

    for d_label, kw in cats.items():
        st.header(f"📂 {d_label}")
        m_df = df_lectures[df_lectures["대분류"].str.contains(kw, na=False)]
        for s_label in m_df["중분류"].unique():
            with st.expander(f"➕ {s_label}", expanded=True):
                s_df = m_df[m_df["중분류"] == s_label]
                for i, row in s_df.iterrows():
                    std_cat = "필수교육" if "필수" in kw else ("업종공통" if "업종공통" in kw else "업종특화")
                    key = f"{std_cat}_{s_label}_{row['강의명']}_{i}"
                    is_m, is_d = (std_cat == "필수교육"), False
                    if std_cat == "업종공통" and not st.session_state.get(key) and cur_common_h >= TARGET_COMMON_HOURS: is_d = True
                    if std_cat == "업종특화" and not st.session_state.get(key):
                        if cur_special_h >= TARGET_SPECIAL_HOURS or (sel_special_ind and sel_special_ind != s_label): is_d = True
                    if st.session_state.submitted: is_d = True

                    chk = st.checkbox(f"{row['강의명']} ({row['시간']}H)", key=key, value=is_m, disabled=is_d)
                    if is_m and not chk and not st.session_state.submitted:
                        st.session_state[key] = True
                        st.rerun()
                    if chk:
                        temp = row.to_dict(); temp['표준대분류'] = std_cat
                        selected_data.append(temp)

    # 사이드바 현황판 (4단 고정)
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
    
    if st.session_state.submitted:
        st.sidebar.info("✅ 신청이 완료되었습니다.")
        st.sidebar.button("전송 완료", type="primary", disabled=True)
    elif valid:
        st.sidebar.success("✅ 모든 이수 요건 충족!")
        if st.sidebar.button("최종 수강목록 확인 및 제출", type="primary"):
            if u_info['name'] and u_info['biz'] and u_info['phone']:
                show_confirmation_dialog(r_df, u_info)
            else: st.sidebar.error("기본 정보를 입력해 주세요.")
    else: st.sidebar.warning("⚠️ 이수 요건 부족")
else:
    st.error("강의목록.xlsx 파일을 찾을 수 없습니다.")