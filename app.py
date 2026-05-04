"""
🍽️  맛집 정보 입력기 - Streamlit + Gemini API
웹 / 모바일 최적화 버전
"""

import streamlit as st
import google.generativeai as genai
import pandas as pd
import json
import os
from datetime import date

# ──────────────────────────────────────────────────────
# ✅ 설정 (Streamlit Cloud Secrets 사용)
# ──────────────────────────────────────────────────────
GEMINI_API_KEY  = st.secrets["GEMINI_API_KEY"]
ACCESS_PASSWORD = st.secrets["ACCESS_PASSWORD"]
EXCEL_FILE  = "맛집정보.xlsx"
SHEET_NAME  = "Sheet1"

# ──────────────────────────────────────────────────────
# 페이지 설정
# ──────────────────────────────────────────────────────
st.set_page_config(
    page_title="🍽️ 맛집 정보 입력기",
    page_icon="🍽️",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# ──────────────────────────────────────────────────────
# CSS - 웹/모바일 반응형
# ──────────────────────────────────────────────────────
st.markdown("""
<style>
    html, body, [class*="css"] { font-family: 'Noto Sans KR', sans-serif; }
    .block-container { padding: 1rem 0.8rem 3rem 0.8rem !important; max-width: 800px; }

    .main-title { font-size: 1.5rem; font-weight: 700; margin-bottom: 0; }
    .subtitle   { font-size: 0.9rem; opacity: 0.7; margin-bottom: 0.5rem; }

    .section-title {
        font-size: 1rem; font-weight: 700;
        border-left: 4px solid #ff6b35; padding-left: 10px;
        margin: 1.4rem 0 0.6rem;
    }

    /* 정보 박스 - 라이트/다크 자동 대응 */
    .info-box {
        border-radius: 12px;
        padding: 1.1rem 1rem;
        margin: 0.4rem 0;
        font-size: 0.95rem;
        line-height: 2.2;
        border: 1.5px solid rgba(128,128,128,0.3);
        background: rgba(128,128,128,0.08) !important;
    }

    /* 메뉴 카드 */
    .menu-card {
        border-radius: 14px;
        padding: 1.1rem 1rem;
        margin: 0.6rem 0;
        border: 1.5px solid rgba(128,128,128,0.3);
        background: rgba(128,128,128,0.08) !important;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    }
    .menu-label { font-size: 0.8rem; color: #ff6b35 !important; font-weight: 700; margin-bottom: 4px; }
    .menu-name  { font-size: 1.1rem; font-weight: 800; margin: 0.2rem 0 0.4rem; }
    .menu-sub   { font-size: 0.85rem; opacity: 0.7; margin-bottom: 0.5rem; }
    .menu-desc  { font-size: 0.9rem; line-height: 1.7; margin-bottom: 0.5rem; opacity: 0.9; }
    .menu-price { font-size: 1rem; font-weight: 800; color: #ff6b35 !important; }

    /* 배지 - 라이트/다크 자동 대응 */
    .badge {
        display: inline-block;
        border-radius: 20px; padding: 4px 12px;
        font-size: 0.82rem; font-weight: 600;
        margin: 3px 2px;
        background: rgba(255,152,0,0.2) !important;
        border: 1.5px solid rgba(255,152,0,0.5);
        color: #ff9800 !important;
    }
    .badge-blue  { background: rgba(33,150,243,0.2) !important; border-color: rgba(33,150,243,0.5) !important; color: #2196f3 !important; }
    .badge-green { background: rgba(76,175,80,0.2) !important;  border-color: rgba(76,175,80,0.5) !important;  color: #4caf50 !important; }
    .badge-red   { background: rgba(244,67,54,0.2) !important;  border-color: rgba(244,67,54,0.5) !important;  color: #f44336 !important; }

    .restaurant-title { font-size: 1.6rem; font-weight: 800; margin: 0.5rem 0 0.4rem; }

    /* 버튼 */
    .stButton > button {
        min-height: 52px !important;
        font-size: 1.05rem !important;
        font-weight: 600 !important;
        border-radius: 12px !important;
    }

    /* 입력창 */
    .stTextInput > div > div > input {
        font-size: 1.05rem !important;
        min-height: 52px !important;
        border-radius: 12px !important;
    }

    #MainMenu, footer, header { visibility: hidden; }
</style>
""", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────
# 로그인
# ──────────────────────────────────────────────────────
def check_login():
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False

    if not st.session_state.logged_in:
        st.markdown("## 🍽️ 맛집 정보 입력기")
        st.markdown("승인된 사용자만 접속 가능합니다.")
        pw = st.text_input("비밀번호", type="password", placeholder="비밀번호 입력",
                           label_visibility="collapsed")
        if st.button("로그인", use_container_width=True, type="primary"):
            if pw == ACCESS_PASSWORD:
                st.session_state.logged_in = True
                st.rerun()
            else:
                st.error("❌ 비밀번호가 올바르지 않습니다.")
        st.stop()


# ──────────────────────────────────────────────────────
# Gemini 초기화
# ──────────────────────────────────────────────────────
@st.cache_resource
def init_gemini():
    genai.configure(api_key=GEMINI_API_KEY)
    return genai.GenerativeModel("gemini-2.5-flash")


# ──────────────────────────────────────────────────────
# Gemini 프롬프트
# ──────────────────────────────────────────────────────
def query_restaurant(restaurant_name: str) -> dict:
    model = init_gemini()
    prompt = f"""
당신은 전 세계 맛집 전문 리서처입니다.
아래 식당에 대해 최대한 정확하게 조사하여 JSON으로만 답변하세요.
다른 텍스트, 마크다운, 설명 없이 오직 JSON만 출력하세요.
정보가 없는 항목은 반드시 빈 문자열("")로 처리하세요.

입력형식: "국가, 식당명" 또는 "식당명"
입력값: {restaurant_name}

[언어 표기 원칙]
- 한국 식당: 한글 작성 원칙. 상호/메뉴가 외국어인 경우에만 병기. 주소는 한글만.
- 해외 식당 (영어권): 한글/영어 병기
- 해외 식당 (비영어권): 한글/영어/현지어 병기

[추천사유 검증 기준]
- 미슐랭: 등급(3스타/2스타/1스타/빕구르망/셀렉티드) + 선정연도
- 블루리본/레드리본: 한국 식당만 해당
- 로컬인증: 해외 식당의 현지 맛집 인증 및 선정연도
- 노포: 창업연도 표기. 한국은 백년가게/오래가게 등록여부 확인
- 방송출연 필수 확인 목록:
  한국: 생활의달인, 수요미식회, 허영만(식객), 백종원(Street Food Fighter/랩소디 시리즈/기타),
        정용진회장, 전현무(전현무계획), 흑백요리사/한식대첩/마스터셰프코리아, 유튜버(더들리)
  해외: 한국방송 출연여부 + 현지방송 출연여부 모두 확인

[메뉴 가격 표기]
- 한국: 원화(원)
- 해외: 현지통화 + 한화 병기 (예: 120HKD / 약 21,000원)

다음 JSON 형식으로만 출력하세요:

{{
  "식당명": "한국식당:한글명 / 해외식당:한글명(영어명/현지어명)",
  "음식국적": "한식/일식/중화/양식/이태리/프랑스/태국/베트남 등 한글",
  "카테고리": "육류/해산물/면류/밥류/국탕/빵간식/기타 중 해당항목 한글",
  "기준메뉴": "Signature1의 주재료 한글 (예: 소고기, 돼지고기, 해수어, 밀가루 등)",

  "추천사유": {{
    "미슐랭": "등급+연도 (예: 빕구르망(2023년)) 또는 빈문자열",
    "블루리본": "블루리본 / 레드리본 또는 빈문자열 (한국식당만)",
    "로컬인증": "현지 맛집 인증명+선정연도 또는 빈문자열 (해외식당만)",
    "노포": "창업연도 (예: 1965년 창업) 또는 빈문자열",
    "백년가게": "Y 또는 빈문자열 (한국식당만)",
    "오래가게": "Y 또는 빈문자열 (한국식당만)",
    "방송_생활의달인": "Y 또는 빈문자열",
    "방송_수요미식회": "Y 또는 빈문자열",
    "방송_허영만": "Y 또는 빈문자열",
    "방송_백종원": "출연방송명 (예: Street Food Fighter 2019, 냉면랩소디) 또는 빈문자열",
    "방송_정용진": "Y 또는 빈문자열",
    "방송_전현무": "Y 또는 빈문자열",
    "방송_요리경연": "출연대회명 (예: 흑백요리사 시즌1) 또는 빈문자열",
    "방송_더들리": "Y 또는 빈문자열",
    "방송_기타": "기타 방송/유튜브 출연내용 또는 빈문자열"
  }},

  "signature1": {{
    "메뉴명": "한국:한글 / 해외:한글(영어/현지어)",
    "주재료": "주재료 한글",
    "요리방법": "요리방법 한글 (예: 구이, 찜, 볶음, 튀김, 국물 등)",
    "설명": "메뉴 설명 및 특징 한글 2~3문장",
    "가격": "한국:15,000원 / 해외:120HKD(약 21,000원) 형식"
  }},
  "signature2": {{
    "메뉴명": "한국:한글 / 해외:한글(영어/현지어)",
    "주재료": "주재료 한글",
    "요리방법": "요리방법 한글",
    "설명": "메뉴 설명 및 특징 한글 2~3문장",
    "가격": "현지통화(한화) 형식"
  }},
  "signature3": {{
    "메뉴명": "한국:한글 / 해외:한글(영어/현지어)",
    "주재료": "주재료 한글",
    "요리방법": "요리방법 한글",
    "설명": "메뉴 설명 및 특징 한글 2~3문장",
    "가격": "현지통화(한화) 형식"
  }},

  "국가": "한국:한국 / 해외:한글(영어/현지어)",
  "시도": "한국:시도 한글 / 해외:영어(현지어)",
  "지역": "한국:세부지역 한글 / 해외:영어(현지어) (예: 센트럴(中環))",
  "주소": "한국:한글주소만 / 해외:영어주소(현지어주소)",
  "MRT역": "역명+출구번호. 한국:한글 / 해외:한글(영어/현지어). 없으면 빈문자열",
  "MRT도보": "도보이동방법 한글. 건물/도로명은 영어/현지어 병기. 없으면 빈문자열",
  "전화번호": "한국:(지역번호)전화번호 / 해외:+국가번호 (지역번호) 전화번호",
  "예약여부": "필수예약 / 권장 / 불필요 + 예약방법 한글 설명",
  "대기피하는시간": "웨이팅 피하는 시간대 한글 설명",
  "팁및정보": "식당 관련 유용한 팁 및 이용정보 한글",
  "영업시간": "요일별 영업시간",
  "브레이크타임": "브레이크타임 (없으면 빈문자열)",
  "라스트오더": "라스트오더 시간 (없으면 빈문자열)",
  "정기휴무": "휴무일",
  "가격대": "~10K / 10K~20K / 20K~40K / 40K~60K / 60K~80K / 80K~ 중 해당",
  "Homepage": "공식 홈페이지 또는 SNS URL (없으면 빈문자열)"
}}
"""
    response = model.generate_content(prompt)
    raw = response.text.strip()
    if "```" in raw:
        parts = raw.split("```")
        for part in parts:
            part = part.strip()
            if part.startswith("json"):
                part = part[4:].strip()
            try:
                return json.loads(part)
            except:
                continue
    return json.loads(raw)


# ──────────────────────────────────────────────────────
# 엑셀 저장
# ──────────────────────────────────────────────────────
def append_to_excel(info: dict):
    sig1 = info.get("signature1", {})
    sig2 = info.get("signature2", {})
    sig3 = info.get("signature3", {})
    tr   = info.get("추천사유", {})

    def yn(val): return "*" if str(val).strip().upper() in ("Y","YES","*") else ""

    new_row = {
        "프렌차이즈":     "",
        "식당명":         info.get("식당명", ""),
        "국적":           info.get("음식국적", ""),
        "카테고리":       info.get("카테고리", ""),
        "기준메뉴":       info.get("기준메뉴", ""),
        "Signature메뉴1": sig1.get("메뉴명", ""),
        "Signature1설명": sig1.get("설명", ""),
        "Signature1가격": sig1.get("가격", ""),
        "Signature메뉴2": sig2.get("메뉴명", ""),
        "Signature2설명": sig2.get("설명", ""),
        "Signature2가격": sig2.get("가격", ""),
        "Signature메뉴3": sig3.get("메뉴명", ""),
        "Signature3설명": sig3.get("설명", ""),
        "Signature3가격": sig3.get("가격", ""),
        "국가":           info.get("국가", ""),
        "시도":           info.get("시도", ""),
        "지역":           info.get("지역", ""),
        "주소":           info.get("주소", ""),
        "MRT역":          info.get("MRT역", ""),
        "MRT도보":        info.get("MRT도보", ""),
        "전화번호":       info.get("전화번호", ""),
        "예약여부":       info.get("예약여부", ""),
        "대기피하는시간": info.get("대기피하는시간", ""),
        "팁및정보":       info.get("팁및정보", ""),
        "영업시간":       info.get("영업시간", ""),
        "브레이크타임":   info.get("브레이크타임", ""),
        "라스트오더":     info.get("라스트오더", ""),
        "정기휴무":       info.get("정기휴무", ""),
        "가격대":         info.get("가격대", ""),
        "노포":           tr.get("노포", ""),
        "백년가게":       yn(tr.get("백년가게","")),
        "오래가게":       yn(tr.get("오래가게","")),
        "미쉐린가이드":   tr.get("미슐랭", ""),
        "블루리본":       tr.get("블루리본",""),
        "로컬인증":       tr.get("로컬인증",""),
        "생활의달인":     yn(tr.get("방송_생활의달인","")),
        "수요미식회":     yn(tr.get("방송_수요미식회","")),
        "허영만":         yn(tr.get("방송_허영만","")),
        "백종원":         tr.get("방송_백종원",""),
        "정용진":         yn(tr.get("방송_정용진","")),
        "전현무":         yn(tr.get("방송_전현무","")),
        "요리경연":       tr.get("방송_요리경연",""),
        "더들리":         yn(tr.get("방송_더들리","")),
        "기타방송":       tr.get("방송_기타",""),
        "방문여부":       "N",
        "맛":             "",
        "식당시설":       "",
        "주차":           "",
        "Homepage":       info.get("Homepage", ""),
        "입력날짜":       date.today(),
    }

    if os.path.exists(EXCEL_FILE):
        df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    else:
        df = pd.DataFrame([new_row])

    df.to_excel(EXCEL_FILE, sheet_name=SHEET_NAME, index=False)


# ──────────────────────────────────────────────────────
# 결과 화면 (모바일 최적화)
# ──────────────────────────────────────────────────────
def show_result(info: dict):
    tr   = info.get("추천사유", {})
    sig1 = info.get("signature1", {})
    sig2 = info.get("signature2", {})
    sig3 = info.get("signature3", {})

    def yn(val): return str(val).strip().upper() in ("Y","YES","*")

    # 식당명 + 기본 배지
    st.markdown(f'<div class="restaurant-title">🍽️ {info.get("식당명","")}</div>',
                unsafe_allow_html=True)
    st.markdown(
        f'<span class="badge badge-blue">{info.get("음식국적","")}</span>'
        f'<span class="badge badge-blue">{info.get("카테고리","")}</span>'
        f'<span class="badge badge-blue">{info.get("가격대","")}</span>',
        unsafe_allow_html=True
    )
    st.divider()

    # 추천사유 배지
    badges_html = ""
    if tr.get("미슐랭"):
        badges_html += f'<span class="badge badge-red">⭐ {tr["미슐랭"]}</span> '
    if tr.get("블루리본"):
        badges_html += f'<span class="badge badge-green">🎗️ {tr["블루리본"]}</span> '
    if tr.get("로컬인증"):
        badges_html += f'<span class="badge badge-green">🏆 {tr["로컬인증"]}</span> '
    if tr.get("노포"):
        badges_html += f'<span class="badge">🏛️ {tr["노포"]}</span> '
    if yn(tr.get("백년가게","")):
        badges_html += '<span class="badge">🏅 백년가게</span> '
    if yn(tr.get("오래가게","")):
        badges_html += '<span class="badge">🏅 오래가게</span> '

    방송목록 = [
        ("생활의달인","방송_생활의달인"), ("수요미식회","방송_수요미식회"),
        ("허영만","방송_허영만"), ("정용진","방송_정용진"),
        ("전현무","방송_전현무"), ("더들리","방송_더들리"),
    ]
    for name, key in 방송목록:
        if yn(tr.get(key,"")):
            badges_html += f'<span class="badge badge-red">📺 {name}</span> '
    if tr.get("방송_백종원"):
        badges_html += f'<span class="badge badge-red">📺 백종원({tr["방송_백종원"]})</span> '
    if tr.get("방송_요리경연"):
        badges_html += f'<span class="badge badge-red">🏆 {tr["방송_요리경연"]}</span> '
    if tr.get("방송_기타"):
        badges_html += f'<span class="badge">📺 {tr["방송_기타"]}</span> '

    if badges_html:
        st.markdown('<p class="section-title">📌 추천사유</p>', unsafe_allow_html=True)
        st.markdown(f'<div style="line-height:2.2">{badges_html}</div>',
                    unsafe_allow_html=True)

    # 시그니처 메뉴 (세로 배치 - 모바일 최적)
    st.markdown('<p class="section-title">🍴 시그니처 메뉴</p>', unsafe_allow_html=True)
    for label, sig in [("Signature 1 ⭐", sig1), ("Signature 2", sig2), ("Signature 3", sig3)]:
        if sig.get("메뉴명"):
            st.markdown(
                f'<div class="menu-card">'
                f'<div class="menu-label">{label}</div>'
                f'<div class="menu-name">{sig["메뉴명"]}</div>'
                f'<div class="menu-sub">주재료: {sig.get("주재료","")} &nbsp;|&nbsp; 조리법: {sig.get("요리방법","")}</div>'
                f'<div class="menu-desc">{sig.get("설명","")}</div>'
                f'<div class="menu-price">💰 {sig.get("가격","")}</div>'
                f'</div>',
                unsafe_allow_html=True
            )

    # 위치 정보
    st.markdown('<p class="section-title">📍 위치 정보</p>', unsafe_allow_html=True)
    loc_html = (
        f'<div class="info-box">'
        f'🌏 <b>국가:</b> {info.get("국가","")}<br>'
        f'📌 <b>시/도:</b> {info.get("시도","")}<br>'
        f'📍 <b>지역:</b> {info.get("지역","")}<br>'
        f'🏠 <b>주소:</b> {info.get("주소","")}<br>'
    )
    if info.get("MRT역"):
        loc_html += f'🚇 <b>MRT/지하철:</b> {info.get("MRT역","")}<br>'
    if info.get("MRT도보"):
        loc_html += f'🚶 <b>도보이동:</b> {info.get("MRT도보","")}<br>'
    loc_html += '</div>'
    st.markdown(loc_html, unsafe_allow_html=True)

    # 운영 정보
    st.markdown('<p class="section-title">📋 운영 정보</p>', unsafe_allow_html=True)
    ops_html = (
        f'<div class="info-box">'
        f'📞 <b>전화:</b> {info.get("전화번호","")}<br>'
        f'🕐 <b>영업시간:</b> {info.get("영업시간","")}<br>'
    )
    if info.get("브레이크타임"):
        ops_html += f'☕ <b>브레이크타임:</b> {info.get("브레이크타임","")}<br>'
    if info.get("라스트오더"):
        ops_html += f'⏱️ <b>라스트오더:</b> {info.get("라스트오더","")}<br>'
    ops_html += (
        f'🚫 <b>정기휴무:</b> {info.get("정기휴무","")}<br>'
        f'📅 <b>예약:</b> {info.get("예약여부","")}<br>'
    )
    if info.get("대기피하는시간"):
        ops_html += f'⏰ <b>대기 피하는 시간:</b> {info.get("대기피하는시간","")}<br>'
    ops_html += '</div>'
    st.markdown(ops_html, unsafe_allow_html=True)

    # 팁 & 홈페이지
    if info.get("팁및정보"):
        st.markdown('<p class="section-title">💡 팁 & 정보</p>', unsafe_allow_html=True)
        st.info(info.get("팁및정보",""))
    if info.get("Homepage"):
        st.markdown(f'🌐 **홈페이지:** [{info["Homepage"]}]({info["Homepage"]})')


# ──────────────────────────────────────────────────────
# 메인
# ──────────────────────────────────────────────────────
def main():
    check_login()

    # 헤더
    col_title, col_logout = st.columns([5, 1])
    with col_title:
        st.markdown('<p class="main-title">🍽️ 맛집 정보 입력기</p>', unsafe_allow_html=True)
        st.markdown('<p class="subtitle">Gemini AI가 자동으로 맛집 정보를 조사합니다</p>',
                    unsafe_allow_html=True)
    with col_logout:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("로그아웃", use_container_width=True):
            st.session_state.logged_in = False
            st.rerun()

    st.divider()

    # 검색창
    restaurant_name = st.text_input(
        "식당명",
        placeholder="예) 진주회관  /  홍콩, Tim Ho Wan",
        label_visibility="collapsed"
    )
    search_btn = st.button("🔍 정보 조회", use_container_width=True, type="primary")

    if search_btn and not restaurant_name:
        st.warning("⚠️ 식당명을 입력해주세요.")

    if search_btn and restaurant_name:
        with st.spinner(f"'{restaurant_name}' 정보 조사 중... (20~40초 소요)"):
            try:
                info = query_restaurant(restaurant_name)
                st.session_state["last_result"] = info
            except json.JSONDecodeError:
                st.error("❌ 응답 파싱 오류입니다. 다시 시도해주세요.")
                st.stop()
            except Exception as e:
                st.error(f"❌ 오류 발생: {e}")
                st.stop()

    # 결과
    if "last_result" in st.session_state:
        info = st.session_state["last_result"]
        st.divider()
        show_result(info)
        st.divider()

        if st.button("💾 엑셀에 저장", type="primary", use_container_width=True):
            try:
                append_to_excel(info)
                st.success(f"✅ '{info.get('식당명','')}' 저장 완료!")
                if os.path.exists(EXCEL_FILE):
                    with open(EXCEL_FILE, "rb") as f:
                        st.download_button(
                            "📥 엑셀 파일 다운로드",
                            data=f,
                            file_name=EXCEL_FILE,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
            except Exception as e:
                st.error(f"저장 오류: {e}")


if __name__ == "__main__":
    main()
