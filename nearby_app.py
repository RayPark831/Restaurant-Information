"""
🍽️  맛집 탐색기 - 내 주변 맛집 검색 + 상세정보
Google Places API + Gemini API + Streamlit
웹/모바일 최적화
"""

import streamlit as st
import google.generativeai as genai
import pandas as pd
import json
import os
import requests
import math
from datetime import date

# ──────────────────────────────────────────────────────
# ✅ 설정
# ──────────────────────────────────────────────────────
GEMINI_API_KEY   = st.secrets["GEMINI_API_KEY"]
ACCESS_PASSWORD  = st.secrets["ACCESS_PASSWORD"]
PLACES_API_KEY   = st.secrets["PLACES_API_KEY"]
EXCEL_FILE  = "맛집정보.xlsx"
SHEET_NAME  = "Sheet1"

# ──────────────────────────────────────────────────────
# 페이지 설정
# ──────────────────────────────────────────────────────
st.set_page_config(
    page_title="🍽️ 내 주변 맛집",
    page_icon="🍽️",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# ──────────────────────────────────────────────────────
# CSS
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

    .info-box {
        border-radius: 12px; padding: 1.1rem 1rem; margin: 0.4rem 0;
        font-size: 0.95rem; line-height: 2.2;
        border: 1.5px solid rgba(128,128,128,0.3);
        background: rgba(128,128,128,0.08) !important;
    }

    /* 식당 리스트 카드 */
    .rest-card {
        border-radius: 14px; padding: 1rem;
        margin: 0.6rem 0;
        border: 1.5px solid rgba(128,128,128,0.3);
        background: rgba(128,128,128,0.08) !important;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    }
    .rest-name   { font-size: 1.05rem; font-weight: 800; margin-bottom: 0.3rem; }
    .rest-dist   { font-size: 0.82rem; color: #ff6b35 !important; font-weight: 600; }
    .rest-reason { font-size: 0.85rem; opacity: 0.8; margin: 0.3rem 0; }
    .rest-menu   { font-size: 0.85rem; opacity: 0.9; line-height: 1.8; }

    /* 메뉴 카드 */
    .menu-card {
        border-radius: 14px; padding: 1.1rem 1rem; margin: 0.6rem 0;
        border: 1.5px solid rgba(128,128,128,0.3);
        background: rgba(128,128,128,0.08) !important;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    }
    .menu-label { font-size: 0.8rem; color: #ff6b35 !important; font-weight: 700; margin-bottom: 4px; }
    .menu-name  { font-size: 1.1rem; font-weight: 800; margin: 0.2rem 0 0.4rem; }
    .menu-sub   { font-size: 0.85rem; opacity: 0.7; margin-bottom: 0.5rem; }
    .menu-desc  { font-size: 0.9rem; line-height: 1.7; margin-bottom: 0.5rem; opacity: 0.9; }
    .menu-price { font-size: 1rem; font-weight: 800; color: #ff6b35 !important; }

    .badge {
        display: inline-block; border-radius: 20px; padding: 3px 10px;
        font-size: 0.78rem; font-weight: 600; margin: 2px;
        background: rgba(255,152,0,0.2) !important;
        border: 1.5px solid rgba(255,152,0,0.5); color: #ff9800 !important;
    }
    .badge-blue  { background: rgba(33,150,243,0.2) !important; border-color: rgba(33,150,243,0.5) !important; color: #2196f3 !important; }
    .badge-green { background: rgba(76,175,80,0.2) !important;  border-color: rgba(76,175,80,0.5) !important;  color: #4caf50 !important; }
    .badge-red   { background: rgba(244,67,54,0.2) !important;  border-color: rgba(244,67,54,0.5) !important;  color: #f44336 !important; }

    .restaurant-title { font-size: 1.6rem; font-weight: 800; margin: 0.5rem 0 0.4rem; }

    .stButton > button {
        min-height: 52px !important; font-size: 1.05rem !important;
        font-weight: 600 !important; border-radius: 12px !important;
    }
    .stTextInput > div > div > input {
        font-size: 1.05rem !important; min-height: 52px !important;
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
        st.markdown("## 🍽️ 내 주변 맛집 탐색기")
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
# 거리 계산 (Haversine)
# ──────────────────────────────────────────────────────
def calc_distance(lat1, lng1, lat2, lng2):
    R = 6371
    d_lat = math.radians(lat2 - lat1)
    d_lng = math.radians(lng2 - lng1)
    a = math.sin(d_lat/2)**2 + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.sin(d_lng/2)**2
    return R * 2 * math.asin(math.sqrt(a))


# ──────────────────────────────────────────────────────
# Google Places API - 반경 내 식당 검색
# ──────────────────────────────────────────────────────
def search_nearby_restaurants(lat, lng, radius_km, keyword=""):
    radius_m = int(radius_km * 1000)
    url = "https://maps.googleapis.com/maps/api/place/nearbysearch/json"
    params = {
        "location": f"{lat},{lng}",
        "radius": radius_m,
        "type": "restaurant",
        "key": PLACES_API_KEY,
        "language": "ko",
    }
    if keyword:
        params["keyword"] = keyword

    results = []
    while True:
        resp = requests.get(url, params=params).json()
        results.extend(resp.get("results", []))
        next_token = resp.get("next_page_token")
        if not next_token or len(results) >= 60:
            break
        import time; time.sleep(2)
        params = {"pagetoken": next_token, "key": PLACES_API_KEY, "language": "ko"}

    # 거리 계산 후 정렬
    restaurants = []
    for r in results:
        loc = r["geometry"]["location"]
        dist = calc_distance(lat, lng, loc["lat"], loc["lng"])
        restaurants.append({
            "place_id": r.get("place_id",""),
            "name": r.get("name",""),
            "address": r.get("vicinity",""),
            "rating": r.get("rating", 0),
            "user_ratings": r.get("user_ratings_total", 0),
            "lat": loc["lat"],
            "lng": loc["lng"],
            "distance": dist,
        })

    return sorted(restaurants, key=lambda x: x["distance"])


# ──────────────────────────────────────────────────────
# Gemini - 1단계: 간단 리스트 정보
# ──────────────────────────────────────────────────────
def get_simple_info(restaurant_name: str, address: str, distance: float) -> dict:
    model = init_gemini()
    prompt = f"""
당신은 맛집 전문 리서처입니다.
아래 식당의 간단한 정보를 JSON으로만 출력하세요. 다른 텍스트 없이 JSON만 출력하세요.

식당명: {restaurant_name}
주소: {address}
현재위치에서 거리: {distance:.1f}km

{{
  "식당명": "정확한 식당명",
  "선정사유": "미슐랭/블루리본/노포/방송출연 등 추천 이유 한줄 요약. 없으면 '로컬 맛집'",
  "거리": "{distance:.1f}km",
  "signature1": {{
    "메뉴명": "대표메뉴1",
    "가격": "가격 (모르면 '가격 미상')"
  }},
  "signature2": {{
    "메뉴명": "대표메뉴2",
    "가격": "가격"
  }},
  "signature3": {{
    "메뉴명": "대표메뉴3",
    "가격": "가격"
  }},
  "평점": "{restaurant_name}의 구글 평점 또는 네이버 평점",
  "카테고리": "음식 종류 한글"
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
# Gemini - 2단계: 상세 정보
# ──────────────────────────────────────────────────────
def get_detail_info(restaurant_name: str) -> dict:
    model = init_gemini()
    prompt = f"""
당신은 전 세계 맛집 전문 리서처입니다.
아래 식당에 대해 최대한 정확하게 조사하여 JSON으로만 답변하세요.
다른 텍스트, 마크다운, 설명 없이 오직 JSON만 출력하세요.
정보가 없는 항목은 반드시 빈 문자열("")로 처리하세요.

식당명: {restaurant_name}

[언어 표기 원칙]
- 한국 식당: 한글 작성 원칙. 주소는 한글만.
- 해외 식당 (영어권): 한글/영어 병기
- 해외 식당 (비영어권): 한글/영어/현지어 병기

[추천사유 검증]
- 미슐랭: 등급+선정연도
- 블루리본/레드리본: 한국 식당만
- 노포: 창업연도. 한국은 백년가게/오래가게 확인
- 방송: 생활의달인/수요미식회/허영만/백종원(출연방송명)/정용진/전현무/흑백요리사/더들리
- 해외: 한국방송 + 현지방송 모두 확인

[메뉴 가격] 한국:원화 / 해외:현지통화+한화 병기

{{
  "식당명": "식당명",
  "음식국적": "한식/일식/중화 등",
  "카테고리": "육류/해산물/면류/밥류/국탕/빵간식/기타",
  "기준메뉴": "Signature1 주재료",
  "추천사유": {{
    "미슐랭": "", "블루리본": "", "로컬인증": "",
    "노포": "", "백년가게": "", "오래가게": "",
    "방송_생활의달인": "", "방송_수요미식회": "", "방송_허영만": "",
    "방송_백종원": "", "방송_정용진": "", "방송_전현무": "",
    "방송_요리경연": "", "방송_더들리": "", "방송_기타": ""
  }},
  "signature1": {{"메뉴명": "", "주재료": "", "요리방법": "", "설명": "", "가격": ""}},
  "signature2": {{"메뉴명": "", "주재료": "", "요리방법": "", "설명": "", "가격": ""}},
  "signature3": {{"메뉴명": "", "주재료": "", "요리방법": "", "설명": "", "가격": ""}},
  "국가": "", "시도": "", "지역": "", "주소": "",
  "MRT역": "", "MRT도보": "",
  "전화번호": "",
  "예약여부": "", "대기피하는시간": "", "팁및정보": "",
  "영업시간": "", "브레이크타임": "", "라스트오더": "", "정기휴무": "",
  "가격대": "~10K / 10K~20K / 20K~40K / 40K~60K / 60K~80K / 80K~",
  "Homepage": ""
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
# 리스트 카드 표시
# ──────────────────────────────────────────────────────
def show_simple_card(info: dict, idx: int):
    sig1 = info.get("signature1", {})
    sig2 = info.get("signature2", {})
    sig3 = info.get("signature3", {})

    menus_html = ""
    for sig in [sig1, sig2, sig3]:
        if sig.get("메뉴명"):
            menus_html += (
                f'<span style="font-weight:600">{sig["메뉴명"]}</span>'
                f'<span style="opacity:0.6;font-size:0.8rem"> {sig.get("가격","")}</span>　'
            )

    st.markdown(
        f'<div class="rest-card">'
        f'<div class="rest-name">🍽️ {info.get("식당명","")}'
        f'<span class="rest-dist">　📍 {info.get("거리","")}</span></div>'
        f'<div style="margin-bottom:0.3rem">'
        f'<span class="badge badge-blue">{info.get("카테고리","")}</span> '
        f'</div>'
        f'<div class="rest-reason">✅ {info.get("선정사유","")}</div>'
        f'<div class="rest-menu">🍴 {menus_html}</div>'
        f'</div>',
        unsafe_allow_html=True
    )


# ──────────────────────────────────────────────────────
# 상세 정보 표시
# ──────────────────────────────────────────────────────
def show_detail(info: dict):
    tr   = info.get("추천사유", {})
    sig1 = info.get("signature1", {})
    sig2 = info.get("signature2", {})
    sig3 = info.get("signature3", {})

    def yn(val): return str(val).strip().upper() in ("Y","YES","*")

    st.markdown(f'<div class="restaurant-title">🍽️ {info.get("식당명","")}</div>',
                unsafe_allow_html=True)
    st.markdown(
        f'<span class="badge badge-blue">{info.get("음식국적","")}</span>'
        f'<span class="badge badge-blue">{info.get("카테고리","")}</span>'
        f'<span class="badge badge-blue">{info.get("가격대","")}</span>',
        unsafe_allow_html=True
    )
    st.divider()

    # 추천사유
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
        ("생활의달인","방송_생활의달인"),("수요미식회","방송_수요미식회"),
        ("허영만","방송_허영만"),("정용진","방송_정용진"),
        ("전현무","방송_전현무"),("더들리","방송_더들리"),
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
        st.markdown(f'<div style="line-height:2.2">{badges_html}</div>', unsafe_allow_html=True)

    # 시그니처 메뉴
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

    # 위치
    st.markdown('<p class="section-title">📍 위치 정보</p>', unsafe_allow_html=True)
    loc_html = (
        f'<div class="info-box">'
        f'🌏 <b>국가:</b> {info.get("국가","")}<br>'
        f'📌 <b>시/도:</b> {info.get("시도","")}<br>'
        f'📍 <b>지역:</b> {info.get("지역","")}<br>'
        f'🏠 <b>주소:</b> {info.get("주소","")}<br>'
    )
    if info.get("MRT역"):
        loc_html += f'🚇 <b>지하철:</b> {info.get("MRT역","")}<br>'
    if info.get("MRT도보"):
        loc_html += f'🚶 <b>도보:</b> {info.get("MRT도보","")}<br>'
    loc_html += '</div>'
    st.markdown(loc_html, unsafe_allow_html=True)

    # 운영
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

    if info.get("팁및정보"):
        st.markdown('<p class="section-title">💡 팁 & 정보</p>', unsafe_allow_html=True)
        st.info(info.get("팁및정보",""))
    if info.get("Homepage"):
        st.markdown(f'🌐 **홈페이지:** [{info["Homepage"]}]({info["Homepage"]})')


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
        "식당명": info.get("식당명",""), "국적": info.get("음식국적",""),
        "카테고리": info.get("카테고리",""), "기준메뉴": info.get("기준메뉴",""),
        "Signature메뉴1": sig1.get("메뉴명",""), "Signature1설명": sig1.get("설명",""), "Signature1가격": sig1.get("가격",""),
        "Signature메뉴2": sig2.get("메뉴명",""), "Signature2설명": sig2.get("설명",""), "Signature2가격": sig2.get("가격",""),
        "Signature메뉴3": sig3.get("메뉴명",""), "Signature3설명": sig3.get("설명",""), "Signature3가격": sig3.get("가격",""),
        "국가": info.get("국가",""), "시도": info.get("시도",""), "지역": info.get("지역",""),
        "주소": info.get("주소",""), "MRT역": info.get("MRT역",""), "전화번호": info.get("전화번호",""),
        "예약여부": info.get("예약여부",""), "대기피하는시간": info.get("대기피하는시간",""),
        "팁및정보": info.get("팁및정보",""), "영업시간": info.get("영업시간",""),
        "브레이크타임": info.get("브레이크타임",""), "라스트오더": info.get("라스트오더",""),
        "정기휴무": info.get("정기휴무",""), "가격대": info.get("가격대",""),
        "노포": tr.get("노포",""), "백년가게": yn(tr.get("백년가게","")),
        "미쉐린가이드": tr.get("미슐랭",""), "블루리본": tr.get("블루리본",""),
        "생활의달인": yn(tr.get("방송_생활의달인","")), "수요미식회": yn(tr.get("방송_수요미식회","")),
        "허영만": yn(tr.get("방송_허영만","")), "백종원": tr.get("방송_백종원",""),
        "정용진": yn(tr.get("방송_정용진","")), "전현무": yn(tr.get("방송_전현무","")),
        "요리경연": tr.get("방송_요리경연",""), "기타방송": tr.get("방송_기타",""),
        "Homepage": info.get("Homepage",""), "입력날짜": date.today(),
    }
    if os.path.exists(EXCEL_FILE):
        df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    else:
        df = pd.DataFrame([new_row])
    df.to_excel(EXCEL_FILE, sheet_name=SHEET_NAME, index=False)


# ──────────────────────────────────────────────────────
# 위치 가져오기 (브라우저 GPS)
# ──────────────────────────────────────────────────────
GPS_HTML = """
<html>
<head><meta charset="utf-8"></head>
<body style="margin:0;padding:8px;font-family:sans-serif;">
<button onclick="getLocation()" style="
    background:#ff6b35;color:white;border:none;
    padding:12px 20px;border-radius:10px;
    font-size:15px;cursor:pointer;width:100%;
    min-height:48px;
">📡 GPS 위치 감지 시작</button>
<div id="status" style="margin-top:8px;font-size:13px;color:#666;"></div>
<div style="display:flex;align-items:center;gap:8px;margin-top:6px;">
  <div id="coords" style="font-size:15px;font-weight:bold;color:#ff6b35;"></div>
  <button id="copy_btn" onclick="copyCoords()" style="
    display:none;
    background:#ff6b35;color:white;border:none;
    padding:6px 12px;border-radius:8px;
    font-size:13px;cursor:pointer;
  ">📋 복사</button>
</div>
<div id="copy_ok" style="font-size:12px;color:#4caf50;display:none;">✅ 복사됨!</div>
<script>
function getLocation() {
    var btn = document.querySelector('button');
    var status = document.getElementById('status');
    var coords = document.getElementById('coords');
    
    btn.innerText = '⏳ 위치 감지 중...';
    btn.style.background = '#888';
    status.innerText = '브라우저에서 위치 허용을 눌러주세요';
    
    if (!navigator.geolocation) {
        status.innerText = '❌ 이 브라우저는 GPS를 지원하지 않습니다';
        btn.innerText = '📡 GPS 위치 감지 시작';
        btn.style.background = '#ff6b35';
        return;
    }
    
    navigator.geolocation.getCurrentPosition(
        function(pos) {
            var lat = pos.coords.latitude.toFixed(6);
            var lng = pos.coords.longitude.toFixed(6);
            var result = lat + ', ' + lng;
            
            status.innerText = '✅ 위치 감지 성공! 아래 좌표를 복사하세요:';
            coords.innerText = result;
            btn.innerText = '✅ 감지 완료 - 다시 감지하려면 클릭';
            btn.style.background = '#4caf50';
            
            // 복사 버튼 표시
            document.getElementById('copy_btn').style.display = 'inline-block';
            
            // 클립보드 자동 복사 시도
            if (navigator.clipboard) {
                navigator.clipboard.writeText(result).then(function() {
                    status.innerHTML = '✅ 클립보드에 자동 복사됨! 입력창에 붙여넣기 하세요';
                }).catch(function() {
                    status.innerHTML = '📋 복사 버튼을 눌러 좌표를 복사하세요';
                });
            } else {
                status.innerHTML = '📋 복사 버튼을 눌러 좌표를 복사하세요';
            }
        },
        function(err) {
            var msg = '';
            if (err.code === 1) msg = '위치 접근이 거부되었습니다. 브라우저 설정에서 위치 허용 후 다시 시도하세요';
            else if (err.code === 2) msg = '위치를 찾을 수 없습니다';
            else msg = '시간 초과. 다시 시도해주세요';
            status.innerText = '❌ ' + msg;
            btn.innerText = '📡 GPS 위치 감지 시작';
            btn.style.background = '#ff6b35';
        },
        {enableHighAccuracy: true, timeout: 15000, maximumAge: 0}
    );
}
function copyCoords() {
    var coords = document.getElementById('coords').innerText;
    var copyOk = document.getElementById('copy_ok');
    if (navigator.clipboard) {
        navigator.clipboard.writeText(coords).then(function() {
            copyOk.style.display = 'block';
            setTimeout(function(){ copyOk.style.display = 'none'; }, 2000);
        });
    } else {
        // 구형 브라우저 대응
        var el = document.createElement('textarea');
        el.value = coords;
        document.body.appendChild(el);
        el.select();
        document.execCommand('copy');
        document.body.removeChild(el);
        copyOk.style.display = 'block';
        setTimeout(function(){ copyOk.style.display = 'none'; }, 2000);
    }
}
</script>
</body>
</html>
"""


# ──────────────────────────────────────────────────────
# 메인
# ──────────────────────────────────────────────────────
def main():
    check_login()

    # 세션 초기화
    if "nearby_list"   not in st.session_state: st.session_state.nearby_list   = []
    if "simple_infos"  not in st.session_state: st.session_state.simple_infos  = []
    if "detail_info"   not in st.session_state: st.session_state.detail_info   = None
    if "user_lat"      not in st.session_state: st.session_state.user_lat      = None
    if "user_lng"      not in st.session_state: st.session_state.user_lng      = None

    # 헤더
    col_title, col_logout = st.columns([5, 1])
    with col_title:
        st.markdown('<p class="main-title">🍽️ 내 주변 맛집 탐색기</p>', unsafe_allow_html=True)
        st.markdown('<p class="subtitle">현재 위치 기반으로 주변 맛집을 찾아드립니다</p>',
                    unsafe_allow_html=True)
    with col_logout:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("로그아웃", use_container_width=True):
            st.session_state.logged_in = False
            st.rerun()

    st.divider()

    # ── 탭 구성 ──
    tab1, tab2 = st.tabs(["📍 주변 맛집 검색", "🔍 식당명으로 직접 검색"])

    # ════════════════════════════════════════
    # TAB 1: 주변 맛집 검색
    # ════════════════════════════════════════
    with tab1:
        st.markdown("#### 📍 현재 위치 설정")

        # GPS 자동 감지 - 개선된 iframe 방식
        st.markdown("**① GPS 자동 감지**")
        st.caption("아래 버튼 클릭 → 위치 허용 → 좌표가 자동으로 클립보드에 복사됩니다")
        st.components.v1.html(GPS_HTML, height=110)

        st.markdown("**② 좌표 입력** (GPS 감지 후 붙여넣기 또는 직접 입력)")
        loc_input = st.text_input(
            "위치 (위도,경도)",
            placeholder="예) 37.7419, 127.0300",
            label_visibility="collapsed",
            key="loc_input"
        )

        with st.expander("📌 좌표 찾는 다른 방법"):
            st.markdown("""
- **구글 지도** → 원하는 위치 **길게 누르기** → 하단에 좌표 표시 → 복사
- **네이버 지도** → 현재위치 검색 → URL의 `@위도,경도` 부분 복사
- **의정부 브라운스톤흥선**: `37.7419, 127.0300`
            """)

        if loc_input and "," in loc_input:
            try:
                parts = loc_input.strip().split(",")
                lat, lng = float(parts[0].strip()), float(parts[1].strip())
                st.session_state.user_lat = lat
                st.session_state.user_lng = lng
                st.success(f"✅ 위치 설정됨: {lat:.4f}, {lng:.4f}")
            except:
                st.error("위도,경도 형식으로 입력해주세요. 예) 37.7419, 127.0300")

        st.divider()

        # 반경 & 카테고리 설정
        st.markdown("#### 🔧 검색 설정")
        col_r, col_k = st.columns([1, 1])
        with col_r:
            radius = st.number_input(
                "반경 (km)",
                min_value=0.5, max_value=50.0,
                value=3.0, step=0.5,
                format="%.1f"
            )
        with col_k:
            keyword = st.text_input(
                "카테고리/키워드 (선택)",
                placeholder="예) 한식, 냉면, 스시...",
                label_visibility="visible"
            )

        # 검색 버튼
        if st.button("🔍 주변 맛집 검색", use_container_width=True, type="primary"):
            if not st.session_state.user_lat:
                st.error("❌ 먼저 위치를 설정해주세요!")
            else:
                with st.spinner(f"반경 {radius}km 내 식당 검색 중..."):
                    try:
                        restaurants = search_nearby_restaurants(
                            st.session_state.user_lat,
                            st.session_state.user_lng,
                            radius, keyword
                        )
                        st.session_state.nearby_list  = restaurants
                        st.session_state.simple_infos = []
                        st.success(f"✅ {len(restaurants)}개 식당 발견!")
                    except Exception as e:
                        st.error(f"❌ 검색 오류: {e}")


        # 세션 - 페이지 인덱스 초기화
        if "page_index" not in st.session_state: st.session_state.page_index = 0

        # 검색 결과 리스트
        if st.session_state.nearby_list:
            st.divider()
            total = len(st.session_state.nearby_list)
            st.markdown(f'<p class="section-title">📋 검색 결과 총 {total}개</p>',
                        unsafe_allow_html=True)

            # 현재 페이지 범위
            page_start = st.session_state.page_index * 10
            page_end   = min(page_start + 10, total)
            current_batch = st.session_state.nearby_list[page_start:page_end]

            # 조회 버튼
            btn_label = f"🤖 {page_start+1}~{page_end}번 맛집 정보 조회"
            if st.button(btn_label, use_container_width=True, type="primary"):
                simple_infos = list(st.session_state.simple_infos)
                progress = st.progress(0, text="Gemini 정보 조회 중...")
                for i, r in enumerate(current_batch):
                    try:
                        info = get_simple_info(r["name"], r["address"], r["distance"])
                        info["_place_data"] = r
                    except:
                        info = {
                            "식당명": r["name"], "선정사유": "로컬 맛집",
                            "거리": f'{r["distance"]:.1f}km', "카테고리": "식당",
                            "signature1": {"메뉴명": "-", "가격": "-"},
                            "signature2": {"메뉴명": "", "가격": ""},
                            "signature3": {"메뉴명": "", "가격": ""},
                            "_place_data": r
                        }
                    names = [x.get("식당명","") for x in simple_infos]
                    if info["식당명"] not in names:
                        simple_infos.append(info)
                    progress.progress((i+1)/len(current_batch),
                                      text=f"({page_start+i+1}/{total}) {r['name']} 조회 중...")
                st.session_state.simple_infos = simple_infos
                progress.empty()
                st.rerun()

            # 간단 카드 + 상세보기 버튼
            if st.session_state.simple_infos:
                st.markdown(f'<p class="section-title">🍽️ 맛집 목록 ({len(st.session_state.simple_infos)}개 조회됨)</p>',
                            unsafe_allow_html=True)
                st.caption("아래 버튼을 클릭하면 상세 정보를 바로 볼 수 있습니다")

                for i, info in enumerate(st.session_state.simple_infos):
                    show_simple_card(info, i)
                    if st.button(f"📄 {info.get('식당명','')} 상세보기",
                                 key=f"detail_btn_{i}",
                                 use_container_width=True):
                        with st.spinner(f"'{info.get('식당명','')}' 상세 정보 조사 중..."):
                            try:
                                detail = get_detail_info(info.get("식당명",""))
                                st.session_state.detail_info = detail
                                st.rerun()
                            except Exception as e:
                                st.error(f"❌ {e}")

                # 추가 10개 버튼
                if page_end < total:
                    st.divider()
                    remain = min(10, total - page_end)
                    if st.button(f"➕ 추가 {remain}개 더 보기 ({page_end+1}~{min(page_end+10, total)}번)",
                                 use_container_width=True):
                        st.session_state.page_index += 1
                        st.rerun()
                else:
                    st.caption("✅ 모든 식당을 조회했습니다.")

    # ════════════════════════════════════════
    # TAB 2: 직접 검색
    # ════════════════════════════════════════
    with tab2:
        st.markdown("#### 🔍 식당명으로 직접 검색")
        restaurant_name = st.text_input(
            "식당명",
            placeholder="예) 진주회관  /  홍콩, Tim Ho Wan",
            key="direct_input",
            label_visibility="collapsed"
        )
        if st.button("🔍 정보 조회", use_container_width=True, type="primary", key="direct_btn"):
            if restaurant_name:
                with st.spinner(f"'{restaurant_name}' 정보 조사 중... (20~40초 소요)"):
                    try:
                        detail = get_detail_info(restaurant_name)
                        st.session_state.detail_info = detail
                    except Exception as e:
                        st.error(f"❌ {e}")
            else:
                st.warning("식당명을 입력해주세요.")

    # ── 상세 정보 표시 (탭 공통) ──
    if st.session_state.detail_info:
        st.divider()
        show_detail(st.session_state.detail_info)
        st.divider()

        if st.button("💾 엑셀에 저장", type="primary", use_container_width=True):
            try:
                append_to_excel(st.session_state.detail_info)
                st.success(f"✅ '{st.session_state.detail_info.get('식당명','')}' 저장 완료!")
                if os.path.exists(EXCEL_FILE):
                    with open(EXCEL_FILE, "rb") as f:
                        st.download_button(
                            "📥 엑셀 파일 다운로드", data=f,
                            file_name=EXCEL_FILE,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
            except Exception as e:
                st.error(f"저장 오류: {e}")


if __name__ == "__main__":
    main()
