"""
🍽️  맛집 탐색기 - 내 주변 맛집 검색 + 상세정보
Google Places API + Gemini API + Streamlit
웹/모바일 최적화 - 최종버전
"""

import streamlit as st
import google.generativeai as genai
import pandas as pd
import json
import os
import requests
import math
import time
import re
from datetime import date

# ──────────────────────────────────────────────────────
# ✅ 설정
# ──────────────────────────────────────────────────────
GEMINI_API_KEY  = st.secrets["GEMINI_API_KEY"]
ACCESS_PASSWORD = st.secrets["ACCESS_PASSWORD"]
PLACES_API_KEY  = st.secrets["PLACES_API_KEY"]
EXCEL_FILE  = "맛집정보.xlsx"
SHEET_NAME  = "Sheet1"
BATCH_SIZE  = 5
MAX_RESULTS = 20

# ──────────────────────────────────────────────────────
# 페이지 설정
# ──────────────────────────────────────────────────────
st.set_page_config(
    page_title="🍽️ 내 주변 맛집",
    page_icon="🍽️",
    layout="centered",
    initial_sidebar_state="collapsed"
)

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
    .rest-card {
        border-radius: 14px; padding: 1rem; margin: 0.6rem 0;
        border: 1.5px solid rgba(128,128,128,0.3);
        background: rgba(128,128,128,0.08) !important;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    }
    .rest-rank   { font-size: 1.2rem; font-weight: 900; color: #ff6b35; }
    .rest-name   { font-size: 1.05rem; font-weight: 800; margin-bottom: 0.3rem; }
    .rest-dist   { font-size: 0.82rem; color: #ff6b35 !important; font-weight: 600; }
    .menu-card {
        border-radius: 14px; padding: 1.1rem 1rem; margin: 0.6rem 0;
        border: 1.5px solid rgba(128,128,128,0.3);
        background: rgba(128,128,128,0.08) !important;
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
    .stButton > button { min-height: 52px !important; font-size: 1.05rem !important; font-weight: 600 !important; border-radius: 12px !important; }
    .stTextInput > div > div > input { font-size: 1.05rem !important; min-height: 52px !important; border-radius: 12px !important; }
    #MainMenu, footer, header { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────────────
# GPS HTML
# ──────────────────────────────────────────────────────
GPS_HTML = """
<html><head><meta charset="utf-8"></head>
<body style="margin:0;padding:8px;font-family:sans-serif;">
<button onclick="getLocation()" style="background:#ff6b35;color:white;border:none;
    padding:12px 20px;border-radius:10px;font-size:15px;cursor:pointer;width:100%;min-height:48px;">
    📡 GPS 위치 감지 시작</button>
<div id="status" style="margin-top:8px;font-size:13px;color:#888;"></div>
<div style="margin-top:10px;">
  <div id="coords" style="font-size:16px;font-weight:bold;color:#ff6b35;margin-bottom:8px;word-break:break-all;"></div>
  <button id="copy_btn" onclick="copyCoords()" style="display:none;background:#ff6b35;color:white;border:none;
    padding:10px 20px;border-radius:10px;font-size:14px;cursor:pointer;width:100%;margin-top:4px;">
    📋 좌표 복사하기</button>
  <div id="copy_ok" style="font-size:13px;color:#4caf50;margin-top:6px;display:none;font-weight:bold;">
    ✅ 복사됐습니다! 아래 입력창에 붙여넣기 하세요</div>
</div>
<script>
function getLocation() {
    var btn=document.querySelector('button'), status=document.getElementById('status');
    btn.innerText='⏳ 위치 감지 중...'; btn.style.background='#888';
    status.innerText='브라우저에서 위치 허용을 눌러주세요';
    if (!navigator.geolocation) { status.innerText='❌ GPS 미지원 브라우저'; btn.innerText='📡 GPS 위치 감지 시작'; btn.style.background='#ff6b35'; return; }
    navigator.geolocation.getCurrentPosition(
        function(pos) {
            var lat=pos.coords.latitude.toFixed(6), lng=pos.coords.longitude.toFixed(6), result=lat+', '+lng;
            document.getElementById('coords').innerText=result;
            document.getElementById('copy_btn').style.display='inline-block';
            btn.innerText='✅ 감지 완료 - 다시 감지하려면 클릭'; btn.style.background='#4caf50';
            if (navigator.clipboard && window.isSecureContext) {
                navigator.clipboard.writeText(result).then(function(){ status.innerText='✅ 클립보드에 자동 복사됨! 입력창에 붙여넣기 하세요'; })
                .catch(function(){ status.innerText='📋 복사 버튼을 눌러 좌표를 복사하세요'; });
            } else { status.innerText='📋 복사 버튼을 눌러 좌표를 복사하세요'; }
        },
        function(err) {
            var msg=err.code===1?'위치 접근 거부됨. 브라우저 설정에서 위치 허용 후 재시도':err.code===2?'위치를 찾을 수 없습니다':'시간 초과. 다시 시도해주세요';
            status.innerText='❌ '+msg; btn.innerText='📡 GPS 위치 감지 시작'; btn.style.background='#ff6b35';
        },
        {enableHighAccuracy:true, timeout:15000, maximumAge:0}
    );
}
function copyCoords() {
    var coords=document.getElementById('coords').innerText.trim();
    var copyOk=document.getElementById('copy_ok'), btn=document.getElementById('copy_btn');
    function showCopied() { copyOk.style.display='block'; btn.style.background='#4caf50'; btn.innerText='✅ 복사됨!';
        setTimeout(function(){ copyOk.style.display='none'; btn.style.background='#ff6b35'; btn.innerText='📋 좌표 복사하기'; },2500); }
    if (navigator.clipboard && window.isSecureContext) { navigator.clipboard.writeText(coords).then(showCopied).catch(function(){ fallbackCopy(coords); showCopied(); }); }
    else { fallbackCopy(coords); showCopied(); }
}
function fallbackCopy(text) {
    var el=document.createElement('textarea'); el.value=text; el.style.position='fixed'; el.style.opacity='0';
    document.body.appendChild(el); el.focus(); el.select(); try{document.execCommand('copy');}catch(e){} document.body.removeChild(el);
}
</script></body></html>
"""

# ──────────────────────────────────────────────────────
# 로그인
# ──────────────────────────────────────────────────────
def check_login():
    if "logged_in" not in st.session_state: st.session_state.logged_in = False
    if not st.session_state.logged_in:
        st.markdown("## 🍽️ 내 주변 맛집 탐색기")
        st.markdown("승인된 사용자만 접속 가능합니다.")
        pw = st.text_input("비밀번호", type="password", placeholder="비밀번호 입력", label_visibility="collapsed")
        if st.button("로그인", use_container_width=True, type="primary"):
            if pw == ACCESS_PASSWORD: st.session_state.logged_in = True; st.rerun()
            else: st.error("❌ 비밀번호가 올바르지 않습니다.")
        st.stop()

# ──────────────────────────────────────────────────────
# Gemini
# ──────────────────────────────────────────────────────
@st.cache_resource
def init_gemini():
    genai.configure(api_key=GEMINI_API_KEY)
    return genai.GenerativeModel("gemini-2.5-flash")

# ──────────────────────────────────────────────────────
# 거리 계산
# ──────────────────────────────────────────────────────
def calc_distance(lat1, lng1, lat2, lng2):
    R = 6371
    d_lat = math.radians(lat2-lat1); d_lng = math.radians(lng2-lng1)
    a = math.sin(d_lat/2)**2 + math.cos(math.radians(lat1))*math.cos(math.radians(lat2))*math.sin(d_lng/2)**2
    return R*2*math.asin(math.sqrt(a))

# ──────────────────────────────────────────────────────
# 중요도 점수
# ──────────────────────────────────────────────────────
def calc_priority_score(info: dict) -> int:
    score = 0
    tr = info.get("추천사유", {})
    michelin = tr.get("미슐랭","")
    if "3스타" in michelin: score += 100
    elif "2스타" in michelin: score += 80
    elif "1스타" in michelin: score += 60
    elif "빕구르망" in michelin: score += 40
    elif "셀렉티드" in michelin: score += 20
    if tr.get("블루리본"):
        score += 50 if "블루리본" in tr["블루리본"] else 30
    if tr.get("로컬인증"): score += 40
    nopo = tr.get("노포","")
    if nopo:
        years = re.findall(r'\d{4}', nopo)
        if years:
            nopo_years = date.today().year - int(years[0])
            if nopo_years >= 50: score += 50
            elif nopo_years >= 30: score += 30
            elif nopo_years >= 10: score += 10
    if tr.get("백년가게"): score += 20
    if tr.get("오래가게"): score += 15
    if tr.get("방송_생활의달인"): score += 40
    if tr.get("방송_수요미식회"): score += 40
    if tr.get("방송_백종원"):     score += 40
    if tr.get("방송_더들리"):     score += 40
    if tr.get("방송_허영만"):     score += 25
    if tr.get("방송_정용진"):     score += 25
    if tr.get("방송_전현무"):     score += 25
    if tr.get("방송_요리경연"):   score += 25
    if tr.get("방송_빅페이스"):   score += 20
    if tr.get("방송_나의시선"):   score += 10
    if tr.get("방송_기타"):       score += 10
    waiting = info.get("대기여부","")
    if "상시" in waiting: score += 30
    elif "피크" in waiting: score += 20
    return score

# ──────────────────────────────────────────────────────
# 추천사유 HTML (중요도 순)
# ──────────────────────────────────────────────────────
def build_reason_html(tr: dict, waiting: str = "") -> str:
    items = []
    if tr.get("미슐랭"):   items.append(("red",   f"⭐ {tr['미슐랭']}"))
    if tr.get("블루리본"): items.append(("red",   f"🎗️ {tr['블루리본']}"))
    if tr.get("로컬인증"): items.append(("red",   f"🏆 {tr['로컬인증']}"))
    nopo = tr.get("노포","")
    if nopo:
        years = re.findall(r'\d{4}', nopo)
        nopo_years = date.today().year - int(years[0]) if years else 0
        priority = "red" if nopo_years >= 50 else "orange" if nopo_years >= 30 else "blue"
        label = f"🏛️ {nopo}" + (f" ({nopo_years}년)" if nopo_years > 0 else "")
        items.append((priority, label))
    if tr.get("백년가게"): items.append(("red",    "🏅 백년가게"))
    if tr.get("오래가게"): items.append(("orange", "🏅 오래가게"))
    if tr.get("방송_생활의달인"): items.append(("red",    "📺 생활의달인"))
    if tr.get("방송_수요미식회"): items.append(("red",    "📺 수요미식회"))
    if tr.get("방송_백종원"):     items.append(("red",    f"📺 백종원({tr['방송_백종원']})"))
    if tr.get("방송_더들리"):     items.append(("red",    "📺 더들리"))
    if tr.get("방송_허영만"):     items.append(("orange", "📺 허영만"))
    if tr.get("방송_정용진"):     items.append(("orange", "📺 정용진"))
    if tr.get("방송_전현무"):     items.append(("orange", "📺 전현무"))
    if tr.get("방송_요리경연"):   items.append(("orange", f"🏆 {tr['방송_요리경연']}"))
    if tr.get("방송_빅페이스"):   items.append(("orange", "📺 빅페이스"))
    if tr.get("방송_나의시선"):   items.append(("blue",   "📺 나의시선"))
    if tr.get("방송_기타"):       items.append(("blue",   f"📺 {tr['방송_기타']}"))
    if "상시" in waiting:   items.append(("red",    "⏰ 상시웨이팅"))
    elif "피크" in waiting: items.append(("orange", "⏰ 피크타임웨이팅"))
    if not items: return '<span class="badge badge-blue">로컬 맛집</span>'
    color_map = {"red": "badge-red", "orange": "badge", "blue": "badge-blue"}
    return " ".join(f'<span class="badge {color_map[p]}">{t}</span>' for p,t in items)

# ──────────────────────────────────────────────────────
# Google Places 검색 (노포 키워드 처리)
# ──────────────────────────────────────────────────────
def search_nearby_restaurants(lat, lng, radius_km, keyword=""):
    radius_m = int(radius_km * 1000)
    url = "https://maps.googleapis.com/maps/api/place/nearbysearch/json"

    nopo_filter = None
    nopo_keywords = ["노포", "백년가게", "오래가게", "오래된", "전통 식당"]
    nopo_year_filter = 0

    if keyword:
        for nk in nopo_keywords:
            if nk in keyword:
                nopo_filter = keyword
                keyword = ""
                years = re.findall(r'\d+', nopo_filter)
                if years: nopo_year_filter = int(years[0])
                break

    params = {"location": f"{lat},{lng}", "radius": radius_m, "type": "restaurant",
              "key": PLACES_API_KEY, "language": "ko"}
    if keyword: params["keyword"] = keyword

    results = []
    resp = requests.get(url, params=params).json()
    results.extend(resp.get("results", []))
    # B) 경량화: 1페이지만 검색 (최대 20개)

    restaurants = []
    for r in results[:MAX_RESULTS]:
        loc = r["geometry"]["location"]
        dist = calc_distance(lat, lng, loc["lat"], loc["lng"])
        restaurants.append({
            "place_id": r.get("place_id",""), "name": r.get("name",""),
            "address": r.get("vicinity",""), "rating": r.get("rating",0),
            "user_ratings": r.get("user_ratings_total",0),
            "lat": loc["lat"], "lng": loc["lng"], "distance": dist,
            "nopo_filter": nopo_filter, "nopo_year": nopo_year_filter,
        })
    return sorted(restaurants, key=lambda x: x["distance"]), nopo_filter

# ──────────────────────────────────────────────────────
# Gemini 배치 조회 (속도 최적화)
# ──────────────────────────────────────────────────────
def get_simple_info_batch(restaurants: list) -> list:
    model = init_gemini()
    names_list = "\n".join([
        f"{i+1}. 식당명: {r['name']}, 주소: {r['address']}, 거리: {r['distance']:.1f}km"
        for i, r in enumerate(restaurants)
    ])
    prompt = f"""맛집 리서처. 아래 식당들 정보를 JSON 배열로만 출력. 다른 텍스트 없이 JSON만.

{names_list}

[{{"순번":1,"식당명":"","카테고리":"","거리":"","추천사유":{{"미슐랭":"","블루리본":"","로컬인증":"","노포":"예)1965년창업","백년가게":"Y또는빈값","오래가게":"Y또는빈값","방송_생활의달인":"Y또는빈값","방송_수요미식회":"Y또는빈값","방송_백종원":"출연방송명","방송_허영만":"Y또는빈값","방송_정용진":"Y또는빈값","방송_전현무":"Y또는빈값","방송_요리경연":"대회명","방송_더들리":"Y또는빈값","방송_빅페이스":"Y또는빈값","방송_기타":""}},"대기여부":"상시웨이팅필수/피크타임웨이팅/웨이팅없음","signature1":{{"메뉴명":"","가격":""}},"signature2":{{"메뉴명":"","가격":""}},"signature3":{{"메뉴명":"","가격":""}}}}]
"""
    try:
        response = model.generate_content(prompt)
        raw = response.text.strip()
        if "```" in raw:
            parts = raw.split("```")
            for part in parts:
                part = part.strip()
                if part.startswith("json"): part = part[4:].strip()
                try: return json.loads(part)
                except: continue
        return json.loads(raw)
    except: return []

# ──────────────────────────────────────────────────────
# Gemini 상세 정보
# ──────────────────────────────────────────────────────
def get_detail_info(restaurant_name: str) -> dict:
    model = init_gemini()
    prompt = f"""맛집 전문 리서처. 식당명: {restaurant_name}
JSON만 출력. 정보없는항목=빈문자열. 한국=한글, 해외영어권=한글/영어, 비영어권=한글/영어/현지어. 가격:한국=원화,해외=현지통화+한화.
{{"식당명":"","음식국적":"","카테고리":"","기준메뉴":"","추천사유":{{"미슐랭":"","블루리본":"","로컬인증":"","노포":"","백년가게":"","오래가게":"","방송_생활의달인":"","방송_수요미식회":"","방송_허영만":"","방송_백종원":"","방송_정용진":"","방송_전현무":"","방송_요리경연":"","방송_더들리":"","방송_빅페이스":"","방송_나의시선":"","방송_기타":""}},"signature1":{{"메뉴명":"","주재료":"","요리방법":"","설명":"","가격":""}},"signature2":{{"메뉴명":"","주재료":"","요리방법":"","설명":"","가격":""}},"signature3":{{"메뉴명":"","주재료":"","요리방법":"","설명":"","가격":""}},"국가":"","시도":"","지역":"","주소":"","MRT역":"","MRT도보":"","전화번호":"","예약여부":"","대기피하는시간":"","팁및정보":"","영업시간":"","브레이크타임":"","라스트오더":"","정기휴무":"","가격대":"","Homepage":""}}
"""
    response = model.generate_content(prompt)
    raw = response.text.strip()
    if "```" in raw:
        parts = raw.split("```")
        for part in parts:
            part = part.strip()
            if part.startswith("json"): part = part[4:].strip()
            try: return json.loads(part)
            except: continue
    return json.loads(raw)

# ──────────────────────────────────────────────────────
# 간단 카드
# ──────────────────────────────────────────────────────
def show_simple_card(info: dict, rank: int):
    tr   = info.get("추천사유", {})
    sig1 = info.get("signature1", {})
    sig2 = info.get("signature2", {})
    sig3 = info.get("signature3", {})
    waiting = info.get("대기여부","")
    reason_html = build_reason_html(tr, waiting)
    menus_html = "".join([
        f'<b>{s["메뉴명"]}</b> <span style="opacity:0.6;font-size:0.8rem">{s.get("가격","")}</span>　'
        for s in [sig1,sig2,sig3] if s.get("메뉴명") and s["메뉴명"] != "-"
    ])
    st.markdown(
        f'<div class="rest-card">'
        f'<div style="display:flex;justify-content:space-between;align-items:flex-start;">'
        f'<span class="rest-rank">#{rank}</span>'
        f'<span class="rest-dist">📍 {info.get("거리","")}</span>'
        f'</div>'
        f'<div class="rest-name">🍽️ {info.get("식당명","")}</div>'
        f'<span class="badge badge-blue">{info.get("카테고리","")}</span><br>'
        f'<div style="line-height:2.2;margin:6px 0">{reason_html}</div>'
        f'<div style="font-size:0.85rem;margin-top:6px">🍴 {menus_html}</div>'
        f'</div>', unsafe_allow_html=True
    )

# ──────────────────────────────────────────────────────
# 상세 정보
# ──────────────────────────────────────────────────────
def show_detail(info: dict):
    tr   = info.get("추천사유", {})
    sig1 = info.get("signature1", {})
    sig2 = info.get("signature2", {})
    sig3 = info.get("signature3", {})
    waiting = info.get("대기피하는시간","") + info.get("팁및정보","")
    st.markdown(f'<div class="restaurant-title">🍽️ {info.get("식당명","")}</div>', unsafe_allow_html=True)
    st.markdown(
        f'<span class="badge badge-blue">{info.get("음식국적","")}</span>'
        f'<span class="badge badge-blue">{info.get("카테고리","")}</span>'
        f'<span class="badge badge-blue">{info.get("가격대","")}</span>',
        unsafe_allow_html=True
    )
    st.divider()
    reason_html = build_reason_html(tr, waiting)
    if reason_html:
        st.markdown('<p class="section-title">📌 추천사유 (중요도순)</p>', unsafe_allow_html=True)
        st.markdown(f'<div style="line-height:2.5">{reason_html}</div>', unsafe_allow_html=True)
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
                f'</div>', unsafe_allow_html=True
            )
    st.markdown('<p class="section-title">📍 위치 정보</p>', unsafe_allow_html=True)
    loc_html = f'<div class="info-box">🌏 <b>국가:</b> {info.get("국가","")}<br>📌 <b>시/도:</b> {info.get("시도","")}<br>📍 <b>지역:</b> {info.get("지역","")}<br>🏠 <b>주소:</b> {info.get("주소","")}<br>'
    if info.get("MRT역"):  loc_html += f'🚇 <b>지하철:</b> {info.get("MRT역","")}<br>'
    if info.get("MRT도보"): loc_html += f'🚶 <b>도보:</b> {info.get("MRT도보","")}<br>'
    st.markdown(loc_html + '</div>', unsafe_allow_html=True)
    st.markdown('<p class="section-title">📋 운영 정보</p>', unsafe_allow_html=True)
    ops = f'<div class="info-box">📞 <b>전화:</b> {info.get("전화번호","")}<br>🕐 <b>영업시간:</b> {info.get("영업시간","")}<br>'
    if info.get("브레이크타임"): ops += f'☕ <b>브레이크타임:</b> {info.get("브레이크타임","")}<br>'
    if info.get("라스트오더"):   ops += f'⏱️ <b>라스트오더:</b> {info.get("라스트오더","")}<br>'
    ops += f'🚫 <b>정기휴무:</b> {info.get("정기휴무","")}<br>📅 <b>예약:</b> {info.get("예약여부","")}<br>'
    if info.get("대기피하는시간"): ops += f'⏰ <b>대기 피하는 시간:</b> {info.get("대기피하는시간","")}<br>'
    st.markdown(ops + '</div>', unsafe_allow_html=True)
    if info.get("팁및정보"):
        st.markdown('<p class="section-title">💡 팁 & 정보</p>', unsafe_allow_html=True)
        st.info(info.get("팁및정보",""))
    if info.get("Homepage"):
        st.markdown(f'🌐 **홈페이지:** [{info["Homepage"]}]({info["Homepage"]})')

# ──────────────────────────────────────────────────────
# 엑셀 저장
# ──────────────────────────────────────────────────────
def append_to_excel(info: dict):
    sig1=info.get("signature1",{}); sig2=info.get("signature2",{}); sig3=info.get("signature3",{}); tr=info.get("추천사유",{})
    def yn(val): return "*" if str(val).strip().upper() in ("Y","YES","*") else ""
    new_row = {
        "식당명": info.get("식당명",""), "국적": info.get("음식국적",""), "카테고리": info.get("카테고리",""), "기준메뉴": info.get("기준메뉴",""),
        "Signature메뉴1": sig1.get("메뉴명",""), "Signature1설명": sig1.get("설명",""), "Signature1가격": sig1.get("가격",""),
        "Signature메뉴2": sig2.get("메뉴명",""), "Signature2설명": sig2.get("설명",""), "Signature2가격": sig2.get("가격",""),
        "Signature메뉴3": sig3.get("메뉴명",""), "Signature3설명": sig3.get("설명",""), "Signature3가격": sig3.get("가격",""),
        "국가": info.get("국가",""), "시도": info.get("시도",""), "지역": info.get("지역",""), "주소": info.get("주소",""),
        "MRT역": info.get("MRT역",""), "전화번호": info.get("전화번호",""), "예약여부": info.get("예약여부",""),
        "대기피하는시간": info.get("대기피하는시간",""), "팁및정보": info.get("팁및정보",""),
        "영업시간": info.get("영업시간",""), "브레이크타임": info.get("브레이크타임",""),
        "라스트오더": info.get("라스트오더",""), "정기휴무": info.get("정기휴무",""), "가격대": info.get("가격대",""),
        "노포": tr.get("노포",""), "백년가게": yn(tr.get("백년가게","")), "오래가게": yn(tr.get("오래가게","")),
        "미쉐린가이드": tr.get("미슐랭",""), "블루리본": tr.get("블루리본",""), "로컬인증": tr.get("로컬인증",""),
        "생활의달인": yn(tr.get("방송_생활의달인","")), "수요미식회": yn(tr.get("방송_수요미식회","")),
        "허영만": yn(tr.get("방송_허영만","")), "백종원": tr.get("방송_백종원",""),
        "정용진": yn(tr.get("방송_정용진","")), "전현무": yn(tr.get("방송_전현무","")),
        "요리경연": tr.get("방송_요리경연",""), "더들리": yn(tr.get("방송_더들리","")),
        "빅페이스": yn(tr.get("방송_빅페이스","")), "기타방송": tr.get("방송_기타",""),
        "Homepage": info.get("Homepage",""), "입력날짜": date.today(),
    }
    if os.path.exists(EXCEL_FILE):
        df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    else:
        df = pd.DataFrame([new_row])
    df.to_excel(EXCEL_FILE, sheet_name=SHEET_NAME, index=False)

# ──────────────────────────────────────────────────────
# 세션 리셋
# ──────────────────────────────────────────────────────
def reset_search():
    st.session_state.nearby_list  = []
    st.session_state.simple_infos = []
    st.session_state.page_index   = 0
    st.session_state.detail_info  = None
    st.session_state.nopo_filter  = None

# ──────────────────────────────────────────────────────
# 메인
# ──────────────────────────────────────────────────────
def main():
    check_login()

    for key, default in [
        ("nearby_list",[]), ("simple_infos",[]), ("page_index",0),
        ("detail_info",None), ("user_lat",None), ("user_lng",None),
        ("nopo_filter",None), ("prev_settings",None), ("detail_cache",{})
    ]:
        if key not in st.session_state: st.session_state[key] = default

    col_title, col_logout = st.columns([5,1])
    with col_title:
        st.markdown('<p class="main-title">🍽️ 내 주변 맛집 탐색기</p>', unsafe_allow_html=True)
        st.markdown('<p class="subtitle">현재 위치 기반으로 주변 맛집을 찾아드립니다</p>', unsafe_allow_html=True)
    with col_logout:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("로그아웃", use_container_width=True):
            st.session_state.logged_in = False; st.rerun()

    st.divider()
    tab1, tab2 = st.tabs(["📍 주변 맛집 검색", "🔍 식당명으로 직접 검색"])

    with tab1:
        st.markdown("#### 📍 현재 위치 설정")
        st.markdown("**① GPS 자동 감지**")
        st.caption("아래 버튼 클릭 → 위치 허용 → 좌표가 자동으로 클립보드에 복사됩니다")
        st.components.v1.html(GPS_HTML, height=160)
        st.markdown("**② 좌표 입력**")
        loc_input = st.text_input("위치", placeholder="예) 37.7419, 127.0300", label_visibility="collapsed", key="loc_input")
        with st.expander("📌 좌표 찾는 다른 방법"):
            st.markdown("- **구글 지도** → 원하는 위치 **길게 누르기** → 하단에 좌표 표시\n- **의정부 브라운스톤흥선**: `37.7419, 127.0300`")

        if loc_input and "," in loc_input:
            try:
                parts = loc_input.strip().split(",")
                lat, lng = float(parts[0].strip()), float(parts[1].strip())
                if st.session_state.user_lat != lat or st.session_state.user_lng != lng:
                    st.session_state.user_lat = lat; st.session_state.user_lng = lng
                    reset_search()
                st.success(f"✅ 위치 설정됨: {lat:.4f}, {lng:.4f}")
            except: st.error("위도,경도 형식으로 입력해주세요.")

        st.divider()
        st.markdown("#### 🔧 검색 설정")
        col_r, col_k = st.columns([1,1])
        with col_r: radius = st.number_input("반경 (km)", min_value=0.5, max_value=50.0, value=3.0, step=0.5, format="%.1f")
        with col_k: keyword = st.text_input("카테고리/키워드 (선택)", placeholder="예) 한식, 냉면, 노포 30년이상...", key="keyword_input")

        # 설정 변경 감지
        current_settings = (st.session_state.get("user_lat"), st.session_state.get("user_lng"), radius, keyword)
        if st.session_state.prev_settings and st.session_state.prev_settings != current_settings:
            reset_search()
        st.session_state.prev_settings = current_settings

        if st.button("🔍 주변 맛집 검색", use_container_width=True, type="primary"):
            if not st.session_state.user_lat:
                st.error("❌ 먼저 위치를 설정해주세요!")
            else:
                with st.spinner(f"반경 {radius}km 내 식당 검색 중..."):
                    try:
                        restaurants, nopo_filter = search_nearby_restaurants(
                            st.session_state.user_lat, st.session_state.user_lng, radius, keyword)
                        reset_search()
                        st.session_state.nearby_list = restaurants
                        st.session_state.nopo_filter = nopo_filter
                        if nopo_filter:
                            st.info(f"💡 노포 검색: Google Places에서 직접 노포 검색이 불가하여 반경 내 전체 식당을 검색 후 Gemini가 노포 여부를 판별합니다.")
                        st.success(f"✅ {len(restaurants)}개 식당 발견!")
                    except Exception as e: st.error(f"❌ 검색 오류: {e}")

        if st.session_state.nearby_list:
            st.divider()
            total = len(st.session_state.nearby_list)
            page_start = st.session_state.page_index * BATCH_SIZE
            page_end   = min(page_start + BATCH_SIZE, total)
            current_batch = st.session_state.nearby_list[page_start:page_end]
            already_names = [x.get("식당명","") for x in st.session_state.simple_infos]
            batch_not_queried = [r for r in current_batch if r["name"] not in already_names]

            if batch_not_queried:
                btn_label = f"🤖 {page_start+1}~{page_end}번 맛집 정보 조회 ({len(batch_not_queried)}개)"
                if st.button(btn_label, use_container_width=True, type="primary"):
                    with st.spinner(f"Gemini가 {len(batch_not_queried)}개 식당 일괄 조회 중..."):
                        batch_results = get_simple_info_batch(batch_not_queried)
                        for i, r in enumerate(batch_not_queried):
                            matched = next((x for x in batch_results if str(x.get("순번",""))==str(i+1)), None)
                            if matched:
                                matched["거리"] = f"{r['distance']:.1f}km"
                                matched["_place_data"] = r
                                st.session_state.simple_infos.append(matched)
                            else:
                                st.session_state.simple_infos.append({
                                    "식당명": r["name"], "카테고리": "식당",
                                    "거리": f"{r['distance']:.1f}km", "추천사유": {}, "대기여부": "",
                                    "signature1": {"메뉴명":"-","가격":"-"},
                                    "signature2": {"메뉴명":"","가격":""},
                                    "signature3": {"메뉴명":"","가격":""},
                                    "_place_data": r
                                })
                    st.rerun()

            if st.session_state.simple_infos:
                sorted_infos = sorted(st.session_state.simple_infos, key=calc_priority_score, reverse=True)
                st.markdown(f'<p class="section-title">🍽️ 맛집 목록 - 중요도순 ({len(sorted_infos)}개)</p>', unsafe_allow_html=True)
                st.caption("🔴 High중요도  🟠 Mid중요도  🔵 Low중요도")

                for i, info in enumerate(sorted_infos):
                    show_simple_card(info, i+1)
                    if st.button(f"📄 {info.get('식당명','')} 상세보기", key=f"detail_btn_{i}", use_container_width=True):
                        name = info.get("식당명","")
                        # C) 캐시 확인
                        if name in st.session_state.detail_cache:
                            st.session_state.detail_info = st.session_state.detail_cache[name]
                            st.rerun()
                        else:
                            with st.spinner(f"'{name}' 상세 정보 조사 중..."):
                                try:
                                    detail = get_detail_info(name)
                                    st.session_state.detail_cache[name] = detail
                                    st.session_state.detail_info = detail
                                    st.rerun()
                                except Exception as e: st.error(f"❌ {e}")

                if page_end < total:
                    st.divider()
                    remain = min(BATCH_SIZE, total-page_end)
                    if st.button(f"➕ 추가 {remain}개 더 보기 ({page_end+1}~{min(page_end+BATCH_SIZE,total)}번)", use_container_width=True):
                        st.session_state.page_index += 1; st.rerun()
                else:
                    st.caption(f"✅ 전체 {total}개 식당을 모두 조회했습니다.")

    with tab2:
        st.markdown("#### 🔍 식당명으로 직접 검색")
        restaurant_name = st.text_input("식당명", placeholder="예) 진주회관  /  홍콩, Tim Ho Wan", key="direct_input", label_visibility="collapsed")
        if st.button("🔍 정보 조회", use_container_width=True, type="primary", key="direct_btn"):
            if restaurant_name:
                # C) 캐시 확인
                if restaurant_name in st.session_state.detail_cache:
                    st.session_state.detail_info = st.session_state.detail_cache[restaurant_name]
                    st.rerun()
                else:
                    with st.spinner(f"'{restaurant_name}' 정보 조사 중..."):
                        try:
                            detail = get_detail_info(restaurant_name)
                            st.session_state.detail_cache[restaurant_name] = detail
                            st.session_state.detail_info = detail
                        except Exception as e: st.error(f"❌ {e}")
            else: st.warning("식당명을 입력해주세요.")

    if st.session_state.detail_info:
        st.divider()
        show_detail(st.session_state.detail_info)
        st.divider()
        col1, col2 = st.columns([1,1])
        with col1:
            if st.button("💾 엑셀에 저장", type="primary", use_container_width=True):
                try:
                    append_to_excel(st.session_state.detail_info)
                    st.success(f"✅ '{st.session_state.detail_info.get('식당명','')}' 저장 완료!")
                    if os.path.exists(EXCEL_FILE):
                        with open(EXCEL_FILE, "rb") as f:
                            st.download_button("📥 엑셀 다운로드", data=f, file_name=EXCEL_FILE,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                except Exception as e: st.error(f"저장 오류: {e}")
        with col2:
            if st.button("✖️ 상세정보 닫기", use_container_width=True):
                st.session_state.detail_info = None; st.rerun()

if __name__ == "__main__":
    main()
