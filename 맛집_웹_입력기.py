"""
🍽️  맛집 정보 웹 입력기 - Streamlit + Gemini API
비밀번호 로그인 + 식당 정보 조회 + 엑셀 자동 저장

실행방법:
  pip install streamlit google-generativeai pandas openpyxl
  streamlit run 맛집_웹_입력기.py
"""

import streamlit as st
import google.generativeai as genai
import pandas as pd
import json
import os
from datetime import date

# ──────────────────────────────────────────────────────
# ✅ 설정 - 여기만 수정하세요
# ──────────────────────────────────────────────────────
GEMINI_API_KEY = "AIzaSyCk7occLrHme-cl0UQrl-FuJkJqFJgDrlw"
ACCESS_PASSWORD = "park8520@@"   # 예) "matzip2024"
EXCEL_FILE  = "맛집정보.xlsx"
SHEET_NAME  = "한국"

# ──────────────────────────────────────────────────────
# 페이지 설정
# ──────────────────────────────────────────────────────
st.set_page_config(
    page_title="맛집 정보 입력기",
    page_icon="🍽️",
    layout="wide"
)

# ──────────────────────────────────────────────────────
# CSS 스타일
# ──────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-title { font-size: 2rem; font-weight: 700; margin-bottom: 0.2rem; }
    .subtitle   { color: #888; margin-bottom: 2rem; }
    .info-box   { background: #f8f9fa; border-radius: 10px; padding: 1.2rem; margin: 0.5rem 0; }
    .tag        { display: inline-block; background: #e8f4fd; color: #1a6fa8;
                  border-radius: 20px; padding: 2px 12px; font-size: 0.8rem; margin: 2px; }
    .tag-green  { background: #e8f8ee; color: #1a7a3f; }
    .tag-orange { background: #fff3e0; color: #b35c00; }
    .section-title { font-size: 1rem; font-weight: 600; color: #444;
                     border-left: 3px solid #ff6b35; padding-left: 8px; margin: 1rem 0 0.5rem; }
    .login-box  { max-width: 400px; margin: 8rem auto; text-align: center; }
</style>
""", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────
# 로그인 처리
# ──────────────────────────────────────────────────────
def check_login():
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False

    if not st.session_state.logged_in:
        st.markdown('<div class="login-box">', unsafe_allow_html=True)
        st.markdown("## 🍽️ 맛집 정보 입력기")
        st.markdown("접속하려면 비밀번호를 입력하세요.")
        pw = st.text_input("비밀번호", type="password", placeholder="비밀번호 입력")
        if st.button("로그인", use_container_width=True):
            if pw == ACCESS_PASSWORD:
                st.session_state.logged_in = True
                st.rerun()
            else:
                st.error("❌ 비밀번호가 올바르지 않습니다.")
        st.markdown('</div>', unsafe_allow_html=True)
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

식당명: {restaurant_name}

규칙:
- 한국 식당: 모든 항목 한글로 작성 (필요시 영어 병기)
- 영어권 외국 식당: 한글/영어 병기
- 영어권 외 외국 식당: 한글/영어/현지어 병기
- 정보가 없는 항목은 빈 문자열("") 처리
- 가격은 현지 통화 기준으로 작성

{{
  "식당명": "한글명 (영어명 / 현지어명)",
  "음식국적": "한식 / 일식 / 중화 / 양식 / 이태리 / 프랑스 / 태국 / 베트남 등",
  "카테고리": "육류 / 해산물 / 면류 / 밥류 / 빵간식 / 기타 등 세부 카테고리",
  "기준메뉴": "주재료 분류 (예: 소고기, 돼지고기, 해수어 등)",

  "추천사유": {{
    "미슐랭": "3스타(2023년) / 2스타 / 1스타 / 빕구르망 / 셀렉티드 / 해당없음",
    "블루리본": "블루리본 / 레드리본 / 해당없음",
    "로컬인증": "해당국가 맛집 인증 및 선정연도 (없으면 빈 문자열)",
    "노포": "창업년도 (예: 1965년 창업 / 해당없음)",
    "백년가게": "Y / N",
    "방송_생활의달인": "Y / N",
    "방송_수요미식회": "Y / N",
    "방송_식객허영만": "Y / N",
    "방송_백종원": "Y / N",
    "방송_전현무": "Y / N",
    "방송_정용진": "Y / N",
    "방송_흑백요리사": "Y / N",
    "방송_기타": "기타 방송 출연 내용 (없으면 빈 문자열)"
  }},

  "signature1": {{
    "메뉴명": "한글명 (영어명 / 현지어명)",
    "주재료": "주재료 한글",
    "요리방법": "요리방법 한글 (예: 구이, 찜, 볶음, 튀김 등)",
    "설명": "메뉴 설명 및 특징 (한글)",
    "가격": "현지 가격 (예: 15,000원 / 120HKD)"
  }},
  "signature2": {{
    "메뉴명": "한글명 (영어명 / 현지어명)",
    "주재료": "주재료 한글",
    "요리방법": "요리방법 한글",
    "설명": "메뉴 설명 및 특징 (한글)",
    "가격": "현지 가격"
  }},
  "signature3": {{
    "메뉴명": "한글명 (영어명 / 현지어명)",
    "주재료": "주재료 한글",
    "요리방법": "요리방법 한글",
    "설명": "메뉴 설명 및 특징 (한글)",
    "가격": "현지 가격"
  }},

  "국가": "한글 (영어 / 현지어)",
  "지역": "시도 및 세부지역 (예: 경기도 의정부시 / 홍콩 센트럴지역)",
  "주소": "한글주소 (영어주소 / 현지어주소)",
  "MRT역": "근처 MRT/지하철역명 및 출구번호 (해당없으면 빈 문자열)",
  "MRT도보": "MRT에서 식당까지 도보 이동 방법 한글 (해당없으면 빈 문자열)",
  "국가번호": "숫자만 (예: 82)",
  "지역번호": "숫자만 (예: 2)",
  "전화번호": "전화번호",
  "예약여부": "필수예약 / 권장 / 불필요",
  "대기피하는시간": "대기 피할 수 있는 시간대 (한글)",
  "팁및정보": "식당 관련 유용한 팁 및 정보 (한글)",
  "영업시간": "영업시간",
  "정기휴무": "휴무일",
  "가격대": "~10K / 10K~20K / 20K~40K / 40K~60K / 60K~80K / 80K~",
  "Homepage": "공식 홈페이지 URL (없으면 빈 문자열)"
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
    tr   = info.get("추천사유", {})

    def yn(val): return "*" if str(val).strip().upper() in ("Y","YES","*") else ""

    new_row = {
        "프렌차이즈":       "",
        "식당명":           info.get("식당명", ""),
        "국적":             info.get("음식국적", ""),
        "카테고리":         info.get("카테고리", ""),
        "기준메뉴":         info.get("기준메뉴", ""),
        "Signature\n메뉴1": sig1.get("메뉴명", ""),
        "Signature\n메뉴2": sig2.get("메뉴명", ""),
        "국가":             info.get("국가", ""),
        "시도":             info.get("지역", "").split()[0] if info.get("지역") else "",
        "주소1":            info.get("지역", ""),
        "주소2":            "",
        "주소3":            "",
        "신주소1":          info.get("주소", ""),
        "국가번호":         info.get("국가번호", ""),
        "지역번호":         info.get("지역번호", ""),
        "전화번호":         info.get("전화번호", ""),
        "Remark":           info.get("팁및정보", ""),
        "방문여부":         "N",
        "구정":             "",
        "미국이모":         "",
        "식당시설":         "",
        "주차":             "",
        "맛":               "",
        "가격":             info.get("가격대", ""),
        "노포":             tr.get("노포", "").replace("해당없음",""),
        "백년가게":         yn(tr.get("백년가게","")),
        "미쉐린가이드":     tr.get("미슐랭", "").replace("해당없음",""),
        "생활의달인":       yn(tr.get("방송_생활의달인","")),
        "블루리본":         tr.get("블루리본","").replace("해당없음",""),
        "수요미식회":       yn(tr.get("방송_수요미식회","")),
        "식객허영만":       yn(tr.get("방송_식객허영만","")),
        "백종원":           yn(tr.get("방송_백종원","")),
        "정용진":           yn(tr.get("방송_정용진","")),
        "전현무":           yn(tr.get("방송_전현무","")),
        "유튜브":           "",
        "지인":             "",
        "영업시간":         info.get("영업시간", ""),
        "정기휴무":         info.get("정기휴무", ""),
        "B/T":              "",
        "L/O":              "",
        "입력날짜":         date.today(),
        "지도등록":         "",
        "Homepage":         info.get("Homepage", ""),
    }

    if os.path.exists(EXCEL_FILE):
        df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    else:
        df = pd.DataFrame([new_row])

    df.to_excel(EXCEL_FILE, sheet_name=SHEET_NAME, index=False)


# ──────────────────────────────────────────────────────
# 결과 화면 표시
# ──────────────────────────────────────────────────────
def show_result(info: dict):
    tr   = info.get("추천사유", {})
    sig1 = info.get("signature1", {})
    sig2 = info.get("signature2", {})
    sig3 = info.get("signature3", {})

    def yn(val): return str(val).strip().upper() in ("Y","YES","*")

    st.markdown(f"## 🍽️ {info.get('식당명','')}")
    st.markdown(f"`{info.get('음식국적','')}` `{info.get('카테고리','')}` `{info.get('기준메뉴','')}`")
    st.divider()

    # 추천사유
    badges = []
    if tr.get("미슐랭") and "해당없음" not in tr.get("미슐랭",""):
        badges.append(f"⭐ {tr['미슐랭']}")
    if tr.get("블루리본") and "해당없음" not in tr.get("블루리본",""):
        badges.append(f"🎗️ {tr['블루리본']}")
    if tr.get("로컬인증"):
        badges.append(f"🏆 {tr['로컬인증']}")
    if tr.get("노포") and "해당없음" not in tr.get("노포",""):
        badges.append(f"🏛️ {tr['노포']}")

    방송목록 = [
        ("생활의달인","방송_생활의달인"),("수요미식회","방송_수요미식회"),
        ("식객 허영만","방송_식객허영만"),("백종원","방송_백종원"),
        ("전현무","방송_전현무"),("정용진","방송_정용진"),("흑백요리사","방송_흑백요리사"),
    ]
    방송출연 = [n for n,k in 방송목록 if yn(tr.get(k,""))]
    if 방송출연:
        badges.append(f"📺 {' / '.join(방송출연)}")
    if tr.get("방송_기타"):
        badges.append(f"📺 {tr['방송_기타']}")

    if badges:
        st.markdown('<p class="section-title">📌 추천사유</p>', unsafe_allow_html=True)
        tag_html = " ".join([f'<span class="tag">{b}</span>' for b in badges])
        st.markdown(tag_html, unsafe_allow_html=True)

    # 시그니처 메뉴
    st.markdown('<p class="section-title">🍴 시그니처 메뉴</p>', unsafe_allow_html=True)
    cols = st.columns(3)
    for i, (label, sig) in enumerate([("Signature 1", sig1), ("Signature 2", sig2), ("Signature 3", sig3)]):
        with cols[i]:
            if sig.get("메뉴명"):
                st.markdown(f'<div class="info-box">'
                            f'<b>{label}</b><br>'
                            f'<span style="font-size:1rem;font-weight:600">{sig["메뉴명"]}</span><br>'
                            f'<small>주재료: {sig.get("주재료","")} | {sig.get("요리방법","")}</small><br>'
                            f'<small>{sig.get("설명","")}</small><br>'
                            f'<b style="color:#ff6b35">{sig.get("가격","")}</b>'
                            f'</div>', unsafe_allow_html=True)
            else:
                st.markdown('<div class="info-box" style="color:#ccc;text-align:center">—</div>',
                            unsafe_allow_html=True)

    # 위치 정보
    col1, col2 = st.columns(2)
    with col1:
        st.markdown('<p class="section-title">📍 위치 정보</p>', unsafe_allow_html=True)
        st.markdown(f'<div class="info-box">'
                    f'🌏 <b>국가:</b> {info.get("국가","")}<br>'
                    f'📌 <b>지역:</b> {info.get("지역","")}<br>'
                    f'🏠 <b>주소:</b> {info.get("주소","")}<br>'
                    f'🚇 <b>MRT역:</b> {info.get("MRT역","") or "해당없음"}<br>'
                    f'🚶 <b>도보:</b> {info.get("MRT도보","") or "—"}'
                    f'</div>', unsafe_allow_html=True)

    with col2:
        st.markdown('<p class="section-title">📋 운영 정보</p>', unsafe_allow_html=True)
        국가번호 = info.get("국가번호","")
        지역번호 = info.get("지역번호","")
        전화번호 = info.get("전화번호","")
        st.markdown(f'<div class="info-box">'
                    f'📞 <b>전화:</b> +{국가번호}-{지역번호}-{전화번호}<br>'
                    f'🕐 <b>영업시간:</b> {info.get("영업시간","")}<br>'
                    f'🚫 <b>정기휴무:</b> {info.get("정기휴무","")}<br>'
                    f'📅 <b>예약:</b> {info.get("예약여부","")}<br>'
                    f'⏰ <b>대기 피하는 시간:</b> {info.get("대기피하는시간","")}'
                    f'</div>', unsafe_allow_html=True)

    # 팁 & 홈페이지
    if info.get("팁및정보"):
        st.markdown('<p class="section-title">💡 팁 & 정보</p>', unsafe_allow_html=True)
        st.info(info.get("팁및정보",""))

    if info.get("Homepage"):
        st.markdown(f'🌐 홈페이지: [{info["Homepage"]}]({info["Homepage"]})')


# ──────────────────────────────────────────────────────
# 메인 앱
# ──────────────────────────────────────────────────────
def main():
    check_login()

    # 상단 헤더
    col1, col2 = st.columns([8, 1])
    with col1:
        st.markdown('<p class="main-title">🍽️ 맛집 정보 입력기</p>', unsafe_allow_html=True)
        st.markdown('<p class="subtitle">식당명을 입력하면 Gemini AI가 자동으로 정보를 조사합니다</p>',
                    unsafe_allow_html=True)
    with col2:
        if st.button("로그아웃"):
            st.session_state.logged_in = False
            st.rerun()

    st.divider()

    # 검색 입력
    col_input, col_btn = st.columns([5, 1])
    with col_input:
        restaurant_name = st.text_input(
            "식당명",
            placeholder="예) 진주회관, 스시 사이토, Tim Ho Wan...",
            label_visibility="collapsed"
        )
    with col_btn:
        search_btn = st.button("🔍 조회", use_container_width=True, type="primary")

    # 조회 실행
    if search_btn and restaurant_name:
        with st.spinner(f"'{restaurant_name}' 정보 조사 중... (10~30초 소요)"):
            try:
                info = query_restaurant(restaurant_name)
                st.session_state["last_result"] = info
                st.session_state["last_name"]   = restaurant_name
            except json.JSONDecodeError:
                st.error("❌ 응답 파싱 오류입니다. 다시 시도해주세요.")
                st.stop()
            except Exception as e:
                st.error(f"❌ 오류 발생: {e}")
                st.stop()

    # 결과 표시
    if "last_result" in st.session_state:
        info = st.session_state["last_result"]
        show_result(info)

        st.divider()

        # 저장 버튼
        col_save, col_skip = st.columns([1, 4])
        with col_save:
            if st.button("💾 엑셀에 저장", type="primary", use_container_width=True):
                try:
                    append_to_excel(info)
                    st.success(f"✅ '{info.get('식당명','')}' 정보가 맛집정보.xlsx에 저장되었습니다!")
                    # 저장된 파일 다운로드 버튼
                    if os.path.exists(EXCEL_FILE):
                        with open(EXCEL_FILE, "rb") as f:
                            st.download_button(
                                "📥 엑셀 파일 다운로드",
                                data=f,
                                file_name=EXCEL_FILE,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                except Exception as e:
                    st.error(f"저장 오류: {e}")


if __name__ == "__main__":
    main()
