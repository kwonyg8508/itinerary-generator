import streamlit as st
from datetime import datetime, timedelta
from serpapi import GoogleSearch

# 도시명 -> IATA 공항 코드 매핑
AIRPORT_CODES = {
    "인천": "ICN", "프라하": "PRG", "이스탄불": "IST", 
    "부다페스트": "BUD", "비엔나": "VIE", "마드리드": "MAD", 
    "바르셀로나": "BCN", "로마": "FCO", "밀라노": "MXP",
    "파리": "CDG", "런던": "LHR", "프랑크푸르트": "FRA"
}

def search_google_flights(dep_city, in_city, out_date, ret_date):
    """
    SerpApi를 이용해 구글 플라이트 실시간 스케줄을 가져옵니다.
    """
    serpapi_key = st.secrets.get("SERPAPI_API_KEY", "")
    if not serpapi_key:
        st.error("Streamlit Secrets에 'SERPAPI_API_KEY'를 먼저 등록해 주세요.")
        return []

    dep_code = AIRPORT_CODES.get(dep_city, "ICN")
    arr_code = AIRPORT_CODES.get(in_city, in_city)

    params = {
        "engine": "google_flights",
        "departure_id": dep_code,
        "arrival_id": arr_code,
        "outbound_date": out_date.strftime("%Y-%m-%d"),
        "return_date": ret_date.strftime("%Y-%m-%d"),
        "currency": "KRW",
        "hl": "ko",
        "api_key": serpapi_key
    }

    try:
        search = GoogleSearch(params)
        res = search.get_dict()
        flights = res.get("best_flights", []) + res.get("other_flights", [])
        
        parsed = []
        for f in flights[:5]:  # 상위 5개 추천
            legs = f.get("flights", [])
            if not legs:
                continue
            first_leg = legs[0]
            airline = first_leg.get("airline", "외항사")
            f_num = f"{first_leg.get('airline_code', '')}{first_leg.get('flight_number', '')}"
            dep_t = first_leg.get("departure_airport", {}).get("time", "")
            arr_t = first_leg.get("arrival_airport", {}).get("time", "")
            
            summary = f"[{airline}] {f_num} {dep_code} {dep_t} ➔ {arr_code} {arr_t}"
            parsed.append({"airline": airline, "summary": summary, "flight_num": f_num})
        return parsed
    except Exception as e:
        st.error(f"항공 조회 실패: {e}")
        return []

def render_flight_search_ui():
    """
    Streamlit 상단에 표시될 항공 검색 박스 UI 및 자동 주입 버튼
    """
    with st.expander("✈️ [Add-on] 대리점 문의용 실시간 항공 스케줄 조회 & 자동 주입", expanded=False):
        c1, c2, c3, c4 = st.columns([1.5, 1, 1.2, 1.2])
        with c1:
            base_date = st.date_input("출발 기준일(가계약일)", datetime.now() + timedelta(days=60))
        with c2:
            tour_days = st.number_input("행사 기간(일)", min_value=3, max_value=25, value=9)
        with c3:
            in_city = st.selectbox("IN 도시", list(AIRPORT_CODES.keys())[1:], index=0)
        with c4:
            out_city = st.selectbox("OUT 도시", list(AIRPORT_CODES.keys())[1:], index=2)

        ret_date = base_date + timedelta(days=tour_days - 1)
        st.caption(f"🗓️ 확정 일정: **{base_date.strftime('%Y-%m-%d')} ~ {ret_date.strftime('%Y-%m-%d')} ({tour_days}일)**")

        if st.button("🔎 구글 플라이트 실시간 스케줄 조회"):
            with st.spinner("실시간 운항 스케줄을 조회 중입니다..."):
                results = search_google_flights("인천", in_city, base_date, ret_date)
                st.session_state["flight_results"] = results

        if "flight_results" in st.session_state and st.session_state["flight_results"]:
            st.write("---")
            for idx, item in enumerate(st.session_state["flight_results"]):
                ci, cb = st.columns([4, 1])
                with ci:
                    st.info(f"**옵션 {idx+1}:** {item['summary']}")
                with cb:
                    if st.button("👉 일정표에 주입", key=f"inj_btn_{idx}"):
                        d1 = f"1. 인천 출발 / {item['summary']} 도착 후 호텔 투숙\n"
                        mid = "2. 시내 핵심 관광지 투어 진행\n3. 전일 관광 및 이동\n"
                        last = f"{tour_days-1}. 시내 관광 후 공항 이동 / 기내 숙박\n{tour_days}. 인천 도착 후 해산"
                        st.session_state["routing_text_injected"] = d1 + mid + last
                        st.success("✅ [옵션 B] 루팅 창에 자동 주입되었습니다! 아래를 확인하세요.")
