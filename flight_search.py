import streamlit as st
from datetime import datetime, timedelta
from serpapi import GoogleSearch

# 도시명 -> IATA 공항 코드 매핑
AIRPORT_CODES = {
    "인천": "ICN", "프라하": "PRG", "이스탄불": "IST", 
    "부다페스트": "BUD", "비엔나": "VIE", "마드리드": "MAD", 
    "바르셀로나": "BCN", "로마": "FCO", "밀라노": "MXP",
    "파리": "CDG", "런던": "LHR", "프랑크푸르트": "FRA",
    "자그레브": "ZAG", "취리히": "ZRH", "리스본": "LIS", "아테네": "ATH"
}

def search_single_leg(dep_code, arr_code, flight_date):
    """
    구글 플라이트에서 특정 구간(편도)만 독립적으로 강제 조회
    """
    serpapi_key = st.secrets.get("SERPAPI_API_KEY", "")
    if not serpapi_key:
        st.error("❌ Streamlit Secrets에 'SERPAPI_API_KEY'가 설정되어 있지 않습니다.")
        return []

    params = {
        "engine": "google_flights",
        "departure_id": dep_code,
        "arrival_id": arr_code,
        "outbound_date": flight_date.strftime("%Y-%m-%d"),
        "type": "2",          # 2 = 순수 One-way (편도 강제)
        "currency": "KRW",
        "hl": "ko",
        "gl": "kr",
        "api_key": serpapi_key
    }

    try:
        search = GoogleSearch(params)
        res = search.get_dict()
        
        # 구글 플라이트 응답 파싱
        flights = res.get("best_flights", []) + res.get("other_flights", [])
        if not flights:
            return []

        parsed = []
        for f in flights[:6]:
            legs = f.get("flights", [])
            if not legs:
                continue
                
            first_leg = legs[0]
            airline = first_leg.get("airline", "항공사").strip()
            
            # 실제 출발 공항 및 시간
            dep_airport = first_leg.get("departure_airport", {}).get("id", dep_code)
            dep_t = first_leg.get("departure_airport", {}).get("time", "")
            
            # 실제 도착 공항 및 시간
            last_leg = legs[-1]
            arr_airport = last_leg.get("arrival_airport", {}).get("id", arr_code)
            arr_t = last_leg.get("arrival_airport", {}).get("time", "")
            
            # 편명 조합 (경유 포함)
            f_nums = " / ".join([f"{l.get('airline_code', '')}{l.get('flight_number', '')}" for l in legs if l.get('flight_number')])
            if not f_nums:
                f_nums = airline

            is_transit = len(legs) > 1
            via_text = f"({len(legs)-1}회 경유)" if is_transit else "(직항)"

            summary = f"[{airline}] {f_nums} | {dep_airport} {dep_t} ➔ {arr_airport} {arr_t} {via_text}"
            
            parsed.append({
                "airline": airline,
                "flight_nums": f_nums,
                "summary": summary,
                "dep_airport": dep_airport,
                "arr_airport": arr_airport,
                "dep_time": dep_t,
                "arr_time": arr_t
            })
        return parsed
    except Exception as e:
        st.error(f"항공 조회 오류 ({dep_code} ➔ {arr_code}): {e}")
        return []

def render_flight_search_ui():
    """Streamlit 상단 다구간 항공 검색 및 일정표 자동 주입 UI"""
    with st.expander("✈️ [Add-on] 대리점 문의용 다구간 항공 스케줄 조회 & 자동 주입", expanded=True):
        c1, c2, c3, c4 = st.columns([1.5, 1, 1.2, 1.2])
        with c1:
            # 기본 검색일: 오늘 기준 약 60일 뒤
            base_date = st.date_input("출발일 (IN)", datetime.now() + timedelta(days=60))
        with c2:
            tour_days = st.number_input("행사 기간 (일수)", min_value=3, max_value=25, value=9)
        with c3:
            city_names = list(AIRPORT_CODES.keys())[1:] # 인천 제외 목록
            in_city = st.selectbox("IN 도시 (도착)", city_names, index=city_names.index("프라하") if "프라하" in city_names else 0)
        with c4:
            out_city = st.selectbox("OUT 도시 (출발)", city_names, index=city_names.index("부다페스트") if "부다페스트" in city_names else 2)

        # 귀국 비행기 탑승일(8일차) 및 한국 도착일(9일차) 계산
        ret_fly_date = base_date + timedelta(days=tour_days - 2)
        
        st.caption(f"🗓️ 실시간 매칭 요청: **출국 {base_date.strftime('%Y-%m-%d')} (인천➔{in_city}) | 귀국탑승 {ret_fly_date.strftime('%Y-%m-%d')} ({out_city}➔인천)**")

        if st.button("🔎 다구간(In/Out) 실시간 스케줄 조회", type="primary"):
            dep_code = "ICN"
            in_code = AIRPORT_CODES.get(in_city, "PRG")
            out_code = AIRPORT_CODES.get(out_city, "BUD")
            
            with st.spinner(f"1구간(인천➔{in_city}) 및 2구간({out_city}➔인천)을 독립 조회하여 매칭 중입니다..."):
                outbound_list = search_single_leg(dep_code, in_code, base_date)
                inbound_list = search_single_leg(out_code, dep_code, ret_fly_date)
                
                if not outbound_list or not inbound_list:
                    st.warning("⚠️ 스케줄을 가져오지 못했습니다. 날짜가 미래인지, 해당 노선에 운항 편이 있는지 확인해 주세요.")
                    st.session_state["flight_pairs"] = []
                else:
                    # 동일 항공사 왕복 우선 매칭
                    matched_pairs = []
                    for ob in outbound_list:
                        same_rets = [ib for ib in inbound_list if ib["airline"] == ob["airline"]]
                        if same_rets:
                            matched_pairs.append({"out": ob, "ret": same_rets[0], "same": True})
                        else:
                            matched_pairs.append({"out": ob, "ret": inbound_list[0], "same": False})
                    
                    st.session_state["flight_pairs"] = matched_pairs

        # 검색 결과 카드 표시
        if "flight_pairs" in st.session_state and st.session_state["flight_pairs"]:
            st.write("---")
            st.markdown(f"#### 🎯 실시간 다구간 조합 결과: [IN: {in_city} / OUT: {out_city}]")
            
            for idx, pair in enumerate(st.session_state["flight_pairs"][:5]):
                ob = pair["out"]
                ib = pair["ret"]
                tag = "⭐ 동일항공사 왕복" if pair["same"] else "⚠️ 출귀국 항공사 상이"
                
                ci, cb = st.columns([4, 1])
                with ci:
                    st.info(
                        f"**옵션 {idx+1} [{ob['airline']}] {tag}**\n\n"
                        f"• **출국(1일차):** {ob['summary']}\n\n"
                        f"• **귀국({tour_days-1}일차):** {ib['summary']}"
                    )
                with cb:
                    if st.button("👉 일정표 주입", key=f"inj_pair_{idx}"):
                        d1 = f"1. 인천 출발 / {ob['summary']} 도착 후 호텔 투숙\n"
                        mid = f"2. {in_city} 시내 핵심 관광지 투어 진행\n3. 전일 관광 및 이동\n4. 이동 및 전일 관광\n"
                        d_last_prev = f"{tour_days-1}. {out_city} 시내 관광 후 공항 이동 / {ib['summary']} / 기내 숙박\n"
                        d_last = f"{tour_days}. 인천 국제공항 도착 후 해산"
                        
                        st.session_state["routing_text_injected"] = d1 + mid + d_last_prev + d_last
                        st.success(f"✅ 출국({in_city}) 및 귀국({out_city}) 스케줄이 아래 루팅 창에 주입되었습니다!")
                        st.rerun()
