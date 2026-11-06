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

def search_flight_leg(dep_code, arr_code, flight_date):
    """단일 구간(편도) 스케줄을 실시간 조회하여 상위 항공편 반환"""
    serpapi_key = st.secrets.get("SERPAPI_API_KEY", "")
    if not serpapi_key:
        return []

    params = {
        "engine": "google_flights",
        "departure_id": dep_code,
        "arrival_id": arr_code,
        "outbound_date": flight_date.strftime("%Y-%m-%d"),
        "type": "2",  # One-way 편도 모드 (다구간 출/귀국 분리용)
        "currency": "KRW",
        "hl": "ko",
        "api_key": serpapi_key
    }

    try:
        search = GoogleSearch(params)
        res = search.get_dict()
        flights = res.get("best_flights", []) + res.get("other_flights", [])
        
        parsed = []
        for f in flights[:8]:
            legs = f.get("flights", [])
            if not legs:
                continue
            first_leg = legs[0]
            airline = first_leg.get("airline", "항공사")
            dep_t = first_leg.get("departure_airport", {}).get("time", "")
            
            # 경유 편명 및 도착 처리
            f_nums = " / ".join([f"{l.get('airline_code', '')}{l.get('flight_number', '')}" for l in legs])
            last_leg = legs[-1]
            arr_t = last_leg.get("arrival_airport", {}).get("time", "")
            
            is_transit = len(legs) > 1
            via_text = f"({len(legs)-1}회 경유)" if is_transit else "(직항)"

            summary = f"[{airline}] {f_nums} | {dep_code} {dep_t} ➔ {arr_code} {arr_t} {via_text}"
            parsed.append({
                "airline": airline,
                "flight_nums": f_nums,
                "summary": summary,
                "dep_time": dep_t,
                "arr_time": arr_t
            })
        return parsed
    except Exception as e:
        st.error(f"항공 조회 실패: {e}")
        return []

def render_flight_search_ui():
    """Streamlit 상단 다구간 항공 검색 및 일정표 주입 UI"""
    with st.expander("✈️ [Add-on] 대리점 문의용 다구간 항공 스케줄 조회 & 자동 주입", expanded=False):
        c1, c2, c3, c4 = st.columns([1.5, 1, 1.2, 1.2])
        with c1:
            base_date = st.date_input("출발일 (IN)", datetime.now() + timedelta(days=60))
        with c2:
            tour_days = st.number_input("행사 기간 (일수)", min_value=3, max_value=25, value=9)
        with c3:
            in_city = st.selectbox("IN 도시 (도착)", list(AIRPORT_CODES.keys())[1:], index=0)
        with c4:
            out_city = st.selectbox("OUT 도시 (출발)", list(AIRPORT_CODES.keys())[1:], index=2)

        # 9일 행사 기준: 귀국 비행기 탑승일(8일차) 및 한국 도착일(9일차) 계산
        ret_fly_date = base_date + timedelta(days=tour_days - 2)
        end_date = base_date + timedelta(days=tour_days - 1)
        
        st.caption(f"🗓️ 확정 스케줄: **1일차 출발 {base_date.strftime('%Y-%m-%d')} (인천➔{in_city}) | {tour_days-1}일차 귀국 {ret_fly_date.strftime('%Y-%m-%d')} ({out_city}➔인천)**")

        if st.button("🔎 다구간(In/Out) 실시간 스케줄 조회", type="primary"):
            dep_code = "ICN"
            in_code = AIRPORT_CODES.get(in_city, "PRG")
            out_code = AIRPORT_CODES.get(out_city, "BUD")
            
            with st.spinner("출국편 및 귀국편 실시간 운항 데이터를 매칭 중입니다..."):
                outbound_list = search_flight_leg(dep_code, in_code, base_date)
                inbound_list = search_flight_leg(out_code, dep_code, ret_fly_date)
                
                # 동일 항공사 왕복 매칭 우선 정렬
                matched_pairs = []
                for ob in outbound_list:
                    # 동일 항공사 귀국편 우선 탐색
                    same_airline_rets = [ib for ib in inbound_list if ib["airline"] == ob["airline"]]
                    if same_airline_rets:
                        matched_pairs.append({"out": ob, "ret": same_airline_rets[0], "same": True})
                    elif inbound_list:
                        matched_pairs.append({"out": ob, "ret": inbound_list[0], "same": False})
                
                st.session_state["flight_pairs"] = matched_pairs

        # 결과 렌더링
        if "flight_pairs" in st.session_state and st.session_state["flight_pairs"]:
            st.write("---")
            st.markdown("#### 🎯 추천 다구간 왕복 조합 (동일 항공사 매칭)")
            
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
                        mid = "2. 시내 핵심 관광지 투어 진행\n3. 전일 관광 및 이동\n4. 이동 및 전일 관광\n"
                        d_last_prev = f"{tour_days-1}. 시내 관광 후 공항 이동 / {ib['summary']} / 기내 숙박\n"
                        d_last = f"{tour_days}. 인천 국제공항 도착 후 해산"
                        
                        st.session_state["routing_text_injected"] = d1 + mid + d_last_prev + d_last
                        st.success("✅ 출국(프라하) 및 귀국(부다페스트) 스케줄이 아래 루팅 창에 완벽히 주입되었습니다!")
