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

def normalize_airline_name(name: str) -> str:
    """항공사 영문/한글 표기 불일치 방지를 위한 정규화"""
    name_clean = name.replace(" ", "").upper()
    if "TURKISH" in name_clean or "터키" in name_clean or "THY" in name_clean:
        return "터키항공"
    if "LUFTHANSA" in name_clean or "루프트한자" in name_clean:
        return "루프트한자"
    if "KOREAN" in name_clean or "대한항공" in name_clean:
        return "대한항공"
    if "ASIANA" in name_clean or "아시아나" in name_clean:
        return "아시아나항공"
    if "EMIRATES" in name_clean or "에미레이트" in name_clean:
        return "에미레이트항공"
    if "QATAR" in name_clean or "카타르" in name_clean:
        return "카타르항공"
    if "AIRFRANCE" in name_clean or "에어프랑스" in name_clean:
        return "에어프랑스"
    if "KLM" in name_clean:
        return "KLM네덜란드항공"
    if "FINNAIR" in name_clean or "핀에어" in name_clean:
        return "핀에어"
    return name.strip()

def search_single_leg(dep_code, arr_code, flight_date):
    """구글 플라이트 특정 구간 편도 스케줄 조회"""
    serpapi_key = st.secrets.get("SERPAPI_API_KEY", "")
    if not serpapi_key:
        st.error("❌ Streamlit Secrets에 'SERPAPI_API_KEY'가 설정되어 있지 않습니다.")
        return []

    params = {
        "engine": "google_flights",
        "departure_id": dep_code,
        "arrival_id": arr_code,
        "outbound_date": flight_date.strftime("%Y-%m-%d"),
        "type": "2",  # 순수 One-way 편도 모드
        "currency": "KRW",
        "hl": "ko",
        "gl": "kr",
        "api_key": serpapi_key
    }

    try:
        search = GoogleSearch(params)
        res = search.get_dict()
        flights = res.get("best_flights", []) + res.get("other_flights", [])
        if not flights:
            return []

        parsed = []
        for f in flights[:12]:
            legs = f.get("flights", [])
            if not legs:
                continue
                
            first_leg = legs[0]
            raw_airline = first_leg.get("airline", "항공사")
            norm_airline = normalize_airline_name(raw_airline)
            
            dep_airport = first_leg.get("departure_airport", {}).get("id", dep_code)
            dep_t = first_leg.get("departure_airport", {}).get("time", "")
            
            last_leg = legs[-1]
            arr_airport = last_leg.get("arrival_airport", {}).get("id", arr_code)
            arr_t = last_leg.get("arrival_airport", {}).get("time", "")
            
            f_nums = " / ".join([f"{l.get('airline_code', '')}{l.get('flight_number', '')}" for l in legs if l.get('flight_number')])
            if not f_nums:
                f_nums = norm_airline

            is_transit = len(legs) > 1
            via_text = f"({len(legs)-1}회 경유)" if is_transit else "(직항)"

            summary = f"[{norm_airline}] {f_nums} | {dep_airport} {dep_t} ➔ {arr_airport} {arr_t} {via_text}"
            
            parsed.append({
                "raw_airline": raw_airline,
                "norm_airline": norm_airline,
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
            base_date = st.date_input("출발일 (IN)", datetime.now() + timedelta(days=60))
        with c2:
            tour_days = st.number_input("행사 기간 (일수)", min_value=3, max_value=25, value=9)
        with c3:
            city_names = list(AIRPORT_CODES.keys())[1:]
            in_city = st.selectbox("IN 도시 (도착)", city_names, index=city_names.index("프라하") if "프라하" in city_names else 0)
        with c4:
            out_city = st.selectbox("OUT 도시 (출발)", city_names, index=city_names.index("부다페스트") if "부다페스트" in city_names else 2)

        # 귀국 비행기 탑승일(8일차) 계산
        ret_fly_date = base_date + timedelta(days=tour_days - 2)
        
        st.caption(f"🗓️ 매칭 기준: **1일차 출발 {base_date.strftime('%Y-%m-%d')} (인천➔{in_city}) | {tour_days-1}일차 귀국 {ret_fly_date.strftime('%Y-%m-%d')} ({out_city}➔인천)**")

        if st.button("🔎 동일 항공사 다구간 왕복 조회", type="primary"):
            dep_code = "ICN"
            in_code = AIRPORT_CODES.get(in_city, "PRG")
            out_code = AIRPORT_CODES.get(out_city, "BUD")
            
            with st.spinner(f"출국(인천➔{in_city}) 및 귀국({out_city}➔인천) 동일 항공사 편을 매칭 중입니다..."):
                outbound_list = search_single_leg(dep_code, in_code, base_date)
                inbound_list = search_single_leg(out_code, dep_code, ret_fly_date)
                
                # 🚨 동일 항공사 100% 일치 건만 필터링
                matched_pairs = []
                seen_airlines = set()
                
                for ob in outbound_list:
                    o_air = ob["norm_airline"]
                    # 귀국편 목록에서 동일 항공사 찾기
                    same_rets = [ib for ib in inbound_list if ib["norm_airline"] == o_air]
                    
                    if same_rets and o_air not in seen_airlines:
                        # 동일 항공사 매칭 성공 건만 추가
                        matched_pairs.append({
                            "airline": o_air,
                            "out": ob,
                            "ret": same_rets[0]
                        })
                        seen_airlines.add(o_air)
                
                st.session_state["flight_pairs"] = matched_pairs

        # 검색 결과 렌더링
        if "flight_pairs" in st.session_state:
            pairs = st.session_state["flight_pairs"]
            if not pairs:
                st.warning("⚠️ 해당 날짜/구간에 출·귀국 항공사가 완전히 일치하는 왕복 스케줄이 없습니다. 날짜를 변경해 보세요.")
            else:
                st.write("---")
                st.markdown(f"#### 🎯 출귀국 동일 항공사 왕복 스케줄 ({len(pairs)}개 항공사 발견)")
                
                for idx, pair in enumerate(pairs[:5]):
                    ob = pair["out"]
                    ib = pair["ret"]
                    airline_name = pair["airline"]
                    
                    ci, cb = st.columns([4, 1])
                    with ci:
                        st.info(
                            f"**옵션 {idx+1} [{airline_name}] ⭐ 출귀국 동일 항공사 왕복**\n\n"
                            f"• **출국(1일차):** {ob['summary']}\n\n"
                            f"• **귀국({tour_days-1}일차):** {ib['summary']}"
                        )
                    with cb:
                        if st.button("👉 일정표 주입", key=f"inj_pair_{idx}"):
                            d1 = f"1. 인천 출발 / {ob['summary']} 도착 후 호텔 투숙\n"
                            mid = f"2. {in_city} 시내 핵심 관광지 투어 진행\n3. 전일 관광 및 이동\n4. 이동 및 전일 관광\n"
                            d_last_prev = f"{tour_days-1}. {out_city} 시내 관광 후 공항 이동 / {ib['summary']} / 기내 숙박\n"
                            d_last = f"{tour_days}. 인천 국제공항 ({ib['arr_time']}) 도착 후 해산"
                            
                            st.session_state["routing_text_injected"] = d1 + mid + d_last_prev + d_last
                            st.session_state["selected_input_mode"] = "[옵션 B] 직접 루팅 입력 (정밀 항공/일정 파싱)"
                            st.success(f"✅ [{airline_name}] 스케줄이 옵션 B에 자동 주입되었습니다!")
                            st.rerun()
