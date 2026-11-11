import streamlit as st
from datetime import datetime, timedelta
from serpapi import GoogleSearch

EUROPE_AIRPORTS = {
    "프라하 (PRG)": "PRG", "부다페스트 (BUD)": "BUD", "비엔나 (VIE)": "VIE",
    "자그레브 (ZAG)": "ZAG", "두브로브니크 (DBV)": "DBV", "스플리트 (SPU)": "SPU",
    "바르샤바 (WAW)": "WAW", "크라쿠프 (KRK)": "KRK", "류블랴나 (LJU)": "LJU",
    "파리 샤를드골 (CDG)": "CDG", "런던 히드로 (LHR)": "LHR", "암스테르담 (AMS)": "AMS",
    "마드리드 (MAD)": "MAD", "바르셀로나 (BCN)": "BCN", "리스본 (LIS)": "LIS",
    "로마 피우미치노 (FCO)": "FCO", "밀라노 말펜사 (MXP)": "MXP", "베네치아 (VCE)": "VCE",
    "아테네 (ATH)": "ATH", "프랑크푸르트 (FRA)": "FRA", "뮌헨 (MUC)": "MUC",
    "취리히 (ZRH)": "ZRH", "코펜하겐 (CPH)": "CPH", "헬싱키 (HEL)": "HEL",
    "오슬로 (OSL)": "OSL", "스톡홀름 (ARN)": "ARN", "레이캬비크 (KEF)": "KEF",
    "이스탄불 신공항 (IST)": "IST", "두바이 (DXB)": "DXB", "✏️ 공항코드 직접 입력": "CUSTOM"
}

def normalize_airline_name(name: str) -> str:
    c = name.replace(" ", "").upper()
    if any(k in c for k in ["EMIRATES", "에미레이트"]): return "에미레이트항공"
    if any(k in c for k in ["FINNAIR", "핀에어"]): return "핀에어"
    if any(k in c for k in ["TURKISH", "터키", "THY"]): return "터키항공"
    if any(k in c for k in ["LUFTHANSA", "루프트한자"]): return "루프트한자"
    if any(k in c for k in ["KOREAN", "대한항공"]): return "대한항공"
    if any(k in c for k in ["ASIANA", "아시아나"]): return "아시아나항공"
    return name.strip()

def search_single_leg(dep_code, arr_code, flight_date):
    serpapi_key = st.secrets.get("SERPAPI_API_KEY", "")
    if not serpapi_key:
        st.error("❌ Streamlit Secrets에 'SERPAPI_API_KEY'가 설정되어 있지 않습니다.")
        return []

    params = {
        "engine": "google_flights",
        "departure_id": dep_code,
        "arrival_id": arr_code,
        "outbound_date": flight_date.strftime("%Y-%m-%d"),
        "type": "2",
        "currency": "KRW",
        "hl": "ko",
        "gl": "kr",
        "api_key": serpapi_key
    }

    try:
        search = GoogleSearch(params)
        res = search.get_dict()
        flights = res.get("best_flights", []) + res.get("other_flights", [])
        if not flights: return []

        parsed = []
        for f in flights[:20]:
            legs = f.get("flights", [])
            if not legs: continue
            if len(legs) > 2: continue
            
            if len(legs) == 2:
                arr_first = legs[0].get("arrival_airport", {}).get("time", "")
                dep_second = legs[1].get("departure_airport", {}).get("time", "")
                try:
                    t1 = datetime.strptime(arr_first, "%Y-%m-%d %H:%M")
                    t2 = datetime.strptime(dep_second, "%Y-%m-%d %H:%M")
                    layover_mins = (t2 - t1).total_seconds() / 60
                    if layover_mins < 60 or layover_mins > 300:
                        continue
                except Exception:
                    pass

            first_leg = legs[0]
            raw_airline = first_leg.get("airline", "항공사")
            norm_airline = normalize_airline_name(raw_airline)
            
            dep_airport = first_leg.get("departure_airport", {}).get("id", dep_code)
            dep_t = first_leg.get("departure_airport", {}).get("time", "")
            
            last_leg = legs[-1]
            arr_airport = last_leg.get("arrival_airport", {}).get("id", arr_code)
            arr_t = last_leg.get("arrival_airport", {}).get("time", "")
            
            f_nums = " / ".join([f"{l.get('airline_code', '')}{l.get('flight_number', '')}" for l in legs if l.get('flight_number')])
            if not f_nums: f_nums = norm_airline

            is_transit = len(legs) > 1
            via_text = f"({len(legs)-1}회 경유)" if is_transit else "(직항)"

            summary = f"[{norm_airline}] {f_nums} | {dep_airport} {dep_t} ➔ {arr_airport} {arr_t} {via_text}"
            
            parsed.append({
                "norm_airline": norm_airline,
                "flight_nums": f_nums,
                "summary": summary,
                "dep_airport": dep_airport,
                "arr_airport": arr_airport,
                "dep_time": dep_t,
                "arr_time": arr_t,
                "legs": legs
            })
        return parsed
    except Exception as e:
        st.error(f"항공 조회 오류 ({dep_code} ➔ {arr_code}): {e}")
        return []

def render_flight_search_ui():
    with st.expander("✈️ [Add-on] 유럽 전 지역 실시간 항공 스케줄 조회 & 자동 주입", expanded=True):
        c1, c2 = st.columns([1.5, 1])
        with c1: base_date = st.date_input("출발일 (IN)", datetime.now() + timedelta(days=60))
        with c2: tour_days = st.number_input("행사 기간 (일수)", min_value=3, max_value=25, value=11)

        c3, c4 = st.columns(2)
        with c3:
            in_sel = st.selectbox("IN 도시 선택", list(EUROPE_AIRPORTS.keys()), index=0)
            in_code = st.text_input("IN 공항 코드", "OSL").upper().strip() if EUROPE_AIRPORTS[in_sel] == "CUSTOM" else EUROPE_AIRPORTS[in_sel]
            in_city_label = in_sel.split(" (")[0]
        with c4:
            out_sel = st.selectbox("OUT 도시 선택", list(EUROPE_AIRPORTS.keys()), index=1)
            out_code = st.text_input("OUT 공항 코드", "VNO").upper().strip() if EUROPE_AIRPORTS[out_sel] == "CUSTOM" else EUROPE_AIRPORTS[out_sel]
            out_city_label = out_sel.split(" (")[0]

        ret_fly_date = base_date + timedelta(days=tour_days - 2)

        if st.button("🔎 동일 항공사 왕복 스케줄 실시간 조회", type="primary"):
            with st.spinner("최적 환승(1회 경유, 3~4시간 컷) 항공편을 매칭 중입니다..."):
                outbound_list = search_single_leg("ICN", in_code, base_date)
                inbound_list = search_single_leg(out_code, "ICN", ret_fly_date)
                
                matched_pairs = []
                seen_airlines = set()
                for ob in outbound_list:
                    o_air = ob["norm_airline"]
                    same_rets = [ib for ib in inbound_list if ib["norm_airline"] == o_air]
                    if same_rets and o_air not in seen_airlines:
                        matched_pairs.append({"airline": o_air, "out": ob, "ret": same_rets[0]})
                        seen_airlines.add(o_air)
                st.session_state["flight_pairs"] = matched_pairs

        if "flight_pairs" in st.session_state:
            pairs = st.session_state["flight_pairs"]
            if not pairs:
                st.warning("⚠️ 조건에 맞는 동일 항공사 왕복 스케줄을 찾지 못했습니다.")
            else:
                st.write("---")
                st.markdown(f"#### 🎯 출귀국 동일 항공사 왕복 스케줄 ({len(pairs)}개 항공사)")
                for idx, pair in enumerate(pairs[:5]):
                    ob, ib, airline_name = pair["out"], pair["ret"], pair["airline"]
                    ci, cb = st.columns([4, 1])
                    with ci:
                        st.info(f"**옵션 {idx+1} [{airline_name}] ⭐**\n\n• **출국:** {ob['summary']}\n\n• **귀국:** {ib['summary']}")
                    with cb:
                        if st.button("👉 일정표 주입", key=f"inj_pair_{idx}"):
                            # 🚨 [수정 포인트] 화면에 표시된 정확한 객체 데이터를 기반으로 루팅 텍스트 생성
                            ob_legs = ob.get("legs", [])
                            ib_legs = ib.get("legs", [])
                            
                            # 출국 루팅 세부 생성
                            out_lines = []
                            for l_idx, leg in enumerate(ob_legs):
                                dep_a = leg.get("departure_airport", {}).get("id", "")
                                arr_a = leg.get("arrival_airport", {}).get("id", "")
                                dep_t = leg.get("departure_airport", {}).get("time", "")
                                arr_t = leg.get("arrival_airport", {}).get("time", "")
                                f_num = f"{leg.get('airline_code', '')}{leg.get('flight_number', '')}"
                                
                                if l_idx == 0:
                                    out_lines.append(f"1. 인천 국제공항 미팅 후 출국 / [{airline_name}] {f_num} 인천 {dep_t} ➔ {arr_a} {arr_t} 출발")
                                else:
                                    out_lines.append(f"   - {dep_a} {dep_t} 출발 ➔ {arr_a} {arr_t} 도착 ({f_num})")

                            # 귀국 루팅 세부 생성
                            ret_lines = []
                            for l_idx, leg in enumerate(ib_legs):
                                dep_a = leg.get("departure_airport", {}).get("id", "")
                                arr_a = leg.get("arrival_airport", {}).get("id", "")
                                dep_t = leg.get("departure_airport", {}).get("time", "")
                                arr_t = leg.get("arrival_airport", {}).get("time", "")
                                f_num = f"{leg.get('airline_code', '')}{leg.get('flight_number', '')}"
                                
                                if l_idx == 0:
                                    ret_lines.append(f"{tour_days-1}. {out_city_label} 공항 이동 / [{airline_name}] {f_num} {dep_a} {dep_t} ➔ {arr_a} {arr_t} / 기내 숙박")
                                else:
                                    ret_lines.append(f"   - {dep_a} {dep_t} 출발 ➔ {arr_a} {arr_t} 도착 ({f_num})")

                            injected_full_text = "\n".join(out_lines) + f"\n2. {in_city_label} 도착 후 가이드 미팅 및 투어 진행\n3. 전일 관광 및 이동\n" + "\n".join(ret_lines) + f"\n{tour_days}. 인천 국제공항 도착 후 해산"

                            st.session_state["routing_text_injected"] = injected_full_text
                            st.session_state["selected_input_mode"] = "[옵션 B] 직접 루팅 입력 (정밀 항공/일정 파싱)"
                            st.success(f"✅ [{airline_name}] 정확한 실시간 스케줄이 옵션 B에 주입되었습니다!")
                            st.rerun()
