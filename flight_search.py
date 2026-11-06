import streamlit as st
from datetime import datetime, timedelta
from serpapi import GoogleSearch

# 유럽 주요 50개 공항 IATA 코드 데이터베이스
EUROPE_AIRPORTS = {
    # 동유럽 / 발칸
    "프라하 (PRG)": "PRG", "부다페스트 (BUD)": "BUD", "비엔나 (VIE)": "VIE",
    "자그레브 (ZAG)": "ZAG", "두브로브니크 (DBV)": "DBV", "스플리트 (SPU)": "SPU",
    "바르샤바 (WAW)": "WAW", "크라쿠프 (KRK)": "KRK", "류블랴나 (LJU)": "LJU",
    
    # 서유럽 / 영국 / 아일랜드
    "파리 샤를드골 (CDG)": "CDG", "파리 오를리 (ORY)": "ORY",
    "런던 히드로 (LHR)": "LHR", "런던 개트윅 (LGW)": "LGW",
    "더블린 (DUB)": "DUB", "암스테르담 (AMS)": "AMS", "브뤼셀 (BRU)": "BRU",
    
    # 지중해 / 스페인 / 포르투갈
    "마드리드 (MAD)": "MAD", "바르셀로나 (BCN)": "BCN", "말라가 (AGP)": "AGP",
    "리스본 (LIS)": "LIS", "포르투 (OPO)": "OPO",
    
    # 이탈리아 / 몰타 / 그리스
    "로마 피우미치노 (FCO)": "FCO", "밀라노 말펜사 (MXP)": "MXP", "베네치아 (VCE)": "VCE",
    "피렌체 (FLR)": "FLR", "나폴리 (NAP)": "NAP", "아테네 (ATH)": "ATH",
    "테살로니키 (SKG)": "SKG", "산토리니 (JTR)": "JTR", "몰타 (MLA)": "MLA",
    
    # 독일 / 스위스 / 오스트리아
    "프랑크푸르트 (FRA)": "FRA", "뮌헨 (MUC)": "MUC", "베를린 (BER)": "BER",
    "취리히 (ZRH)": "ZRH", "제네바 (GVA)": "GVA", "인스부르크 (INN)": "INN",
    
    # 북유럽
    "코펜하겐 (CPH)": "CPH", "스톡홀름 (ARN)": "ARN", "오슬로 (OSL)": "OSL",
    "헬싱키 (HEL)": "HEL", "레이캬비크 (KEF)": "KEF",
    
    # 튀르키예 / 중동 거점
    "이스탄불 신공항 (IST)": "IST", "이스탄불 사비하 (SAW)": "SAW",
    "안탈리아 (AYT)": "AYT", "이즈미르 (ADB)": "ADB",
    
    # [직접 입력]
    "✏️ 공항코드 직접 입력": "CUSTOM"
}

def normalize_airline_name(name: str) -> str:
    """항공사명 표준화"""
    c = name.replace(" ", "").upper()
    if any(k in c for k in ["TURKISH", "터키", "THY"]): return "터키항공"
    if any(k in c for k in ["LUFTHANSA", "루프트한자"]): return "루프트한자"
    if any(k in c for k in ["KOREAN", "대한항공"]): return "대한항공"
    if any(k in c for k in ["ASIANA", "아시아나"]): return "아시아나항공"
    if any(k in c for k in ["EMIRATES", "에미레이트"]): return "에미레이트항공"
    if any(k in c for k in ["QATAR", "카타르"]): return "카타르항공"
    if any(k in c for k in ["AIRFRANCE", "에어프랑스"]): return "에어프랑스"
    if "KLM" in c: return "KLM네덜란드항공"
    if any(k in c for k in ["FINNAIR", "핀에어"]): return "핀에어"
    if any(k in c for k in ["ETIHAD", "에티하드"]): return "에티하드항공"
    return name.strip()

def search_single_leg(dep_code, arr_code, flight_date):
    """구글 플라이트 편도 독립 조회"""
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
        if not flights:
            return []

        parsed = []
        for f in flights[:15]:
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
    with st.expander("✈️ [Add-on] 유럽 전 지역 실시간 항공 스케줄 조회 & 자동 주입", expanded=True):
        c1, c2 = st.columns([1.5, 1])
        with c1:
            base_date = st.date_input("출발일 (IN)", datetime.now() + timedelta(days=60))
        with c2:
            tour_days = st.number_input("행사 기간 (일수)", min_value=3, max_value=25, value=9)

        c3, c4 = st.columns(2)
        with c3:
            in_sel = st.selectbox("IN 도시 선택", list(EUROPE_AIRPORTS.keys()), index=0)
            if EUROPE_AIRPORTS[in_sel] == "CUSTOM":
                in_code = st.text_input("IN 공항 3자리 코드 입력 (예: PRG, OPO)", "PRG").upper().strip()
                in_city_label = in_code
            else:
                in_code = EUROPE_AIRPORTS[in_sel]
                in_city_label = in_sel.split(" (")[0]

        with c4:
            out_sel = st.selectbox("OUT 도시 선택", list(EUROPE_AIRPORTS.keys()), index=1)
            if EUROPE_AIRPORTS[out_sel] == "CUSTOM":
                out_code = st.text_input("OUT 공항 3자리 코드 입력 (예: BUD, DBV)", "BUD").upper().strip()
                out_city_label = out_code
            else:
                out_code = EUROPE_AIRPORTS[out_sel]
                out_city_label = out_sel.split(" (")[0]

        ret_fly_date = base_date + timedelta(days=tour_days - 2)
        st.caption(f"🗓️ 실시간 매칭 요청: **1일차 출발 {base_date.strftime('%Y-%m-%d')} (인천➔{in_code}) | {tour_days-1}일차 귀국 {ret_fly_date.strftime('%Y-%m-%d')} ({out_code}➔인천)**")

        if st.button("🔎 동일 항공사 왕복 스케줄 실시간 조회", type="primary"):
            dep_code = "ICN"
            with st.spinner(f"출국(인천➔{in_code}) 및 귀국({out_code}➔인천) 동일 항공사 왕복을 매칭 중입니다..."):
                outbound_list = search_single_leg(dep_code, in_code, base_date)
                inbound_list = search_single_leg(out_code, dep_code, ret_fly_date)
                
                matched_pairs = []
                seen_airlines = set()
                
                for ob in outbound_list:
                    o_air = ob["norm_airline"]
                    same_rets = [ib for ib in inbound_list if ib["norm_airline"] == o_air]
                    if same_rets and o_air not in seen_airlines:
                        matched_pairs.append({
                            "airline": o_air,
                            "out": ob,
                            "ret": same_rets[0]
                        })
                        seen_airlines.add(o_air)
                
                st.session_state["flight_pairs"] = matched_pairs

        # 결과 카드 렌더링
        if "flight_pairs" in st.session_state:
            pairs = st.session_state["flight_pairs"]
            if not pairs:
                st.warning("⚠️ 출·귀국 항공사가 완전히 일치하는 스케줄을 찾지 못했습니다. 날짜를 조정하거나 공항 코드를 확인해 주세요.")
            else:
                st.write("---")
                st.markdown(f"#### 🎯 출귀국 동일 항공사 왕복 스케줄 ({len(pairs)}개 항공사 매칭)")
                
                for idx, pair in enumerate(pairs[:5]):
                    ob = pair["out"]
                    ib = pair["ret"]
                    airline_name = pair["airline"]
                    
                    ci, cb = st.columns([4, 1])
                    with ci:
                        st.info(
                            f"**옵션 {idx+1} [{airline_name}] ⭐ 왕복 동일 항공사**\n\n"
                            f"• **출국(1일차):** {ob['summary']}\n\n"
                            f"• **귀국({tour_days-1}일차):** {ib['summary']}"
                        )
                    with cb:
                        if st.button("👉 일정표 주입", key=f"inj_pair_{idx}"):
                            d1 = f"1. 인천 출발 / {ob['summary']} 도착 후 호텔 투숙\n"
                            mid = f"2. {in_city_label} 시내 핵심 관광지 투어 진행\n3. 전일 관광 및 이동\n4. 이동 및 전일 관광\n"
                            d_last_prev = f"{tour_days-1}. {out_city_label} 시내 관광 후 공항 이동 / {ib['summary']} / 기내 숙박\n"
                            d_last = f"{tour_days}. 인천 국제공항 ({ib['arr_time']}) 도착 후 해산"
                            
                            st.session_state["routing_text_injected"] = d1 + mid + d_last_prev + d_last
                            st.session_state["selected_input_mode"] = "[옵션 B] 직접 루팅 입력 (정밀 항공/일정 파싱)"
                            st.success(f"✅ [{airline_name}] 스케줄이 옵션 B에 자동 주입되었습니다!")
                            st.rerun()
