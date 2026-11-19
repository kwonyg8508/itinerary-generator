import re
from datetime import date, datetime, timedelta

import streamlit as st
from serpapi import GoogleSearch


AIRPORTS = {
    "PRG": "프라하", "BUD": "부다페스트", "VIE": "비엔나",
    "ZAG": "자그레브", "DBV": "두브로브니크", "SPU": "스플리트",
    "WAW": "바르샤바", "KRK": "크라쿠프", "LJU": "류블랴나",
    "CDG": "파리 샤를드골", "LHR": "런던 히드로", "AMS": "암스테르담",
    "MAD": "마드리드", "BCN": "바르셀로나", "LIS": "리스본",
    "FCO": "로마 피우미치노", "MXP": "밀라노 말펜사", "VCE": "베네치아",
    "ATH": "아테네", "FRA": "프랑크푸르트", "MUC": "뮌헨",
    "ZRH": "취리히", "CPH": "코펜하겐", "HEL": "헬싱키",
    "OSL": "오슬로", "ARN": "스톡홀름", "KEF": "레이캬비크",
    "IST": "이스탄불", "DXB": "두바이",
}


def parse_segment(segment):
    departure = segment["departure_airport"]
    arrival = segment["arrival_airport"]
    number = re.sub(r"\s+", "", str(segment["flight_number"])).upper()
    carrier = str(segment.get("airline_code") or "").strip().upper()
    if number.isdigit() and carrier:
        number = carrier + number
    if not re.fullmatch(r"[A-Z0-9]{2}\d+[A-Z]?", number):
        raise ValueError("편명 확인 필요")
    return {
        "number": number,
        "carrier": number[:2],
        "airline": segment.get("airline") or number[:2],
        "dep_airport": departure["id"],
        "arr_airport": arrival["id"],
        "dep": datetime.strptime(departure["time"], "%Y-%m-%d %H:%M"),
        "arr": datetime.strptime(arrival["time"], "%Y-%m-%d %H:%M"),
    }


def search_single_leg(dep_code, arr_code, flight_date):
    key = st.secrets.get("SERPAPI_API_KEY", "")
    if not key:
        raise ValueError("Streamlit Secrets에 SERPAPI_API_KEY를 설정해 주세요.")
    response = GoogleSearch({
        "engine": "google_flights", "departure_id": dep_code,
        "arrival_id": arr_code, "outbound_date": flight_date.isoformat(),
        "type": "2", "currency": "KRW", "hl": "ko", "gl": "kr", "api_key": key,
    }).get_dict()
    if response.get("error"):
        raise ValueError("항공 검색 API가 오류를 반환했습니다.")
    results = []
    for option in (response.get("best_flights") or []) + (response.get("other_flights") or []):
        raw = option.get("flights") or []
        if not 1 <= len(raw) <= 2:
            continue
        try:
            legs = [parse_segment(segment) for segment in raw]
        except (KeyError, TypeError, ValueError):
            continue
        if (legs[0]["dep_airport"], legs[-1]["arr_airport"], legs[0]["dep"].date()) != (
            dep_code, arr_code, flight_date
        ):
            continue
        if len(legs) == 2:
            first, second = legs
            if first["arr_airport"] != second["dep_airport"]:
                continue
            layover = second["dep"] - first["arr"]
            if not timedelta(hours=1) <= layover <= timedelta(hours=5):
                continue
        results.append(legs)
    return results


def flight_rows(pair, start_date):
    rows = {}
    for direction, legs in pair.items():
        for index, leg in enumerate(legs):
            events = (
                ("dep", "항공편 출발" if index == 0 else "환승편 출발"),
                ("arr", "경유 공항 도착" if index < len(legs) - 1 else "도착"),
            )
            for field, label in events:
                local_time = leg[field]
                day = (local_time.date() - start_date).days + 1
                if day < 1:
                    raise ValueError("항공 일정의 현지 날짜가 행사 시작일보다 앞섭니다.")
                airport = leg[field + "_airport"]
                row = {
                    "city": airport,
                    "transport": leg["number"],
                    "time": local_time.strftime("%H:%M"),
                    "schedule": f"{airport} {label}",
                }
                position = "prepend" if direction == "out" else "append"
                rows.setdefault(day, {"prepend": [], "append": []})[position].append(row)
    return rows


def render_flight_search_ui():
    with st.expander("✈️ 유럽 항공 스케줄 조회 및 일정표 적용"):
        first, second = st.columns(2)
        with first:
            outbound_date = st.date_input("인천 출발일", date.today() + timedelta(days=60), key="fs_base_date")
            codes = list(AIRPORTS) + ["직접 입력"]
            destination = st.selectbox("IN 공항", codes, index=22, format_func=lambda code: f"{AIRPORTS[code]} ({code})" if code in AIRPORTS else code)
            if destination == "직접 입력":
                destination = st.text_input("IN 공항 코드", key="fs_in_code").strip().upper()
        with second:
            return_date = st.date_input("현지 출발일 (귀국편)", outbound_date + timedelta(days=9), key="fs_return_date")
            origin = st.selectbox("OUT 공항", codes, format_func=lambda code: f"{AIRPORTS[code]} ({code})" if code in AIRPORTS else code)
            if origin == "직접 입력":
                origin = st.text_input("OUT 공항 코드", key="fs_out_code").strip().upper()

        criteria = outbound_date, return_date, destination, origin
        if st.session_state.get("flight_criteria") != criteria:
            st.session_state.pop("selected_flight_pair", None)
            st.session_state.pop("flight_pairs", None)
            st.session_state["flight_criteria"] = criteria
        if st.button("🔎 동일 항공사 출·귀국 조회", type="primary"):
            st.session_state.pop("selected_flight_pair", None)
            st.session_state.pop("flight_pairs", None)
            if return_date < outbound_date or not all(re.fullmatch(r"[A-Z]{3}", code) for code in (destination, origin)):
                st.warning("출발일과 귀국편 현지 출발일, 공항 코드 세 자리를 확인해 주세요.")
                return
            try:
                with st.spinner("항공 스케줄 조회 중"):
                    outbound = search_single_leg("ICN", destination, outbound_date)
                    inbound = search_single_leg(origin, "ICN", return_date)
            except Exception as error:
                st.error(str(error) if isinstance(error, ValueError) else "항공 조회 중 오류가 발생했습니다.")
                return
            first_return = {}
            for legs in inbound:
                first_return.setdefault(legs[0]["carrier"], legs)
            pairs = {}
            for legs in outbound:
                carrier = legs[0]["carrier"]
                if carrier in first_return:
                    pairs.setdefault(carrier, {"out": legs, "ret": first_return[carrier]})
            st.session_state["flight_pairs"] = list(pairs.values())

        pairs = st.session_state.get("flight_pairs")
        if pairs == []:
            st.warning("조회 결과에서 조건에 맞는 동일 항공사 조합을 찾지 못했습니다.")
        for index, pair in enumerate((pairs or [])[:5]):
            st.markdown(f"**{pair['out'][0]['airline']}**")
            for label, direction in (("출국", "out"), ("귀국", "ret")):
                st.write(label)
                for leg in pair[direction]:
                    st.text(f"{leg['number']}  {leg['dep_airport']} {leg['dep']:%Y-%m-%d %H:%M} → {leg['arr_airport']} {leg['arr']:%Y-%m-%d %H:%M}")
            if st.button("이 항공편을 일정표에 적용", key=f"apply_flights_{index}"):
                st.session_state["selected_flight_pair"] = pair
                st.session_state["itinerary_start_date"] = outbound_date
                st.session_state["generated_file"] = None
        if st.session_state.get("selected_flight_pair"):
            selected = st.session_state["selected_flight_pair"]
            st.success(f"{selected['out'][0]['number']} 출국 / {selected['ret'][0]['number']} 귀국 적용")
            if st.button("항공편 적용 해제"):
                st.session_state.pop("selected_flight_pair", None)
                st.session_state["generated_file"] = None
                st.rerun()
