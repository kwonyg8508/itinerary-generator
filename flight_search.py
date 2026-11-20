import re
from datetime import date, datetime, timedelta

import streamlit as st


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


def search_single_leg(dep_code, arr_code, flight_date, diagnostics=None, *, direct_only=False):
    key = st.secrets.get("SERPAPI_API_KEY", "")
    if not key:
        raise ValueError("Streamlit Secrets에 SERPAPI_API_KEY를 설정해 주세요.")
    try:
        from serpapi import GoogleSearch
    except ImportError as error:
        raise ValueError(
            "항공 검색 라이브러리를 불러오지 못했습니다. "
            "requirements.txt에 google-search-results가 설치되도록 설정해 주세요."
        ) from error

    response = GoogleSearch({
        "engine": "google_flights", "departure_id": dep_code,
        "arrival_id": arr_code, "outbound_date": flight_date.isoformat(),
        "type": "2", "currency": "KRW", "hl": "ko", "gl": "kr", "api_key": key,
    }).get_dict()
    if response.get("error"):
        raise ValueError("항공 검색 API가 오류를 반환했습니다.")
    results = []
    short_connections = []
    for option in (response.get("best_flights") or []) + (response.get("other_flights") or []):
        raw = option.get("flights") or []
        numbers = " / ".join(str(leg.get("flight_number") or "편명 없음") for leg in raw)
        reason = ""
        legs = []
        if direct_only and len(raw) != 1:
            continue
        if not 1 <= len(raw) <= 2:
            reason = "구간 정보 없음 또는 2회 이상 경유"
        else:
            try:
                legs = [parse_segment(segment) for segment in raw]
            except (KeyError, TypeError, ValueError):
                reason = "편명·공항·시간 데이터 형식 불일치"
        if legs:
            if (legs[0]["dep_airport"], legs[-1]["arr_airport"], legs[0]["dep"].date()) != (
                dep_code, arr_code, flight_date
            ):
                reason = "요청한 공항 또는 출발일과 다름"
            elif len(legs) == 2:
                first, second = legs
                layover = second["dep"] - first["arr"]
                if first["arr_airport"] != second["dep_airport"]:
                    reason = "환승 공항이 다름"
                elif not timedelta(hours=1) <= layover <= timedelta(hours=5):
                    reason = f"환승 {layover // timedelta(minutes=1)}분 (허용 60~300분)"
                    if timedelta(0) <= layover < timedelta(hours=1):
                        short_connections.append((first, second["carrier"]))
        if diagnostics is not None:
            diagnostics.append({"API 편명": numbers or "—", "처리": reason or "표시"})
        if not reason:
            results.append(legs)
    if short_connections:
        results.extend(_find_later_connections(short_connections, arr_code, results, diagnostics))
    return results


def _find_later_connections(short_connections, destination, existing, diagnostics):
    def identity(legs):
        return tuple((leg["number"], leg["dep"], leg["arr"]) for leg in legs)

    seen = {identity(legs) for legs in existing}
    searches = {}
    added = []
    for first, onward_carrier in short_connections:
        earliest = first["arr"] + timedelta(hours=1)
        latest = first["arr"] + timedelta(hours=5)
        for when in dict.fromkeys((earliest.date(), latest.date())):
            key = first["arr_airport"], destination, when
            if key not in searches:
                if len(searches) >= 3:
                    continue
                try:
                    searches[key] = search_single_leg(*key, direct_only=True)
                except Exception:
                    searches[key] = []
                    if diagnostics is not None:
                        diagnostics.append({"API 편명": "—", "처리": f"보완 조회 실패 ({key[0]}→{destination}, {when})"})
            for option in searches[key]:
                second = option[0]
                if second["carrier"] != onward_carrier or not earliest <= second["dep"] <= latest:
                    continue
                legs = [first, {**second, "supplemental_connection": True}]
                if identity(legs) in seen:
                    continue
                seen.add(identity(legs))
                added.append(legs)
                if diagnostics is not None:
                    diagnostics.append({"API 편명": f"{first['number']} / {second['number']}", "처리": "구간별 보완 조회로 추가"})
    return added


EXTRA_AIRPORT_CITIES = {
    "ICN": "인천", "GMP": "서울", "PUS": "부산", "DOH": "도하",
    "AUH": "아부다비", "CTA": "카타니아", "PMO": "팔레르모", "MLA": "몰타",
    "TLL": "탈린", "RIX": "리가", "VNO": "빌뉴스", "BER": "베를린",
    "DUS": "뒤셀도르프", "BRU": "브뤼셀", "GVA": "제네바", "LYS": "리옹",
    "NCE": "니스", "MRS": "마르세유", "JTR": "산토리니", "HER": "이라클리온",
    "NAP": "나폴리", "BLQ": "볼로냐", "FLR": "피렌체", "TRN": "토리노",
    "AGP": "말라가", "SVQ": "세비야", "BIO": "빌바오", "OPO": "포르투",
    "BGO": "베르겐", "GOT": "예테보리", "RVN": "로바니에미", "KTT": "키틸래",
    "LUX": "룩셈부르크", "SOF": "소피아", "OTP": "부쿠레슈티", "SAJ": "사라예보",
    "SJJ": "사라예보", "BEG": "베오그라드", "SKP": "스코페", "TIA": "티라나",
    "DUB": "더블린", "EDI": "에든버러", "MAN": "맨체스터", "LHR": "런던",
    "FCO": "로마", "MXP": "밀라노", "CDG": "파리",
}


def _flight_city(code):
    return EXTRA_AIRPORT_CITIES.get(code) or AIRPORTS.get(code) or "공항"


def flight_rows(pair, start_date):
    rows = {}
    for direction, legs in pair.items():
        for leg in legs:
            dep_city = _flight_city(leg["dep_airport"])
            arr_city = _flight_city(leg["arr_airport"])
            events = (
                ("dep", dep_city, f"{dep_city} 출발, {arr_city} 향발"),
                ("arr", arr_city, f"{arr_city} 도착"),
            )
            for field, city, schedule in events:
                local_time = leg[field]
                day = (local_time.date() - start_date).days + 1
                if day < 1:
                    raise ValueError("항공 일정의 현지 날짜가 행사 시작일보다 앞섭니다.")
                row = {
                    "city": city,
                    "transport": leg["number"] if field == "dep" else "",
                    "time": local_time.strftime("%H:%M"),
                    "schedule": schedule,
                    "flight_event": field,
                }
                position = "prepend" if direction == "out" else "append"
                rows.setdefault(day, {"prepend": [], "append": []})[position].append(row)
    return rows


def _flight_option_label(legs):
    prefix = "[구간별 조회 조합] " if any(leg.get("supplemental_connection") for leg in legs) else ""
    return prefix + " / ".join(
        f"{leg['number']} {leg['dep_airport']} {leg['dep']:%m-%d %H:%M} → "
        f"{leg['arr_airport']} {leg['arr']:%m-%d %H:%M}"
        for leg in legs
    )


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
            st.session_state.pop("flight_direction_results", None)
            st.session_state["flight_criteria"] = criteria
        if st.button("🔎 출국편·귀국편 조회", type="primary"):
            st.session_state.pop("selected_flight_pair", None)
            st.session_state.pop("flight_direction_results", None)
            st.session_state["generated_file"] = None
            if return_date < outbound_date or not all(re.fullmatch(r"[A-Z]{3}", code) for code in (destination, origin)):
                st.warning("출발일과 귀국편 현지 출발일, 공항 코드 세 자리를 확인해 주세요.")
                return
            results = {}
            with st.spinner("출국편과 귀국편을 각각 조회하고 있습니다."):
                for direction, dep, arr, when in (
                    ("out", "ICN", destination, outbound_date),
                    ("ret", origin, "ICN", return_date),
                ):
                    details = []
                    try:
                        legs = search_single_leg(dep, arr, when, diagnostics=details)
                        results[direction] = legs, details, ""
                    except Exception as error:
                        message = str(error) if isinstance(error, ValueError) else "항공 API 요청 실패"
                        results[direction] = [], details, message
            st.session_state["flight_direction_results"] = results

        results = st.session_state.get("flight_direction_results")
        if results:
            show_details = st.checkbox("API 반환 편명과 제외 이유 보기", key="fs_show_details")
            chosen = {}
            for direction, label in (("out", "출국편"), ("ret", "귀국편")):
                legs, details, error = results[direction]
                st.markdown(f"**{label} · 표시 {len(legs)}개**")
                if error:
                    st.warning(f"{label} 조회 실패 · {error}")
                elif not legs:
                    st.info("API가 후보를 반환하지 않았습니다." if not details else "API 후보가 모두 제외됐습니다.")
                if show_details and details:
                    st.dataframe(details, hide_index=True)
                if legs:
                    index = st.selectbox(
                        f"{label} 선택", range(len(legs)), key=f"fs_choose_{direction}",
                        format_func=lambda i, options=legs: _flight_option_label(options[i]),
                    )
                    chosen[direction] = legs[index]
            if any(leg.get("supplemental_connection") for legs in chosen.values() for leg in legs):
                st.caption("구간별 조회 조합은 견적용입니다. 단일 예약 및 단체석은 수배 단계에서 확인하세요.")
            if len(chosen) == 2 and st.button("선택한 출·귀국편을 일정표에 적용"):
                st.session_state["selected_flight_pair"] = chosen
                st.session_state["itinerary_start_date"] = outbound_date
                st.session_state["generated_file"] = None
        if st.session_state.get("selected_flight_pair"):
            selected = st.session_state["selected_flight_pair"]
            st.success(f"{selected['out'][0]['number']} 출국 / {selected['ret'][0]['number']} 귀국 적용")
            if st.button("항공편 적용 해제"):
                st.session_state.pop("selected_flight_pair", None)
                st.session_state["generated_file"] = None
                st.rerun()


REGION_ROUTES = (
    (("동유럽", "체코", "헝가리"), (("VIE", "PRG"), ("PRG", "BUD"), ("BUD", "VIE"))),
    (("이탈리아", "돌로미티"), (("MXP", "FCO"), ("FCO", "MXP"))),
    (("스페인", "포르투갈"), (("MAD", "LIS"), ("MAD", "BCN"), ("BCN", "MAD"))),
    (("그리스",), (("ATH", "ATH"),)),
    (("아이슬란드",), (("KEF", "KEF"),)),
    (("북유럽",), (("CPH", "HEL"), ("OSL", "CPH"), ("ARN", "HEL"))),
)


def inquiry_conditions(prompt, today):
    month_match = re.search(r"(?:(20\d\d)\s*년\s*|(내년|올해)\s*)?(1[0-2]|0?[1-9])\s*월", prompt)
    days_match = re.search(r"(\d{1,2})\s*일(?!차|자)", prompt)
    routes = next((routes for keywords, routes in REGION_ROUTES if any(word in prompt for word in keywords)), None)
    if not month_match or not days_match or not routes:
        raise ValueError("출발 월, 총 일정 일수, 지원 지역을 입력해 주세요. 예: 내년 3월 동유럽 10일 16명 국적기")
    explicit_year, relative_year, month = month_match.groups()
    month = int(month)
    year = int(explicit_year) if explicit_year else today.year + (relative_year == "내년")
    if not explicit_year and not relative_year and month < today.month:
        year += 1
    days = int(days_match[1])
    if not 3 <= days <= 25:
        raise ValueError("총 일정 일수는 3~25일로 입력해 주세요.")
    return year, month, days, routes, "국적기" in prompt


def search_inquiry(prompt, today, search=search_single_leg, plan=None):
    if plan is None:
        year, month, days, routes, national = inquiry_conditions(prompt, today)
        exact_date = None
    else:
        year, month, days = (plan[key] for key in ("year", "month", "days"))
        national = "국적기" in prompt or plan.get("national", False)
        routes = list(dict.fromkeys((r["in"].strip().upper(), r["out"].strip().upper()) for r in plan["routes"]))[:3]
        if not 3 <= days <= 25 or not routes or not all(
            re.fullmatch(r"[A-Z]{3}", code) and code != "ICN" for route in routes for code in route
        ):
            raise ValueError("AI가 제안한 공항이나 여행 일수를 확인할 수 없습니다. 문의 조건을 보완해 주세요.")
        exact_date = plan.get("departure_day")
    dates = [date(year, month, exact_date)] if exact_date else [date(year, month, day) for day in (3, 10, 17, 24)]
    dates = [day for day in dates if day >= today]
    if not dates and not exact_date and (year, month) == (today.year, today.month):
        dates = [today]
    if not dates:
        raise ValueError("이미 지난 달입니다. 출발 월을 확인해 주세요.")
    cache = {}
    candidates = []

    def find(origin, destination, when):
        key = origin, destination, when
        if key not in cache:
            cache[key] = search(*key)
        return cache[key]

    for departure in dates:
        arrival_home = departure + timedelta(days=days - 1)
        for destination, origin in routes:
            for outbound in find("ICN", destination, departure):
                carrier = outbound[0]["carrier"]
                if national and carrier not in {"KE", "OZ"}:
                    continue
                for return_date in (arrival_home - timedelta(days=1), arrival_home):
                    if return_date < outbound[-1]["arr"].date():
                        continue
                    for inbound in find(origin, "ICN", return_date):
                        if (not national or inbound[0]["carrier"] in {"KE", "OZ"}) and inbound[-1]["arr"].date() == arrival_home:
                            candidates.append({"out": outbound, "ret": inbound})
    candidates.sort(key=lambda pair: (len(pair["out"]) + len(pair["ret"]), pair["out"][0]["dep"]))
    distinct = {}
    for pair in candidates:
        outbound, inbound = pair["out"], pair["ret"]
        key = outbound[0]["dep"].date(), outbound[-1]["arr_airport"], inbound[0]["dep_airport"]
        distinct.setdefault(key, pair)
    return list(distinct.values())[:3]


def render_inquiry_flights(prompt, planner=None):
    if st.button("간단 요청에서 항공 후보 찾기", key="inquiry_search"):
        st.session_state.pop("inquiry_flight_results", None)
        try:
            with st.spinner("문의의 출발 조건과 공항 조합을 검토하고 항공편을 조회합니다."):
                today = date.today()
                plan = planner(prompt, today) if planner else None
                results = search_inquiry(prompt, today, plan=plan)
                st.session_state["inquiry_airport_plan"] = plan
            st.session_state["inquiry_flight_results"] = (prompt, results)
        except ValueError as error:
            st.warning(str(error))
        except Exception:
            st.error("항공 조회에 실패했습니다. 잠시 후 다시 시도해 주세요.")
    query, results = st.session_state.get("inquiry_flight_results", ("", None))
    if query != prompt or results is None:
        return
    plan = st.session_state.get("inquiry_airport_plan")
    if plan:
        st.caption("조회한 공항 조합 · " + ", ".join(f"{r['in']} IN / {r['out']} OUT" for r in plan["routes"][:3]))
    if not results:
        st.info("대표 날짜에서 일치하는 조합이 없어요. 화면 상단에서 날짜와 공항을 직접 조회할 수 있습니다.")
    for index, pair in enumerate(results):
        for label, direction in (("출국", "out"), ("귀국", "ret")):
            st.write(label)
            if any(leg.get("supplemental_connection") for leg in pair[direction]):
                st.caption("구간별 조회 조합 · 견적용")
            for leg in pair[direction]:
                st.text(f"{leg['number']}  {leg['dep_airport']} {leg['dep']:%m-%d %H:%M} → {leg['arr_airport']} {leg['arr']:%m-%d %H:%M}")
        def apply(selected=pair):
            st.session_state["selected_flight_pair"] = selected
            st.session_state["itinerary_start_date"] = selected["out"][0]["dep"].date()
            st.session_state["generated_file"] = None
        st.button("이 후보를 일정표에 적용", key=f"inquiry_apply_{index}", on_click=apply)
