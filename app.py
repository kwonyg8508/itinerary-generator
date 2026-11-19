import streamlit as st
import openpyxl
from datetime import datetime
from pathlib import Path
import json
import os
import re
import glob
from google import genai
from google.genai import types
from itinerary_excel import render_itinerary

try:
    from flight_search import flight_rows, render_flight_search_ui
except ImportError:
    render_flight_search_ui = None

st.set_page_config(page_title="유럽 일정표 자동 생성 시스템 (Master Edition)", layout="wide", page_icon="🇪🇺")

if "generated_file" not in st.session_state: st.session_state.generated_file = None
if "generated_filename" not in st.session_state: st.session_state.generated_filename = None
if "selected_input_mode" not in st.session_state: st.session_state.selected_input_mode = "[옵션 B] 직접 루팅 입력 (정밀 항공/일정 파싱)"

st.title("🇪🇺 유럽 일정표 자동 생성 시스템 (Master Edition)")
st.caption("실무 완판 규격: 실시간 항공 검색 | 관광지 볼드 서식 완벽 사수 | 식사 명칭 단독 표기")

if render_flight_search_ui:
    render_flight_search_ui()

def get_available_db_files():
    db_folder = "database"
    if not os.path.exists(db_folder): return []
    return [os.path.basename(f) for f in glob.glob(os.path.join(db_folder, "*.xlsx"))]

def get_reference_db_context(user_query: str):
    db_folder = "database"
    excel_files = glob.glob(os.path.join(db_folder, "*.xlsx"))
    if not excel_files: return "", ""
    q_lower = user_query.lower()
    best_file = next((f for f in excel_files if any(kw in q_lower for kw in ["노르웨이", "발틱", "아이슬란드", "북유럽", "그리스", "로마", "이탈리아"])), excel_files[0])
    matched_filename = os.path.basename(best_file)
    try:
        wb = openpyxl.load_workbook(best_file, data_only=True)
        ws = wb.active
        content_summary = []
        for r in range(4, min(ws.max_row + 1, 40)):
            schedule = str(ws.cell(r, 5).value or "").strip()
            if schedule.startswith("▶"):
                content_summary.append(schedule[:50])
            if len(content_summary) >= 10: break
        return matched_filename, "\n".join(content_summary)
    except Exception:
        return "", ""

api_key = st.secrets.get("GEMINI_API_KEY")
client = genai.Client(api_key=api_key) if api_key else None

col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("1. 입력 방식 선택")
    input_mode_options = ["[옵션 A] 간편 단문 요청", "[옵션 B] 직접 루팅 입력 (정밀 항공/일정 파싱)"]
    selected_mode_idx = 1 if st.session_state.selected_input_mode == input_mode_options[1] else 0

    input_mode = st.radio("작성 방식을 선택하세요:", input_mode_options, index=selected_mode_idx, horizontal=True, key="radio_input_mode")
    st.session_state.selected_input_mode = input_mode

    start_date = st.date_input("출발일자 선택", datetime.today().date(), key="itinerary_start_date", disabled=bool(st.session_state.get("selected_flight_pair")))
    tour_title = st.text_input("행사명 (선택)", placeholder="예: 아이슬란드 오로라 일주 7일 SK")
    
    if "[옵션 A]" in input_mode:
        quick_prompt = st.text_input("간단 요청 / 웹 URL 입력", placeholder="예: 아이슬란드 오로라 일주 7일 SK")
    else:
        routing_text = st.text_area(
            "상세 루팅 입력 (조회한 항공편을 적용했으면 지상 일정만 입력하세요)",
            height=450,
            placeholder="""[1일차]
- 인천 국제공항 미팅 후 출국
- [SAS항공] SK506 | 인천 10:30 출발 ➔ 코펜하겐 경유
[2일차]
- 호텔 조식 후 로비 미팅
- 싱벨리오르로 이동 (약 1시간 소요)
- ▶ 유네스코 세계문화유산이자 세계 최신의 의회가 열린 **싱벨리어 국립공원** 관광
..."""
        )

with col2:
    st.subheader("2. 회사 양식 엑셀 출력")
    mascot = Path(__file__).with_name("assistant.gif")
    if mascot.is_file():
        from PIL import Image
        with Image.open(mascot) as image:
            st.image(image.convert("RGBA"), width=100)
    else:
        st.markdown("👓 김똘망")

    generate_btn = st.button("🚀 엑셀 일정표 생성하기", type="primary", use_container_width=True)

    if generate_btn:
        if not api_key: st.error("GEMINI_API_KEY가 등록되지 않았습니다.")
        else:
            raw_text_input = quick_prompt if "[옵션 A]" in input_mode else routing_text
            if not raw_text_input.strip():
                st.warning("⚠️ 입력 내용을 작성해 주세요!")
            else:
                st.session_state.generated_file = None
                combined_text = f"{tour_title} {raw_text_input}".strip()
                
                selected_pair = st.session_state.get("selected_flight_pair")
                ref_filename, ref_context = get_reference_db_context(combined_text)
                
                with st.spinner("김똘망 사원이 일정표를 신속하게 제작하고 있습니다..."):
                    try:
                        flight_context = ""
                        if selected_pair:
                            flight_events = flight_rows(selected_pair, start_date)
                            flight_context = "\n".join(
                                f"{day}일차: {row['schedule']} {row['time']}"
                                for day, events in sorted(flight_events.items())
                                for row in events["prepend"] + events["append"]
                            )
                            flight_context = (
                                "\n[확정 항공 일정: 지상 일정의 날짜 및 시간 배치에만 사용]\n"
                                + flight_context
                                + "\n항공 출발·도착·환승 행은 프로그램이 직접 넣습니다. "
                                "항공편 정보나 공항 도착/출발/환승 행을 JSON에 생성하지 마세요. "
                                "귀국일은 공항으로 이동하는 지상 일정까지만 작성하세요."
                            )
                        system_prompt = f"""
당신은 유럽 단체 패키지 수배 15년 차 수석 오퍼레이터 김똘망입니다.
아래 [사용자 입력 및 항공 루팅]을 절대 법전으로 삼아 고객 전달용 완벽한 엑셀 일정표 JSON을 생성하세요.

[사용자 입력 및 항공 루팅 (1순위 절대 반영)]
행사명: {tour_title}
요청내용: {raw_text_input}

{flight_context}

[참고용 가이드북 스니펫 (표현 참고용)]
{ref_context if ref_context else "(없음)"}

[🚨 랜드사 실무 절대 규칙 - 위반 절대 금지]
1. **항공 시간/편명 원본 사수**:
   - 사용자가 입력한 항공편명과 시간을 훼손 없이 정확히 반영하세요.
   - 항공편명(예: SK506)은 교통편(`transport`) 칸에 기재하고, 출발/도착 시간(예: 10:30)은 시간(`time`) 칸에 기재하세요.
2. **도보 이동수단 전면 삭제**:
   - 시내 관광 중 걷는 것은 교통편에 적지 않으며, 교통편 칸을 빈칸("")으로 비워두세요.
3. **'전용버스' 2행 고정 및 반복 금지**:
   - 2일차부터 차량 투어 시작 시 **2번째 행(index 1)에만** `transport="전용버스"`를 고정하고 반복 표기하지 마세요.
4. **관광 명소 부분 볼드체 서식 (매우 중요)**:
   - 관광지 설명 행은 반드시 **`▶ [설명] **[핵심명소]**`** 형식을 준수하여 핵심 명소 이름에 반드시 마크다운 볼드(`**`)를 적용하세요. (예: `▶ 유네스코 세계문화유산이자 세계 최신의 의회가 열린 **싱벨리어 국립공원** 관광`)
5. **호텔 명기 규칙**:
   - 사용자 입력에 기재된 호텔명과 등급을 그대로 사용하세요. 호텔 정보가 없으면 공란으로 두세요.
6. **식사 칸 데이터 정제 규칙 (매우 중요)**:
   - `meals` 배열에는 오직 **'호텔식', '기내식', '현지식', '자유식', '-'** 등 식사 종류 명칭만 순수하게 담으세요. 
   - 절대 "조식: 호텔식", "석식: 현지식"처럼 앞에 `조식:`, `석식:` 같은 접두어를 붙이지 마세요.
"""

                        json_schema = {
                            "type": "ARRAY",
                            "items": {
                                "type": "OBJECT",
                                "properties": {
                                    "day": {"type": "INTEGER"},
                                    "rows": {
                                        "type": "ARRAY",
                                        "items": {
                                            "type": "OBJECT",
                                            "properties": {
                                                "city": {"type": "STRING"},
                                                "transport": {"type": "STRING"},
                                                "time": {"type": "STRING"},
                                                "schedule": {"type": "STRING"}
                                            },
                                            "required": ["schedule"]
                                        }
                                    },
                                    "meals": {"type": "ARRAY", "items": {"type": "STRING"}},
                                    "hotel": {"type": "STRING"}
                                },
                                "required": ["day", "rows"]
                            }
                        }

                        response = client.models.generate_content(
                            model='gemini-3.8-flash',
                            contents=system_prompt,
                            config=types.GenerateContentConfig(response_mime_type="application/json", response_schema=json_schema)
                        )

                        raw_text = response.text.strip()
                        if "```" in raw_text:
                            raw_text = re.sub(r"^```(?:json)?\s*", "", raw_text, flags=re.MULTILINE)
                            raw_text = re.sub(r"\s*```$", "", raw_text, flags=re.MULTILINE)
                        
                        parsed_raw = json.loads(raw_text)
                        day_dict = {}
                        for item in parsed_raw:
                            d_num = item.get("day", 1)
                            if d_num not in day_dict:
                                day_dict[d_num] = {"day": d_num, "rows": [], "meals": ["-", "-", "-"], "hotel": item.get("hotel") or ""}
                            
                            raw_meals = item.get("meals") or ["-", "-", "-"]
                            cleaned_meals = []
                            for m in raw_meals:
                                m_str = str(m).strip()
                                m_str = re.sub(r"^(조식|중식|석식)\s*[:：\-]?\s*", "", m_str)
                                cleaned_meals.append(m_str if m_str else "-")
                            while len(cleaned_meals) < 3: cleaned_meals.append("-")
                            day_dict[d_num]["meals"] = cleaned_meals[:3]

                            for r_item in item.get("rows", []):
                                sc = str(r_item.get("schedule") or "").strip()
                                if sc in ["석식 후", "석식후"]: continue
                                r_item["schedule"] = sc.replace("[", "").replace("]", "")
                                day_dict[d_num]["rows"].append(r_item)
                            if item.get("hotel"): day_dict[d_num]["hotel"] = str(item.get("hotel")).replace("[", "").replace("]", "").strip()

                        if selected_pair:
                            return_day = (selected_pair["ret"][0]["dep"].date() - start_date).days + 1
                            outbound_arrival_day = (selected_pair["out"][-1]["arr"].date() - start_date).days + 1
                            last_day = max(flight_events)
                            if any(day < 1 or day > last_day for day in day_dict):
                                raise ValueError("지상 일정 날짜가 선택한 항공 일정 범위를 벗어났습니다.")
                            for day in flight_events:
                                day_dict.setdefault(day, {"day": day, "rows": [], "meals": ["-", "-", "-"], "hotel": ""})

                        parsed_days = [day_dict[k] for k in sorted(day_dict.keys())]
                        total_days_cnt = len(parsed_days)

                        for item in parsed_days:
                            d_num = item["day"]
                            rows = item["rows"]
                            
                            cleaned_rows = []
                            for r in rows:
                                sc = str(r.get("schedule") or "").strip()
                                tr = str(r.get("transport") or "").strip()
                                if tr in ["도보", "도보이동"]: continue
                                if any(k in sc for k in ["중식(", "석식(", "조식(", "중식 (", "석식 ("]): continue
                                cleaned_rows.append(r)
                            rows = cleaned_rows

                            seen_cities_in_day = set()
                            for r in rows:
                                c_val = str(r.get("city") or "").strip()
                                if c_val:
                                    if c_val in seen_cities_in_day:
                                        r["city"] = ""
                                    else:
                                        seen_cities_in_day.add(c_val)

                            if 1 < d_num < (last_day if selected_pair else total_days_cnt) and len(rows) > 1:
                                for idx, r in enumerate(rows):
                                    if idx > 1:
                                        if r.get("transport") == "전용버스": r["transport"] = ""
                                        if r.get("time") == "전  일": r["time"] = ""
                                        if str(r.get("schedule", "")).startswith("▶"): r["time"] = ""
                                if not rows[1].get("time") or rows[1].get("time") == "전  일":
                                    rows[1]["transport"] = "전용버스"
                                    rows[1]["time"] = ""

                            item["rows"] = rows

                            if selected_pair and d_num == last_day:
                                item["hotel"] = ""
                            elif selected_pair and (d_num < outbound_arrival_day or d_num >= return_day):
                                item["hotel"] = "HOTEL : 기내숙박"
                            elif not selected_pair and d_num == total_days_cnt - 1:
                                item["hotel"] = "HOTEL : 기내숙박"
                            elif not selected_pair and d_num == total_days_cnt:
                                item["hotel"] = ""

                        if selected_pair:
                            for item in parsed_days:
                                events = flight_events.get(item["day"], {"prepend": [], "append": []})
                                item["rows"] = events["prepend"] + item["rows"] + events["append"]

                        title = tour_title.strip() or f"{start_date:%Y-%m-%d} 유럽투어 일정표"
                        st.session_state.generated_file = render_itinerary(parsed_days, start_date, title)
                        st.session_state.generated_filename = f"{start_date.strftime('%Y-%m-%d')} {tour_title.strip() or '유럽투어'} - 일정표.xlsx"
                        st.rerun()

                    except Exception as e:
                        st.error(f"오류 상세: {str(e)}")

    if st.session_state.generated_file:
        st.success("✅ 파일 준비 완료!")
        st.download_button(
            label="📥 완성된 실무 마스터 엑셀 일정표 다운로드",
            data=st.session_state.generated_file,
            file_name=st.session_state.generated_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
