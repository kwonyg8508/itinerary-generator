import streamlit as st
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta
import io
import json
import os
import re
import glob
import base64
import time
from google import genai
from google.genai import types

# 항공 검색 애드온 모듈 안전 임포트
try:
    from flight_search import render_flight_search_ui
except ImportError:
    render_flight_search_ui = None

# 1. 페이지 설정
st.set_page_config(page_title="유럽 랜드사 일정표 생성기 (Master Edition)", layout="wide", page_icon="🇪🇺")

# 2. 세션 상태(Session State) 초기화
if "generated_file" not in st.session_state:
    st.session_state.generated_file = None
if "generated_filename" not in st.session_state:
    st.session_state.generated_filename = None
if "ui_status" not in st.session_state:
    st.session_state.ui_status = "idle"
if "ui_msg" not in st.session_state:
    st.session_state.ui_msg = "선배님! 7대 실무 수배 룰(시간 보정·08:55 해제·동일항공사 매칭) 세팅 완료되었습니다. 요청해 주세요! ☕"
if "routing_text_injected" not in st.session_state:
    st.session_state.routing_text_injected = ""
if "selected_input_mode" not in st.session_state:
    st.session_state.selected_input_mode = "[옵션 A] 간편 단문 요청 (사내 DB 자동 매칭)"

# 3. 커스텀 CSS
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Pretendard:wght@400;600;700&display=swap');

html, body, [class*="css"] {
    font-family: 'Pretendard', sans-serif;
}

.assistant-card {
    background: #ffffff;
    border: 2px solid #e2e8f0;
    border-radius: 16px;
    padding: 16px;
    box-shadow: 0 4px 15px rgba(0, 0, 0, 0.04);
    display: flex;
    align-items: center;
    gap: 16px;
    margin-bottom: 20px;
}

.gif-wrapper {
    width: 135px;
    height: 125px;
    flex-shrink: 0;
    border-radius: 12px;
    overflow: hidden;
    border: 2px solid #cbd5e1;
    background: #0f172a;
    display: flex;
    align-items: center;
    justify-content: center;
}

.gif-wrapper img {
    width: 100%;
    height: 100%;
    object-fit: cover;
    image-rendering: pixelated;
}

.speech-bubble {
    flex-grow: 1;
    background: #f8fafc;
    border: 1.5px solid #e2e8f0;
    border-radius: 12px;
    padding: 14px 16px;
    position: relative;
}

.speech-bubble::before {
    content: '';
    position: absolute;
    left: -10px;
    top: 50%;
    transform: translateY(-50%);
    border-width: 6px 10px 6px 0;
    border-style: solid;
    border-color: transparent #e2e8f0 transparent transparent;
}

.assistant-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-bottom: 6px;
}

.assistant-name {
    font-size: 14px;
    font-weight: 700;
    color: #1e293b;
    display: flex;
    align-items: center;
    gap: 6px;
}

.assistant-badge {
    font-size: 11px;
    font-weight: 600;
    padding: 3px 8px;
    border-radius: 20px;
}

.badge-idle { background: #e2e8f0; color: #475569; }
.badge-working { background: #fee2e2; color: #ef4444; }
.badge-success { background: #dcfce7; color: #15803d; }

.assistant-message {
    font-size: 13.5px;
    font-weight: 500;
    color: #334155;
    line-height: 1.5;
}

.db-info-box {
    background: #eff6ff;
    border: 1px solid #bfdbfe;
    border-radius: 8px;
    padding: 10px 14px;
    font-size: 13px;
    color: #1e40af;
    margin-bottom: 12px;
}
</style>
""", unsafe_allow_html=True)

st.title("🇪🇺 유럽 일정표 자동 생성 시스템 (Master Edition)")
st.caption("실무 완판 규격: 7대 실무 수배 룰 적용 완료 | 사내 DB 20줄 초경량화 | 명소 부분 볼드")

# [Add-on] 실시간 항공 스케줄 조회 & 자동 주입 엔진 렌더링
if render_flight_search_ui:
    render_flight_search_ui()

# 4. database 폴더 내 사내 엑셀 DB 목록 스캔
def get_available_db_files():
    db_folder = "database"
    if not os.path.exists(db_folder):
        return []
    return [os.path.basename(f) for f in glob.glob(os.path.join(db_folder, "*.xlsx"))]

def get_target_db_context(user_query: str):
    db_folder = "database"
    excel_files = glob.glob(os.path.join(db_folder, "*.xlsx"))
    if not excel_files:
        return None, ""

    q_lower = user_query.lower()
    best_file = None
    best_score = -1

    for filepath in excel_files:
        filename = os.path.basename(filepath)
        f_lower = filename.lower()
        score = 0

        keywords = ["동유럽", "서유럽", "터키", "튀르키예", "스페인", "포르투갈", "그리스", "이탈리아", "스위스", "발칸"]
        for kw in keywords:
            if kw in q_lower and kw in f_lower:
                score += 10

        if ("ke" in q_lower or "대한항공" in q_lower) and ("ke" in f_lower or "대한항공" in f_lower):
            score += 5
        if ("tk" in q_lower or "터키항공" in q_lower) and ("tk" in f_lower or "터키항공" in f_lower):
            score += 5
        if ("oz" in q_lower or "아시아나" in q_lower) and ("oz" in f_lower or "아시아나" in f_lower):
            score += 3

        days_match = re.search(r"(\d+)일", user_query)
        if days_match and (f"{days_match.group(1)}일" in filename or f"{days_match.group(1)}d" in f_lower):
            score += 5

        if score > best_score:
            best_score = score
            best_file = filepath

    if not best_file or best_score <= 0:
        best_file = excel_files[0]

    matched_filename = os.path.basename(best_file)

    try:
        wb = openpyxl.load_workbook(best_file, data_only=True)
        ws = wb.active
        content_summary = []
        
        for r in range(4, min(ws.max_row + 1, 90)):
            city = ws.cell(r, 2).value
            schedule = str(ws.cell(r, 5).value or "").strip()
            
            if schedule.startswith("▶") or (city and len(str(city).strip()) > 1):
                clean_sc = schedule.replace("\n", " ")[:60]
                line = f"도시: {str(city).strip()} | {clean_sc}" if city else f"{clean_sc}"
                content_summary.append(line)
                
            if len(content_summary) >= 20:
                break

        extracted = f"=== [사내 표준 핵심 동선(20줄 압축): {matched_filename}] ===\n" + "\n".join(content_summary)
        return matched_filename, extracted
    except Exception:
        return matched_filename, ""

db_file_list = get_available_db_files()

api_key = st.secrets.get("GEMINI_API_KEY")
client = genai.Client(api_key=api_key) if api_key else None

col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("1. 입력 방식 선택")
    
    # 상단 검색기에서 일정표 주입 시 [옵션 B] 라디오 버튼으로 자동 연동
    input_mode_options = ["[옵션 A] 간편 단문 요청 (사내 DB 자동 매칭)", "[옵션 B] 직접 루팅 입력 (정밀 항공/일정 파싱)"]
    selected_mode_idx = 1 if st.session_state.selected_input_mode == input_mode_options[1] else 0

    input_mode = st.radio(
        "작성 방식을 선택하세요:",
        input_mode_options,
        index=selected_mode_idx,
        horizontal=True,
        key="radio_input_mode"
    )
    st.session_state.selected_input_mode = input_mode
    
    if "[옵션 A]" in input_mode:
        if db_file_list:
            st.markdown(f"""
            <div class="db-info-box">
                📚 <b>연동 대기 사내 엑셀 DB ({len(db_file_list)}개):</b><br>
                {' • '.join([f'`{f}`' for f in db_file_list])}<br>
                <span style="font-size:11.5px; color:#2563eb;">* 요청에 맞춰 단 1개 엑셀의 핵심 20행만 초경량 전송되어 503 에러를 차단합니다.</span>
            </div>
            """, unsafe_allow_html=True)
        else:
            st.info("💡 `database/` 폴더에 사내 엑셀이 있을 경우 자동 매핑됩니다.")

    start_date = st.date_input("출발일자 선택", datetime.today())
    tour_title = st.text_input("행사명 (선택)", placeholder="예: 2026-12-05 동유럽 9일 TK")
    
    if "[옵션 A]" in input_mode:
        quick_prompt = st.text_input(
            "간단 요청 입력",
            placeholder="예: 동유럽 9일 TK 프라하인 부다페스트아웃 / 이탈리아 일주 9일"
        )
    else:
        current_injected_val = st.session_state.get("routing_text_injected", "")
        routing_text = st.text_area(
            "루팅 및 일정 입력 (복사-붙여넣기)",
            value=current_injected_val,
            height=280,
            placeholder="""1. 인천 출발 / [터키항공] TK021 ICN 10:45 ➔ IST 16:45 / 프라하 도착 후 호텔 투숙
2. 프라하 (카를교, 프라하 성) 관광 후 체스키크룸로프로 이동 (약 2시간 40분 소요)
..."""
        )

with col2:
    st.subheader("2. AI 수배팀 작업 현황")
    ui_placeholder = st.empty()

    def render_assistant_ui(status="idle", message="선배님, 요청해 주시면 즉각 빌드하겠습니다! ☕"):
        if status == "success":
            badge_text = "🎉 제작 완료!"
            badge_class = "badge-success"
            target_candidates = ["done.gif", "done.png", "assistant.gif", "working.gif"]
        elif status == "working":
            badge_text = "🔥 폭풍 작업 중"
            badge_class = "badge-working"
            target_candidates = ["working.gif", "assistant.gif", "done.gif"]
        else:
            badge_text = "☕ 대기 중"
            badge_class = "badge-idle"
            target_candidates = ["working.gif", "assistant.gif", "done.gif"]

        media_file = None
        for cand in target_candidates:
            if os.path.exists(cand):
                media_file = cand
                break

        media_b64 = ""
        if media_file:
            with open(media_file, "rb") as f:
                mime = "image/gif" if media_file.endswith(".gif") else "image/png"
                media_b64 = f"data:{mime};base64,{base64.b64encode(f.read()).decode()}"

        img_tag = f'<img src="{media_b64}" alt="김똘망 사원">' if media_b64 else '<span style="color:white;font-size:24px;">👓</span>'

        html_content = f"""
        <div class="assistant-card">
            <div class="gif-wrapper">
                {img_tag}
            </div>
            <div class="speech-bubble">
                <div class="assistant-header">
                    <span class="assistant-name">👓 김똘망 사원 <span style="font-size:12px; font-weight:normal; color:#64748b;">(유럽 OP팀 막내)</span></span>
                    <span class="assistant-badge {badge_class}">{badge_text}</span>
                </div>
                <div class="assistant-message">{message}</div>
            </div>
        </div>
        """
        ui_placeholder.markdown(html_content, unsafe_allow_html=True)

    render_assistant_ui(st.session_state.ui_status, st.session_state.ui_msg)

    generate_btn = st.button("🚀 엑셀 일정표 생성하기", type="primary", use_container_width=True)

    if generate_btn:
        if not api_key:
            st.error("설정(Secrets)에서 GEMINI_API_KEY가 등록되지 않았습니다.")
        elif "[옵션 A]" in input_mode and not quick_prompt.strip():
            st.warning("간단 요청 내용을 입력해 주세요.")
        elif "[옵션 B]" in input_mode and not routing_text.strip():
            st.warning("루팅 정보를 입력해 주세요.")
        else:
            st.session_state.generated_file = None
            
            user_input_text = quick_prompt if "[옵션 A]" in input_mode else routing_text
            matched_file_name, target_db_context = get_target_db_context(user_input_text)
            
            msg = f"'{matched_file_name}' 20줄 핵심 동선과 7대 실무 수배 룰을 결합하여 완판본을 제작합니다! ⚡" if matched_file_name else "실무 수배 규격으로 조립 중입니다! ⚡"
            render_assistant_ui("working", msg)
            
            with st.spinner("김똘망 사원이 일정표를 신속하게 제작하고 있습니다..."):
                try:
                    system_prompt = f"""
당신은 15년 차 유럽 전문 랜드사의 수석 오퍼레이터 김똘망입니다.
아래 제공된 [선택된 사내 기준 핵심 동선] 및 사용자의 입력을 바탕으로 고객 전달용 완벽한 엑셀 일정표 JSON을 생성하세요.

[선택된 사내 기준 핵심 동선]
{target_db_context if target_db_context else "(표준 랜드사 규격 준수)"}

[사용자 입력 요청]
{user_input_text}

[🚨 7대 실무 절대 규칙 - 위반 절대 금지]
1. **도시 이동 표기 규칙 (가장 중요)**:
   - 미사여구(수려한, 고풍스러운 등) 절대 사용 금지!
   - 반드시 **'[도착도시명]으로 이동 (약 O시간 O분 소요)'** 형태로만 작성할 것.
   - 이동 시간은 구글 일반 승용차 기준 시간에서 대형버스 운행을 감안하여 **무조건 +30분 합산**된 시간을 기재할 것.
   - 예: 밀라노 ➔ 베네치아 (구글 3시간) -> `베네치아로 이동 (약 3시간 30분 소요)`

2. **'석식 후' 단독 행 전면 삭제**:
   - 매일 일정 끝에 반복되던 `석식 후` 단독 행은 **절대 작성하지 마세요**.
   - 관광 투어가 끝나면 곧바로 `호텔 투숙 및 휴식` 또는 `호텔로 이동 후 휴식`으로 마감하세요.

3. **이탈리아 투어 필수 예외 룰**:
   - **우피치 미술관 제외**: 피렌체 투어 시 사전 예약 및 대기시간이 극심한 '우피치 미술관'은 기본 일정에서 **반드시 제외**할 것. (시뇨리아 광장, 두오모 성당 등 시내 중심 도보 투어로 대체)
   - **친퀘테레 1개 마을 제한**: 친퀘테레 투어 시 여러 마을을 돌지 말고 **반드시 대표 1개 마을만 방문** 후 다음 목적지로 이동할 것.

4. **교통편 및 시간 규격화**:
   - 2일차부터 마지막 전날까지 도시간 이동 및 투어가 있는 날은 `transport="전용차량"`, `time="전  일"`을 첫 번째 이동 또는 투어 행에 명확히 표기할 것.
   - 세모표(▶) 관광지 행에는 시간(time)을 빈칸("")으로 비워둘 것.

5. **항공 및 귀국 시간 규칙 (절대 준수)**:
   - 1일차 출국 및 마지막 전날 귀국 비행기 편명과 시간을 요청 내용 그대로 정확히 반영할 것.
   - **인천 도착 시간**: 마지막 날 인천 도착 시간은 08:55로 고정하지 말고, **사용자가 입력한 항공편의 실제 도착시간**을 그대로 반영할 것. (도착시간이 지정되지 않은 경우에만 통상적인 오전 도착시간 반영)
   - 귀국 출발 행 다음 행은 `schedule="기내 숙박"`, `hotel="HOTEL : 기내 숙박"`.

6. **관광 명소 서식 (부분 볼드 처리)**:
   - 설명은 보통체로, 맨 뒤 핵심 명소만 **기호로 감쌀 것: ` ▶ [설명] **[핵심명소]**`

7. **식사 및 호텔**:
   - 식사는 "호텔식", "현지식", "한   식", "중국식", "기내식", "-" 중에서만 기재.
   - 호텔: 튀르키예는 "HOTEL : ANY 5* HOTEL CLASS", 그 외 유럽은 "HOTEL : ANY 4* HOTEL CLASS".
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
                                "meals": {
                                    "type": "ARRAY",
                                    "items": {"type": "STRING"}
                                },
                                "hotel": {"type": "STRING"}
                            },
                            "required": ["day", "rows"]
                        }
                    }

                    target_model = 'gemini-3.8-flash'
                    response = None
                    last_error = None

                    for attempt in range(3):
                        try:
                            response = client.models.generate_content(
                                model=target_model,
                                contents=system_prompt,
                                config=types.GenerateContentConfig(
                                    response_mime_type="application/json",
                                    response_schema=json_schema
                                )
                            )
                            if response and response.text:
                                break
                        except Exception as e:
                            last_error = e
                            time.sleep(1.5)
                            continue

                    if not response or not response.text:
                        raise Exception(f"일시적인 생성 지연: {last_error}")

                    raw_text = response.text.strip()
                    if "```" in raw_text:
                        raw_text = re.sub(r"^```(?:json)?\s*", "", raw_text, flags=re.MULTILINE)
                        raw_text = re.sub(r"\s*```$", "", raw_text, flags=re.MULTILINE)
                    
                    parsed_raw = json.loads(raw_text)
                    
                    # 1. Day 병합 및 기본 정제
                    day_dict = {}
                    for item in parsed_raw:
                        d_num = item.get("day", 1)
                        if d_num not in day_dict:
                            day_dict[d_num] = {
                                "day": d_num,
                                "rows": [],
                                "meals": item.get("meals") or ["-", "-", "-"],
                                "hotel": item.get("hotel") or ""
                            }
                        for r_item in item.get("rows", []):
                            sc = str(r_item.get("schedule") or "").strip()
                            # '석식 후' 단독 행 원천 필터링
                            if sc in ["석식 후", "석식후", "석식 후 호텔 투숙"]:
                                continue
                            if r_item.get("schedule") or r_item.get("city") or r_item.get("transport"):
                                sc_clean = sc.replace("[", "").replace("]", "")
                                r_item["schedule"] = sc_clean
                                day_dict[d_num]["rows"].append(r_item)
                        if item.get("hotel"):
                            day_dict[d_num]["hotel"] = str(item.get("hotel")).replace("[", "").replace("]", "").strip()
                        if item.get("meals"):
                            day_dict[d_num]["meals"] = item.get("meals")

                    parsed_days = [day_dict[k] for k in sorted(day_dict.keys())]

                    # 2. 실무 파이썬 정밀 후처리
                    valid_meals = ["호텔식", "현지식", "한   식", "중국식", "기내식", "-"]
                    turkey_keywords = ["이스탄불", "카파도키아", "파묵칼레", "터키", "튀르키예", "안탈리아"]
                    total_days_cnt = len(parsed_days)

                    for item in parsed_days:
                        d_num = item["day"]
                        rows = item["rows"]
                        
                        cleaned_rows = []
                        for r in rows:
                            sc = str(r.get("schedule") or "").strip()
                            # 불필요한 [도시]관광 단독행 및 석식후 제거
                            if re.match(r"^[가-힣a-zA-Z\s]+관광$", sc) and not sc.startswith("▶"):
                                continue
                            if sc in ["석식 후", "석식후"]:
                                continue
                            cleaned_rows.append(r)
                        rows = cleaned_rows

                        # 2일차 이후 아침 첫 행 조식 정돈
                        if d_num > 1 and len(rows) > 0 and d_num != total_days_cnt:
                            first_sc = rows[0].get("schedule", "")
                            if "조식" in first_sc and len(first_sc) > 8:
                                remainder = first_sc.replace("호텔 조식 후", "").replace("조식 후", "").strip()
                                if remainder.startswith("후"):
                                    remainder = remainder[1:].strip()
                                rows[0]["schedule"] = "호텔 조식 후"
                                rows.insert(1, {"city": "", "transport": "", "time": "", "schedule": remainder})

                        # 도시명 중복 표기 방지
                        if d_num > 1:
                            last_city = None
                            for r in rows:
                                curr_c = str(r.get("city") or "").strip()
                                if curr_c:
                                    if curr_c == last_city:
                                        r["city"] = ""
                                    else:
                                        last_city = curr_c

                        # 전용차량 & 전  일 자동 배치
                        bus_placed = False
                        for r in rows:
                            tm = str(r.get("time") or "").strip()
                            tr = str(r.get("transport") or "").strip()
                            sc = str(r.get("schedule") or "").strip()

                            if sc.startswith("▶"):
                                r["time"] = ""

                            if tr == "전용차량":
                                if not bus_placed:
                                    bus_placed = True
                                    r["time"] = "전  일"
                                else:
                                    r["transport"] = ""
                                    if r.get("time") == "전  일":
                                        r["time"] = ""

                        # 전용차량 표기가 아예 누락된 투어일 보정
                        if d_num > 1 and d_num < total_days_cnt and not bus_placed and len(rows) > 1:
                            rows[1]["transport"] = "전용차량"
                            rows[1]["time"] = "전  일"

                        item["rows"] = rows

                        # 식사 정제
                        cleaned_meals = []
                        for m in item.get("meals", []):
                            m_str = str(m or "").strip()
                            if "호텔" in m_str or "조식" in m_str:
                                cleaned_meals.append("호텔식")
                            elif "한식" in m_str:
                                cleaned_meals.append("한   식")
                            elif "중식" in m_str or "중국" in m_str:
                                cleaned_meals.append("중국식")
                            elif "기내" in m_str:
                                cleaned_meals.append("기내식")
                            elif "현지" in m_str or "특식" in m_str:
                                cleaned_meals.append("현지식")
                            elif m_str in valid_meals:
                                cleaned_meals.append(m_str)
                            else:
                                cleaned_meals.append("현지식" if m_str != "-" else "-")
                        while len(cleaned_meals) < 3:
                            cleaned_meals.append("-")
                        item["meals"] = cleaned_meals[:3]

                        # 마지막 전날: 기내 숙박
                        if d_num == total_days_cnt - 1:
                            item["rows"] = [r for r in rows if "인천 도착" not in r.get("schedule", "")]
                            has_night_flight = any("기내 숙박" in r.get("schedule", "") for r in item["rows"])
                            if not has_night_flight:
                                item["rows"].append({"city": "인  천", "transport": "", "time": "", "schedule": "기내 숙박"})
                            item["hotel"] = "HOTEL : 기내 숙박"
                            if len(item["meals"]) == 3:
                                item["meals"][2] = "기내식"

                        # 마지막 날: 08:55 고정 제거! 실제 항공 도착시간 유지
                        elif d_num == total_days_cnt:
                            orig_time = rows[0].get("time", "") if len(rows) > 0 else ""
                            final_arr_time = orig_time if re.match(r"^\d{1,2}:\d{2}$", orig_time) else ""
                            
                            item["rows"] = [
                                {"city": "인  천", "transport": "", "time": final_arr_time, "schedule": "인천 국제공항 도착 후 해산"}
                            ]
                            item["hotel"] = ""
                            item["meals"] = ["기내식", "-", "-"]
                        else:
                            all_text = " ".join([r.get("schedule", "") + " " + str(r.get("city", "")) for r in rows])
                            is_turkey = any(tk in all_text for tk in turkey_keywords)
                            item["hotel"] = "HOTEL : ANY 5* HOTEL CLASS" if is_turkey else "HOTEL : ANY 4* HOTEL CLASS"

                    # 3. 엑셀 워크북 생성
                    wb = openpyxl.Workbook()
                    ws = wb.active
                    ws.title = "일정표"
                    ws.views.sheetView[0].showGridLines = True
                    
                    font_title = Font(name="맑은 고딕", size=15, bold=True)
                    font_header = Font(name="맑은 고딕", size=10, bold=True)
                    font_body = Font(name="맑은 고딕", size=9)
                    font_bold = Font(name="맑은 고딕", size=9, bold=True)
                    font_notice_red = Font(name="맑은 고딕", size=11, bold=True, color="FF0000")
                    
                    in_font_regular = InlineFont(rFont="맑은 고딕", sz=9.0, b=False)
                    in_font_bold = InlineFont(rFont="맑은 고딕", sz=9.0, b=True)
                    
                    fill_title = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")
                    fill_hotel = PatternFill(start_color="EFEFEF", end_color="EFEFEF", fill_type="solid")
                    
                    align_center = Alignment(horizontal="center", vertical="center")
                    align_mid_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
                    
                    thin_side = Side(style="thin", color="000000")
                    med_side = Side(style="medium", color="000000")
                    
                    full_title = tour_title.strip() if tour_title.strip() else f"{start_date.strftime('%Y-%m-%d')} 유럽투어 일정표"
                    ws.row_dimensions[1].height = 28.0
                    ws.cell(1, 1, full_title).font = font_title
                    ws.cell(1, 1).alignment = align_center
                    
                    for c in range(1, 8):
                        cell = ws.cell(1, c)
                        cell.fill = fill_title
                        cell.border = Border(top=med_side, bottom=med_side, left=med_side if c==1 else None, right=med_side if c==7 else None)
                    ws.merge_cells("A1:G1")
                    
                    ws.row_dimensions[2].height = 13.0
                    
                    headers = ["날짜", "도시", "교통편", "시간", "일정", "식사", ""]
                    ws.row_dimensions[3].height = 22.0
                    for c in range(1, 8):
                        cell = ws.cell(3, c, headers[c-1])
                        cell.font = font_header
                        cell.alignment = align_center
                        cell.border = Border(top=med_side, bottom=med_side, left=med_side if c==1 else thin_side, right=med_side if c==7 else thin_side)
                    ws.merge_cells("F3:G3")
                    
                    col_widths = {1: 10.0, 2: 12.75, 3: 11.0, 4: 9.25, 5: 64.25, 6: 3.63, 7: 10.0}
                    for c_idx, width in col_widths.items():
                        ws.column_dimensions[get_column_letter(c_idx)].width = width
                    
                    days_korean = ["월요일", "화요일", "수요일", "목요일", "금요일", "토요일", "일요일"]
                    cur_row = 4
                    
                    for item_idx, item in enumerate(parsed_days):
                        day_num = item.get("day", item_idx + 1)
                        curr_date = start_date + timedelta(days=day_num - 1)
                        day_str = f"제{day_num:02d}일"
                        date_str = curr_date.strftime("%m월 %d일")
                        weekday_str = days_korean[curr_date.weekday()]
                        
                        rows_data = item.get("rows", [])
                        hotel_str = str(item.get("hotel") or "").strip()
                        meals = item.get("meals") or ["-", "-", "-"]
                        
                        is_last_day = (item_idx == len(parsed_days) - 1)
                        first_row_text = str(rows_data[0].get("schedule", "")) if len(rows_data) > 0 else ""
                        is_incheon_arrival = "인천" in first_row_text and "도착" in first_row_text
                        
                        has_hotel = bool(hotel_str) and not is_last_day and not is_incheon_arrival
                        total_rows = max(len(rows_data), 3) + (1 if has_hotel else 0)
                        
                        for r_idx in range(total_rows):
                            r = cur_row + r_idx
                            ws.row_dimensions[r].height = 18.75
                            is_last_row = (r_idx == total_rows - 1)
                            
                            for c in range(1, 8):
                                cell = ws.cell(r, c)
                                cell.font = font_body
                                
                                b_top = thin_side if (is_last_row and has_hotel) else None
                                b_bottom = med_side if is_last_row else None
                                b_left = med_side if c == 1 else thin_side
                                b_right = med_side if c == 7 else thin_side
                                
                                if is_last_row and has_hotel:
                                    cell.fill = fill_hotel
                                cell.border = Border(top=b_top, bottom=b_bottom, left=b_left, right=b_right)
                        
                        cell_d1 = ws.cell(cur_row, 1, day_str)
                        cell_d1.font = font_bold
                        cell_d1.alignment = align_center
                        
                        cell_d2 = ws.cell(cur_row + 1, 1, date_str)
                        cell_d2.font = font_body
                        cell_d2.alignment = align_center
                        
                        cell_d3 = ws.cell(cur_row + 2, 1, weekday_str)
                        cell_d3.font = font_body
                        cell_d3.alignment = align_center
                        
                        for idx, r_data in enumerate(rows_data):
                            r = cur_row + idx
                            sc_text = str(r_data.get("schedule") or "")
                            
                            if r_data.get("city"):
                                ws.cell(r, 2, str(r_data.get("city") or "")).alignment = align_center
                            if r_data.get("transport"):
                                ws.cell(r, 3, str(r_data.get("transport") or "")).alignment = align_center
                            if r_data.get("time"):
                                ws.cell(r, 4, str(r_data.get("time") or "")).alignment = align_center
                                
                            cell_sc = ws.cell(r, 5)
                            cell_sc.alignment = align_mid_left
                            
                            if sc_text.strip().startswith("▶"):
                                bold_match = re.search(r"^(.*?)\*\*(.+?)\*\*(.*)$", sc_text)
                                if bold_match:
                                    p_txt, b_txt, s_txt = bold_match.group(1), bold_match.group(2), bold_match.group(3)
                                    rich_elems = []
                                    if p_txt: rich_elems.append(TextBlock(in_font_regular, p_txt))
                                    if b_txt: rich_elems.append(TextBlock(in_font_bold, b_txt))
                                    if s_txt: rich_elems.append(TextBlock(in_font_regular, s_txt))
                                    cell_sc.value = CellRichText(rich_elems)
                                else:
                                    parts = sc_text.rsplit(" ", 1)
                                    if len(parts) == 2:
                                        cell_sc.value = CellRichText([
                                            TextBlock(in_font_regular, parts[0] + " "),
                                            TextBlock(in_font_bold, parts[1])
                                        ])
                                    else:
                                        cell_sc.value = sc_text
                                        cell_sc.font = font_body
                            else:
                                cell_sc.value = sc_text
                                cell_sc.font = font_body
                        
                        if has_hotel:
                            h_row = cur_row + total_rows - 1
                            h_cell = ws.cell(h_row, 5, hotel_str)
                            h_cell.font = font_bold
                            h_cell.alignment = align_mid_left
                        
                        meal_labels = ["조", "중", "석"]
                        for m_idx in range(3):
                            ws.cell(cur_row + m_idx, 6, meal_labels[m_idx]).alignment = align_center
                            meal_val = meals[m_idx] if m_idx < len(meals) else "-"
                            ws.cell(cur_row + m_idx, 7, str(meal_val)).alignment = align_center
                        
                        cur_row += total_rows
                    
                    notice_row = cur_row + 1
                    ws.row_dimensions[notice_row].height = 22.0
                    notice_cell = ws.cell(notice_row, 1, "* 상기 일정은 현지 사정에 따라 변경될 수 있습니다 *")
                    notice_cell.font = font_notice_red
                    notice_cell.alignment = align_center
                    ws.merge_cells(start_row=notice_row, start_column=1, end_row=notice_row, end_column=7)
                    
                    output = io.BytesIO()
                    wb.save(output)
                    
                    final_title = tour_title.strip() if tour_title.strip() else "유럽투어"
                    file_name = f"{start_date.strftime('%Y-%m-%d')} {final_title} - 일정표.xlsx"
                    
                    st.session_state.generated_file = output.getvalue()
                    st.session_state.generated_filename = file_name
                    st.session_state.ui_status = "success"
                    st.session_state.ui_msg = f"선배님! 7대 실무 수배 룰(시간 보정·08:55 해제)을 완벽히 적용하여 생성을 마쳤습니다! 🙇‍♂️✨"
                    st.rerun()

                except Exception as e:
                    st.session_state.ui_status = "working"
                    st.session_state.ui_msg = f"오류가 발생했습니다: {str(e)}"
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
