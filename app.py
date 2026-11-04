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

# 1. 페이지 설정
st.set_page_config(page_title="유럽 랜드사 일정표 생성기 (Master Edition)", layout="wide", page_icon="🇪🇺")

# 2. 커스텀 CSS
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

.badge-idle {
    background: #e2e8f0;
    color: #475569;
}

.badge-working {
    background: #fee2e2;
    color: #ef4444;
}

.badge-success {
    background: #dcfce7;
    color: #15803d;
}

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
st.caption("실무 완판 규격: 아시아나 DB ➔ 대한항공 자동 전환, 관광지 설명 일반체/명소명 볼드 분리, 기내 숙박 및 도착 도시 자동 연동")

# 3. database 폴더 내 사내 엑셀 DB 자동 스캔
@st.cache_data(show_spinner=False)
def load_company_database():
    db_folder = "database"
    excel_files = glob.glob(os.path.join(db_folder, "*.xlsx"))
    
    extracted_knowledge = []
    file_list = []
    
    for filepath in excel_files:
        filename = os.path.basename(filepath)
        file_list.append(filename)
        try:
            wb = openpyxl.load_workbook(filepath, data_only=True)
            ws = wb.active
            
            content_summary = []
            for r in range(4, min(ws.max_row + 1, 120)):
                city = ws.cell(r, 2).value
                schedule = ws.cell(r, 5).value
                line_parts = []
                if city:
                    line_parts.append(f"도시: {str(city).strip()}")
                if schedule:
                    line_parts.append(f"내용: {str(schedule).strip()}")
                if line_parts:
                    content_summary.append(" | ".join(line_parts))
            
            extracted_knowledge.append(f"=== [사내 표준 파일: {filename}] ===\n" + "\n".join(content_summary[:80]))
        except Exception:
            continue
            
    return file_list, "\n\n".join(extracted_knowledge)

db_files, db_context = load_company_database()

# Gemini 클라이언트 초기화
api_key = st.secrets.get("GEMINI_API_KEY")
client = genai.Client(api_key=api_key) if api_key else None

col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("1. 입력 방식 선택")
    
    input_mode = st.radio(
        "작성 방식을 선택하세요:",
        ["[옵션 A] 간편 단문 요청 (사내 DB 자동 매칭)", "[옵션 B] 직접 루팅 입력 (정밀 항공/일정 파싱)"],
        horizontal=True
    )
    
    if "[옵션 A]" in input_mode:
        if db_files:
            st.markdown(f"""
            <div class="db-info-box">
                📚 <b>연동된 사내 엑셀 DB ({len(db_files)}개):</b><br>
                {' • '.join([f'`{f}`' for f in db_files])}
            </div>
            """, unsafe_allow_html=True)
        else:
            st.info("💡 `database/` 폴더에 사내 엑셀이 있을 경우 자동 매핑됩니다.")

    start_date = st.date_input("출발일자 선택", datetime.today())
    tour_title = st.text_input("행사명 (선택)", placeholder="예: 2026-12-05 동유럽 9일 KE")
    
    if "[옵션 A]" in input_mode:
        quick_prompt = st.text_input(
            "간단 요청 입력",
            placeholder="예: 동유럽 9일 KE 프라하인 비엔나아웃 / 튀르키예 9일 중간항공 1회"
        )
    else:
        routing_text = st.text_area(
            "루팅 및 일정 입력 (복사-붙여넣기)",
            height=280,
            placeholder="""1. 인천/이스탄불/아테네 (TK021 / TK1845 편) 20:45분 아테네 도착
2. 아테네(관광) - 고린도(관광) - 아테네
3. 아테네 / 이스탄불 (TK1844 12:50분 도착) 후 이스탄불 시내 관광
4. 이스탄불 / 카파도키아 (TK2010 0735-0900) - 카파도키아 관광
..."""
        )

with col2:
    st.subheader("2. AI 수배팀 작업 현황")
    ui_placeholder = st.empty()

    def render_assistant_ui(status="working", message="선배님, 요청 주시면 사내 규격 엑셀 바로 뽑아오겠습니다! ☕"):
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

    render_assistant_ui("idle", "선배님! 대한항공/아시아나 통합 룰 및 완판본 서식으로 준비 완료했습니다! ⌨️")

    generate_btn = st.button("🚀 엑셀 일정표 생성하기", type="primary", use_container_width=True)

    if generate_btn:
        if not api_key:
            st.error("설정(Secrets)에서 GEMINI_API_KEY가 등록되지 않았습니다.")
        elif "[옵션 A]" in input_mode and not quick_prompt.strip():
            st.warning("간단 요청 내용을 입력해 주세요.")
        elif "[옵션 B]" in input_mode and not routing_text.strip():
            st.warning("루팅 정보를 입력해 주세요.")
        else:
            render_assistant_ui("working", "사내 DB 매칭 및 통합 항공 스케줄/완판본 서식을 조립 중입니다! ⚡🔥")
            
            with st.spinner("김똘망 사원이 완판본 규격대로 꼼꼼히 제작하고 있습니다..."):
                try:
                    system_prompt = f"""
당신은 15년 차 유럽 전문 랜드사의 수석 오퍼레이터입니다.
아래 제공된 [사내 표준 엑셀 DB] 및 사용자의 입력을 바탕으로 고객 전달용 완벽한 엑셀 일정표 JSON을 생성하세요.

[사내 표준 엑셀 DB]
{db_context if db_context else "(표준 랜드사 규격 준수)"}

[사용자 입력 요청]
{quick_prompt if "[옵션 A]" in input_mode else routing_text}

[핵심 절대 규칙 - 위반 절대 금지]
1. 항공 루팅 및 대한항공/아시아나 통합 규칙:
   - 사용자가 대한항공(KE) 또는 국적기를 요청한 경우:
     * 참조하는 사내 DB 파일이 아시아나(OZ) 베이스이거나 구형 스케줄이더라도, 항공편명(KE)과 시간표를 현재 대한항공 운항 스케줄 기준으로 보정하여 출력하세요.
     * 아시아나 편명(OZ)은 결과표에 절대 표기하지 마세요.
   - 터키항공(TK) 요청 시 경유 4행 규격 준수:
     * 행1: city="인  천", transport="TK021", time="10:45", schedule="인천 국제공항 출발 / 이스탄불 공항 향발"
     * 행2: city="이스탄불", transport="", time="16:45", schedule="이스탄불 공항 도착 후 환승 대기"
     * 행3: city="", transport="TK1845", time="19:15", schedule="이스탄불 공항 출발 / 아테네 공항 향발"
     * 행4: city="아테네", transport="", time="20:45", schedule="아테네 공항 도착 후 가이드 미팅"
   - 귀국편(마지막 전날): 공항 이동 후 항공편 출발 행 작성, 그 다음 행에 schedule="기내 숙박", hotel="HOTEL : 기내 숙박".
   - 마지막 날(인천 도착일): city="인  천", schedule="인천 국제공항 도착 후 해산", hotel="".

2. 일정 흐름 및 세모표(▶) 직행 규칙:
   - '~로 이동' 행 바로 다음에 쓸데없이 '[도시명] 관광' 같은 빈 껍데기 행을 절대 넣지 마세요!
   - 이동 행 바로 다음 행부터 곧바로 ' ▶ [설명] **[핵심명소]**' 형태로 작성하세요.

3. 관광 명소 서식 표기 규칙 (부분 볼드 처리):
   - 설명 부분은 보통체로, 맨 뒤의 핵심 관광지 명칭만 **기호로 감싸세요!
   - 대괄호([])는 절대로 사용하지 마세요.
   (작성 예시):
   - schedule=" ▶ 세계 3대 운하 중 하나인 **고린도 운하**"
   - schedule=" ▶ 수천 년 전 화산재가 굳어져 형성된 기암괴석의 전시장 **괴레메 야외 박물관**"
   - schedule=" ▶ 서양 문명의 정수이자 유네스코 세계문화유산인 **파르테논 신전**"

4. 도시명(city) 배치 원칙:
   - 일자 첫 행에 시작 도시를 적습니다.
   - 이동 중인 행(schedule="~로 이동")에는 city를 빈칸("")으로 둡니다.
   - 이동 후 도착하여 첫 번째 명소(▶)가 시작되는 행의 city 칸에 도착 도시명을 1회만 기재하세요. (중복 방지)

5. 하루 시작과 마무리:
   - 2일차부터 아침 1행은 무조건 schedule="호텔 조식 후" 단독 작성.
   - 관광 투어가 끝나는 날 마무리는 '석식 후' -> '호텔 투숙 및 휴식' (또는 '호텔로 이동' -> '호텔 투숙 및 휴식').

6. 식사 및 호텔:
   - 식사는 "호텔식", "현지식", "한   식", "중국식", "기내식", "-" 중에서만 기재.
   - 호텔: 튀르키예 지역은 "HOTEL : ANY 5* HOTEL CLASS", 그 외 유럽 지역은 "HOTEL : ANY 4* HOTEL CLASS".
   - 귀국 출발일은 "HOTEL : 기내 숙박", 마지막 날은 "" (빈칸).
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

                    candidate_models = ['gemini-2.5-pro', 'gemini-3.5-flash-lite']
                    response = None
                    last_error = None

                    for model_name in candidate_models:
                        try:
                            response = client.models.generate_content(
                                model=model_name,
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
                            time.sleep(1.2)
                            continue

                    if not response or not response.text:
                        raise Exception(f"일시적인 AI 서버 지연이 발생했습니다: {last_error}")

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
                            if r_item.get("schedule") or r_item.get("city") or r_item.get("transport"):
                                sc_clean = str(r_item.get("schedule") or "").replace("[", "").replace("]", "").strip()
                                r_item["schedule"] = sc_clean
                                day_dict[d_num]["rows"].append(r_item)
                        if item.get("hotel"):
                            day_dict[d_num]["hotel"] = str(item.get("hotel")).replace("[", "").replace("]", "").strip()
                        if item.get("meals"):
                            day_dict[d_num]["meals"] = item.get("meals")

                    parsed_days = [day_dict[k] for k in sorted(day_dict.keys())]

                    # 2. 강력한 실무 파이썬 후처리
                    valid_meals = ["호텔식", "현지식", "한   식", "중국식", "기내식", "-"]
                    turkey_keywords = ["이스탄불", "카파도키아", "파묵칼레", "마니사", "쉬린제", "에페소", "터키", "튀르키예", "카이세리", "네브셰히르", "안탈리아", "안탈라", "아피온", "콘야", "이즈미르"]

                    total_days_cnt = len(parsed_days)

                    for item in parsed_days:
                        d_num = item["day"]
                        rows = item["rows"]
                        
                        # (1) 불필요한 '[도시] 관광' 행 제거
                        cleaned_rows = []
                        for r in rows:
                            sc = str(r.get("schedule") or "").strip()
                            if re.match(r"^[가-힣a-zA-Z\s]+관광$", sc) and not sc.startswith("▶"):
                                continue
                            cleaned_rows.append(r)
                        rows = cleaned_rows
                        item["rows"] = rows

                        # (2) 조식 단독 행 분리
                        if d_num > 1 and len(rows) > 0 and d_num != total_days_cnt:
                            first_sc = rows[0].get("schedule", "")
                            if "조식" in first_sc and len(first_sc) > 8:
                                remainder = first_sc.replace("호텔 조식 후", "").replace("조식 후", "").strip()
                                if remainder.startswith("후"):
                                    remainder = remainder[1:].strip()
                                rows[0]["schedule"] = "호텔 조식 후"
                                rows.insert(1, {"city": "", "transport": "", "time": "", "schedule": remainder})

                        # (3) 1일차 터키항공(TK) 경유 4단 규격 강제 보정 (사용자 요청이 TK인 경우)
                        req_text = (quick_prompt if "[옵션 A]" in input_mode else routing_text).upper()
                        if d_num == 1 and ("TK" in req_text or "터키항공" in req_text):
                            rows = [
                                {"city": "인  천", "transport": "TK021", "time": "10:45", "schedule": "인천 국제공항 출발 / 이스탄불 공항 향발"},
                                {"city": "이스탄불", "transport": "", "time": "16:45", "schedule": "이스탄불 공항 도착 후 환승 대기"},
                                {"city": "", "transport": "TK1845", "time": "19:15", "schedule": "이스탄불 공항 출발 / 아테네 공항 향발"},
                                {"city": "아테네", "transport": "", "time": "20:45", "schedule": "아테네 공항 도착 후 가이드 미팅"},
                                {"city": "", "transport": "", "time": "", "schedule": "호텔로 이동"},
                                {"city": "", "transport": "", "time": "", "schedule": "호텔 투숙 및 휴식"}
                            ]
                            item["rows"] = rows
                            item["meals"] = ["기내식", "기내식", "-"]
                            item["hotel"] = "HOTEL : ANY 4* HOTEL CLASS"

                        # (4) 도시 중복 제거
                        if d_num > 1:
                            last_city = None
                            for r in rows:
                                curr_c = str(r.get("city") or "").strip()
                                if curr_c:
                                    if curr_c == last_city:
                                        r["city"] = ""
                                    else:
                                        last_city = curr_c

                        # (5) 시간 및 전용차량 정돈
                        bus_placed = False
                        for r in rows:
                            tm = str(r.get("time") or "").strip()
                            tr = str(r.get("transport") or "").strip()
                            sc = str(r.get("schedule") or "").strip()

                            if not re.match(r"^\d{1,2}:\d{2}$", tm) and tm not in ["전  일", "오  전", "오  후"]:
                                r["time"] = ""

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

                        # (6) 식사 4종 규격화
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
                            elif "현지" in m_str or "외식" in m_str or "특식" in m_str or "케밥" in m_str:
                                cleaned_meals.append("현지식")
                            elif m_str in valid_meals:
                                cleaned_meals.append(m_str)
                            else:
                                cleaned_meals.append("현지식" if m_str != "-" else "-")
                        while len(cleaned_meals) < 3:
                            cleaned_meals.append("-")
                        item["meals"] = cleaned_meals[:3]

                        # (7) 귀국일(마지막 전날) 및 마지막 날 고정
                        if d_num == total_days_cnt - 1:
                            item["rows"] = [r for r in rows if "인천 도착" not in r.get("schedule", "")]
                            has_night_flight = any("기내 숙박" in r.get("schedule", "") for r in item["rows"])
                            if not has_night_flight:
                                item["rows"].append({"city": "인  천", "transport": "", "time": "", "schedule": "기내 숙박"})
                            item["hotel"] = "HOTEL : 기내 숙박"
                            if len(item["meals"]) == 3:
                                item["meals"][2] = "기내식"

                        elif d_num == total_days_cnt:
                            item["rows"] = [
                                {"city": "인  천", "transport": "", "time": "08:55", "schedule": "인천 국제공항 도착 후 해산"}
                            ]
                            item["hotel"] = ""
                            item["meals"] = ["기내식", "-", "-"]
                        else:
                            all_text = " ".join([r.get("schedule", "") + " " + str(r.get("city", "")) for r in rows])
                            is_turkey = any(tk in all_text for tk in turkey_keywords)
                            if is_turkey:
                                item["hotel"] = "HOTEL : ANY 5* HOTEL CLASS"
                            else:
                                item["hotel"] = "HOTEL : ANY 4* HOTEL CLASS"

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
                    
                    # In-line 부분 서식을 위한 전용 폰트
                    in_font_regular = InlineFont(rFont="맑은 고딕", sz=9.0, b=False)
                    in_font_bold = InlineFont(rFont="맑은 고딕", sz=9.0, b=True)
                    
                    fill_title = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")
                    fill_hotel = PatternFill(start_color="EFEFEF", end_color="EFEFEF", fill_type="solid")
                    
                    align_center = Alignment(horizontal="center", vertical="center")
                    align_mid_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
                    
                    thin_side = Side(style="thin", color="000000")
                    med_side = Side(style="medium", color="000000")
                    
                    # 1행 타이틀바
                    full_title = tour_title.strip() if tour_title.strip() else f"{start_date.strftime('%Y-%m-%d')} 유럽투어 일정표"
                    ws.row_dimensions[1].height = 28.0
                    ws.cell(1, 1, full_title).font = font_title
                    ws.cell(1, 1).alignment = align_center
                    
                    for c in range(1, 8):
                        cell = ws.cell(1, c)
                        cell.fill = fill_title
                        cell.border = Border(top=med_side, bottom=med_side, left=med_side if c==1 else None, right=med_side if c==7 else None)
                    ws.merge_cells("A1:G1")
                    
                    # 2행 간격행
                    ws.row_dimensions[2].height = 13.0
                    
                    # 3행 헤더행
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
                    total_days_count = len(parsed_days)
                    
                    for item_idx, item in enumerate(parsed_days):
                        day_num = item.get("day", item_idx + 1)
                        curr_date = start_date + timedelta(days=day_num - 1)
                        day_str = f"제{day_num:02d}일"
                        date_str = curr_date.strftime("%m월 %d일")
                        weekday_str = days_korean[curr_date.weekday()]
                        
                        rows_data = item.get("rows", [])
                        hotel_str = str(item.get("hotel") or "").strip()
                        meals = item.get("meals") or ["-", "-", "-"]
                        
                        is_last_day = (item_idx == total_days_count - 1)
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
                        
                        # 1열 날짜/요일 기입
                        cell_d1 = ws.cell(cur_row, 1, day_str)
                        cell_d1.font = font_bold
                        cell_d1.alignment = align_center
                        
                        cell_d2 = ws.cell(cur_row + 1, 1, date_str)
                        cell_d2.font = font_body
                        cell_d2.alignment = align_center
                        
                        cell_d3 = ws.cell(cur_row + 2, 1, weekday_str)
                        cell_d3.font = font_body
                        cell_d3.alignment = align_center
                        
                        # 2~5열 일정
                        for idx, r_data in enumerate(rows_data):
                            r = cur_row + idx
                            sc_text = str(r_data.get("schedule") or "")
                            
                            if r_data.get("city"):
                                cell_city = ws.cell(r, 2, str(r_data.get("city") or ""))
                                cell_city.alignment = align_center
                            if r_data.get("transport"):
                                cell_tr = ws.cell(r, 3, str(r_data.get("transport") or ""))
                                cell_tr.alignment = align_center
                            if r_data.get("time"):
                                cell_tm = ws.cell(r, 4, str(r_data.get("time") or ""))
                                cell_tm.alignment = align_center
                                
                            cell_sc = ws.cell(r, 5)
                            cell_sc.alignment = align_mid_left
                            
                            # 관광 명소 앞부분 일반체, 뒤의 명소명만 볼드체 렌더링
                            if sc_text.strip().startswith("▶"):
                                bold_match = re.search(r"^(.*?)\*\*(.+?)\*\*(.*)$", sc_text)
                                if bold_match:
                                    prefix_text = bold_match.group(1)
                                    bold_target = bold_match.group(2)
                                    suffix_text = bold_match.group(3)
                                    
                                    rich_text_elements = []
                                    if prefix_text:
                                        rich_text_elements.append(TextBlock(in_font_regular, prefix_text))
                                    if bold_target:
                                        rich_text_elements.append(TextBlock(in_font_bold, bold_target))
                                    if suffix_text:
                                        rich_text_elements.append(TextBlock(in_font_regular, suffix_text))
                                    
                                    cell_sc.value = CellRichText(rich_text_elements)
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
                        
                        # 호텔
                        if has_hotel:
                            h_row = cur_row + total_rows - 1
                            h_cell = ws.cell(h_row, 5, hotel_str)
                            h_cell.font = font_bold
                            h_cell.alignment = align_mid_left
                        
                        # 식사
                        meal_labels = ["조", "중", "석"]
                        for m_idx in range(3):
                            ws.cell(cur_row + m_idx, 6, meal_labels[m_idx]).alignment = align_center
                            meal_val = meals[m_idx] if m_idx < len(meals) else "-"
                            ws.cell(cur_row + m_idx, 7, str(meal_val)).alignment = align_center
                        
                        cur_row += total_rows
                    
                    # 최하단 안내 문구: 한 줄 띄고 11pt 붉은색 볼드 가운데 정렬
                    notice_row = cur_row + 1
                    ws.row_dimensions[notice_row].height = 22.0
                    notice_cell = ws.cell(notice_row, 1, "* 상기 일정은 현지 사정에 따라 변경될 수 있습니다 *")
                    notice_cell.font = font_notice_red
                    notice_cell.alignment = align_center
                    ws.merge_cells(start_row=notice_row, start_column=1, end_row=notice_row, end_column=7)
                    
                    output = io.BytesIO()
                    wb.save(output)
                    output.seek(0)
                    
                    render_assistant_ui("success", "선배님! 요청하신 통합 항공 스케줄과 완판본 서식으로 엑셀 제작을 완료했습니다! 🙇‍♂️✨")
                    
                    final_title = tour_title.strip() if tour_title.strip() else "유럽투어"
                    file_name = f"{start_date.strftime('%Y-%m-%d')} {final_title} - 일정표.xlsx"
                    
                    st.download_button(
                        label="📥 완성된 실무 마스터 엑셀 일정표 다운로드",
                        data=output,
                        file_name=file_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True
                    )
                except Exception as e:
                    render_assistant_ui("working", f"오류가 발생했습니다: {str(e)}")
                    st.error(f"오류 상세: {str(e)}")
