import streamlit as st
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta
import io
import json
import os
import re
import glob
import base64
from google import genai
from google.genai import types

# 1. 페이지 설정
st.set_page_config(page_title="유럽 랜드사 일정표 생성기 (Phase 2)", layout="wide", page_icon="🇪🇺")

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

st.title("🇪🇺 유럽 일정표 자동 생성 시스템 (Phase 2)")
st.caption("사내 표준 DB 매칭과 정밀 항공 루팅 생성 모드를 완전히 분리하여 운영합니다.")

# 3. database 폴더 내 사내 엑셀 DB 자동 스캔 함수
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
    tour_title = st.text_input("행사명 (선택)", placeholder="예: 2026-12-05 터키/그리스 9일 TK")
    
    if "[옵션 A]" in input_mode:
        quick_prompt = st.text_input(
            "간단 요청 입력",
            placeholder="예: 동유럽 9일인데 프라하 2박 들어가게 짜줘 / 스포 9일짜리"
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

    render_assistant_ui("idle", "선배님! 옵션에 맞춰 최적화된 엔진으로 작업 준비 완료했습니다! ⌨️")

    generate_btn = st.button("🚀 엑셀 일정표 생성하기", type="primary", use_container_width=True)

    if generate_btn:
        if not api_key:
            st.error("설정(Secrets)에서 GEMINI_API_KEY가 등록되지 않았습니다.")
        elif "[옵션 A]" in input_mode and not quick_prompt.strip():
            st.warning("간단 요청 내용을 입력해 주세요.")
        elif "[옵션 B]" in input_mode and not routing_text.strip():
            st.warning("루팅 정보를 입력해 주세요.")
        else:
            render_assistant_ui("working", "입력 데이터 정밀 분석 및 행별 서식 매핑 작업 진행 중입니다! ⌨️🔥")
            
            with st.spinner("김똘망 사원이 규격 엑셀을 꼼꼼하게 제작 중입니다..."):
                try:
                    if "[옵션 A]" in input_mode:
                        system_prompt = f"""
당신은 15년 차 유럽 전문 랜드사의 수석 오퍼레이터입니다.
아래 제공된 [사내 표준 엑셀 DB]를 최우선 기준으로 삼아 요청받은 일정을 생성하세요.

[사내 표준 엑셀 DB]
{db_context if db_context else "(사내 DB 없음 - 표준 규격 준수)"}

[사용자 요청]
{quick_prompt}

[작성 규칙]
1. 사내 DB 파일 중 가장 부합하는 템플릿의 전체 동선, 도시별 명소 문구(`▶ ~한 ...`), 식사, 호텔 규격을 그대로 적용하세요.
2. 도시명(city)은 도시가 바뀔 때만 첫 행에 적고, 이어지는 일정은 빈칸("")으로 둡니다.
3. 전일 육로 관광일의 2행에만 transport="전용차량", time="전  일"을 단 1회 표기합니다.
4. 마지막 날 인천 도착 후에는 '자택 휴식' 등의 불필요한 행을 절대 생성하지 마세요.
"""
                    else:
                        system_prompt = f"""
당신은 15년 차 유럽 전문 랜드사의 수석 오퍼레이터입니다.
사용자가 입력한 아래 루팅 정보를 바탕으로, 대형 여행사 제출용 최고급 상세 일정표를 생성하세요.
사내 DB에 얽매이지 말고, 텍스트에 적힌 항공 편명/시간/경유 동선을 100% 보존하고 각 도시 명소를 풍성하게 확장하세요.

[사용자 입력 루팅 원본]
{routing_text}

[핵심 작성 및 파싱 절대 규칙 - 위반 금지]
1. 날짜(day) 단위 통합:
   - 각 일자(제1일, 제2일...)는 반드시 하나의 객체로 만들고, 그 날의 모든 일정 행들을 'rows' 배열 안에 순서대로 넣으세요. 절대로 1개 일정을 여러 day로 쪼개지 마세요.
2. 항공 스케줄(편명/시간/경유) 원본 100% 보존 및 매핑:
   - 텍스트에 명시된 항공 편명(예: TK021, TK1845, TK1844, TK2010, TK020)은 반드시 해당 항공 이동 행의 transport 칸에 기재하세요.
   - 시간(time) 칸에는 사용자가 기재한 '실제 시각(예: 20:45, 12:50, 07:35, 09:00, 17:05, 08:55)'만 정확히 적으세요.
   - '오전', '오후', '저녁', '늦은 오후' 같은 단어는 time 칸에 절대 넣지 마세요 (무조건 빈칸("") 처리).
   - 1일차 경유편의 경우:
     * 행1: city="인천", transport="", time="", schedule="인천 국제공항 집결 후 출국 수속"
     * 행2: city="", transport="TK021", time="", schedule="인천 출발 / 이스탄불 향발 (비행시간 약 11시간 30분)"
     * 행3: city="이스탄불", transport="", time="", schedule="이스탄불 공항 도착 후 환승 대기"
     * 행4: city="", transport="TK1845", time="", schedule="이스탄불 출발 / 아테네 향발"
     * 행5: city="아테네", transport="", time="20:45", schedule="아테네 공항 도착 후 가이드 미팅"
     * 행6: city="", transport="", time="", schedule="호텔 투숙 및 휴식"
3. 교통편('전용차량') 및 시간('전  일') 표기 엄격 제한:
   - 전일 육로 관광일의 경우: 오직 그 날의 2번째 행(조식 후 첫 이동/관광 시작 행)에만 단 1회 transport="전용차량", time="전  일"을 적습니다.
   - 같은 날 다른 행에 transport="전용차량"을 중복으로 적거나 time에 '오후', '저녁'을 적는 행위를 금지합니다.
4. 도시별 지리적 정확성 및 명소 풍성도 (각 도시마다 최소 3~5개 명소 필수):
   - 각 도시/지역마다 대표 유료/무료 명소를 ' ▶ ~한 [명소명](내부/외관/조망)' 형식으로 풍성하게 개별 행으로 작성하세요.
   - 타 도시의 명소를 엉뚱하게 섞지 마세요 (예: 아테네 일정에 이스탄불 성소피아를 넣는 행위 절대 금지!).
   - 그리스 명소 기준: 파르테논 신전(내부), 에레크테이온 신전, 아크로폴리스 언덕, 근대 올림픽 경기장(조망), 신타그마 광장, 플라카 지구 등.
   - 고린도 명소 기준: 고린도 운하(조망), 고대 고린도 유적지 및 아폴론 신전 등.
   - 이스탄불 명소 기준: 블루모스크(내부), 아야소피아(외관 또는 내부), 톱카프 궁전(외관), 보스포러스 해협 유람선, 히포드롬 광장, 그랜드 바자르 등.
   - 카파도키아 명소 기준: 괴레메 야외 박물관(내부), 데린쿠유 지하도시(내부), 우치히사르 요새(조망), 파샤바 계곡, 데브란트 계곡 등.
   - 파묵칼레/에페소/쉬린제 명소 기준: 파묵칼레 석회붕 및 히에라폴리스, 에페소 켈수스 도서관, 원형 대극장, 하드리아누스 신전, 쉬린제 마을 등.
5. 마지막 날 인천 도착 규격:
   - 마지막 날 rows는 1개 행만 작성: city="인천", transport="", time="08:55", schedule="인천 국제공항 도착 후 해산"
   - hotel은 빈칸("") 처리.
"""

                    # 엄격한 JSON 스키마 강제 정의 (구조 변형 원천 방지)
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

                    response = client.models.generate_content(
                        model='gemini-3.5-flash-lite',
                        contents=system_prompt,
                        config=types.GenerateContentConfig(
                            response_mime_type="application/json",
                            response_schema=json_schema
                        )
                    )
                    
                    raw_text = response.text.strip()
                    if "```" in raw_text:
                        raw_text = re.sub(r"^```(?:json)?\s*", "", raw_text, flags=re.MULTILINE)
                        raw_text = re.sub(r"\s*```$", "", raw_text, flags=re.MULTILINE)
                    
                    parsed_raw = json.loads(raw_text)
                    
                    # Day 병합 보정 로직 (혹시라도 분리되어 반환된 day들을 안전하게 합침)
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
                            # 유효한 내용이 있는 행만 추가
                            if r_item.get("schedule") or r_item.get("city") or r_item.get("transport"):
                                day_dict[d_num]["rows"].append(r_item)
                        if item.get("hotel"):
                            day_dict[d_num]["hotel"] = item.get("hotel")
                        if item.get("meals") and any(m != "-" for m in item.get("meals")):
                            day_dict[d_num]["meals"] = item.get("meals")

                    parsed_days = [day_dict[k] for k in sorted(day_dict.keys())]
                    
                    # 엑셀 워크북 생성
                    wb = openpyxl.Workbook()
                    ws = wb.active
                    ws.title = "일정표"
                    ws.views.sheetView[0].showGridLines = True
                    
                    font_title = Font(name="맑은 고딕", size=15, bold=True)
                    font_header = Font(name="맑은 고딕", size=10, bold=True)
                    font_body = Font(name="맑은 고딕", size=9)
                    font_bold = Font(name="맑은 고딕", size=9, bold=True)
                    font_notice = Font(name="맑은 고딕", size=9, italic=True, color="555555")
                    
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
                    
                    # 열 너비 설정
                    col_widths = {1: 9.25, 2: 12.75, 3: 11.0, 4: 9.25, 5: 64.25, 6: 3.63, 7: 10.0}
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
                        
                        # 마지막 날 인천 도착 시 호텔칸 제외
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
                        
                        # 1열 날짜
                        ws.cell(cur_row, 1, day_str).font = font_bold
                        ws.cell(cur_row, 1).alignment = align_center
                        ws.cell(cur_row + 1, 1, date_str).alignment = align_center
                        ws.cell(cur_row + 2, 1, weekday_str).alignment = align_center
                        
                        # 2~5열 일정
                        for idx, r_data in enumerate(rows_data):
                            r = cur_row + idx
                            if r_data.get("city"):
                                cell_city = ws.cell(r, 2, str(r_data.get("city") or ""))
                                cell_city.alignment = align_center
                            if r_data.get("transport"):
                                cell_tr = ws.cell(r, 3, str(r_data.get("transport") or ""))
                                cell_tr.alignment = align_center
                            if r_data.get("time"):
                                cell_tm = ws.cell(r, 4, str(r_data.get("time") or ""))
                                cell_tm.alignment = align_center
                            cell_sc = ws.cell(r, 5, str(r_data.get("schedule") or ""))
                            cell_sc.alignment = align_mid_left
                        
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
                    
                    # 최하단 안내 문구 추가
                    notice_row = cur_row + 1
                    ws.row_dimensions[notice_row].height = 20.0
                    notice_cell = ws.cell(notice_row, 1, "* 상기 일정은 현지 사정에 따라 변경될 수 있습니다 *")
                    notice_cell.font = font_notice
                    notice_cell.alignment = align_center
                    ws.merge_cells(start_row=notice_row, start_column=1, end_row=notice_row, end_column=7)
                    
                    output = io.BytesIO()
                    wb.save(output)
                    output.seek(0)
                    
                    render_assistant_ui("success", "선배님! JSON 스키마를 고정하여 날짜 쪼개짐과 빈칸 버그를 완벽히 해결했습니다! 🙇‍♂️✨")
                    
                    final_title = tour_title.strip() if tour_title.strip() else "유럽투어"
                    file_name = f"{start_date.strftime('%Y-%m-%d')} {final_title} - 일정표.xlsx"
                    
                    st.download_button(
                        label="📥 완성된 사내 표준 엑셀 일정표 다운로드",
                        data=output,
                        file_name=file_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True
                    )
                except Exception as e:
                    render_assistant_ui("working", f"오류가 발생했습니다: {str(e)}")
                    st.error(f"오류 상세: {str(e)}")
