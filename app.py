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
st.caption("사내 표준 엑셀 DB를 기반으로 단문 요청 및 정밀 루팅을 규격 일정표로 변환합니다.")

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
                hotel = ws.cell(r, 5).value if "HOTEL" in str(ws.cell(r, 5).value or "") else ""
                
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
    
    if db_files:
        st.markdown(f"""
        <div class="db-info-box">
            📚 <b>탑재된 사내 표준 엑셀 DB ({len(db_files)}개):</b><br>
            {' • '.join([f'`{f}`' for f in db_files])}
        </div>
        """, unsafe_allow_html=True)
    else:
        st.info("💡 `database/` 폴더에 사내 엑셀 파일을 추가하면 자동으로 지식 DB가 확장됩니다.")

    input_mode = st.radio(
        "원하시는 작성 방식을 선택하세요:",
        ["[옵션 A] 간편 단문 요청 (사내 DB 자동 매칭)", "[옵션 B] 직접 루팅 입력 (기존 방식)"],
        horizontal=True
    )
    
    start_date = st.date_input("출발일자 선택", datetime.today())
    tour_title = st.text_input("행사명 (선택)", placeholder="예: 2026-12-05 SM투어) 동유럽 3국 9일")
    
    if "[옵션 A]" in input_mode:
        quick_prompt = st.text_input(
            "간단 요청 입력",
            placeholder="예: 동유럽 9일인데 프라하 2박 들어가게 짜줘 / 스포 9일짜리"
        )
        routing_input_text = f"[사용자 간편 요청]: {quick_prompt}"
    else:
        routing_text = st.text_area(
            "루팅 및 일정 입력 (복사-붙여넣기)",
            height=280,
            placeholder="""1. 인천/프라하 (항공편)
2. 프라하(시내 관광)
3. 프라하-체스키크룸로프-비엔나
4. 비엔나(쉔부른 궁전 관광)
5. 비엔나-부다페스트
..."""
        )
        routing_input_text = f"[사용자 직접 입력 루팅]:\n{routing_text}"

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

    render_assistant_ui("working", "선배님! 1안이든 2안이든 사내 엑셀 DB 싹 훑어서 완벽히 만들어 올리겠습니다! ⌨️")

    generate_btn = st.button("🚀 엑셀 일정표 생성하기", type="primary", use_container_width=True)

    if generate_btn:
        if not api_key:
            st.error("설정(Secrets)에서 GEMINI_API_KEY가 등록되지 않았습니다.")
        elif "[옵션 A]" in input_mode and not quick_prompt.strip():
            st.warning("간단 요청 내용을 입력해 주세요.")
        elif "[옵션 B]" in input_mode and not routing_text.strip():
            st.warning("루팅 정보를 입력해 주세요.")
        else:
            render_assistant_ui("working", "사내 엑셀 DB 대조하면서 안경 고쳐 쓰고 정밀 작업 들어갑니다! ⌨️🔥")
            
            with st.spinner("김똘망 사원이 사내 DB를 분석하여 규격 엑셀을 제작 중입니다..."):
                try:
                    system_prompt = f"""
당신은 15년 차 유럽 전문 랜드사의 수석 오퍼레이터입니다.
아래에 제공된 [사내 표준 엑셀 DB]를 최우선 기준으로 삼아 요청받은 일정을 정밀 JSON 배열로 출력하세요.

[사내 표준 엑셀 DB]
{db_context if db_context else "(등록된 사내 엑셀 DB 없음 - 랜드사 표준으로 생성)"}

[입력 요청]
{routing_input_text}

[사내 DB 반영 및 작성 절대 규칙]
1. 사내 DB 지식 최우선 적용:
   - 사용자가 '간편 요청(옵션 A)'을 한 경우: 사내 DB 파일 중 가장 부합하는 템플릿의 전체 동선과 일자 구성을 뼈대로 채택하세요.
   - 사용자가 '직접 루팅 입력(옵션 B)'을 한 경우: 입력된 동선을 따르되, 방문 도시의 관광지 문구(`▶ ~한 ...`)는 사내 DB의 표현과 필수 입장지를 그대로 가져와 적용하세요.
2. 도시명(city) 표기 규칙:
   - 같은 도시의 일정이 계속 이어지는 동안에는 city 칸을 무조건 빈칸("")으로 둡니다. 도시가 바뀔 때만 최초 1회 기재합니다.
3. 시간란(time) 표기 규칙:
   - 항공편 출발/도착 시간(HH:MM)과 조식 후 첫 관광 시작 행의 '전  일'만 기재합니다. 나머지는 무조건 빈칸("")입니다.
4. 교통편('전용차량') 규칙:
   - 일반 육로 관광일의 2행(첫 관광/이동 시작 행)에만 transport="전용차량", time="전  일"을 단 1회 표기합니다.
5. 관광지 서식:
   - 모든 대표 명소는 ' ▶ ~한 [명소명]' 형태의 개별 행으로 작성하세요.
6. 식사 및 호텔:
   - 식사는 호텔식/현지식 원칙. 호텔은 마지막 행="호텔 투숙 및 휴식", hotel="HOTEL : ANY 4* HOTEL CLASS" 규격을 따르세요.

[출력 포맷: JSON 배열만 출력]
[
  {{
    "day": 1,
    "rows": [
      {{"city": "인천", "transport": "", "time": "", "schedule": "인천 국제공항 제1여객터미널 집결"}},
      {{"city": "", "transport": "항공편명", "time": "23:35", "schedule": "인천 출발"}}
    ],
    "meals": ["-", "-", "기내식"],
    "hotel": "HOTEL : 기내 숙박"
  }}
]
"""
                    response = client.models.generate_content(
                        model='gemini-3.5-flash-lite',
                        contents=system_prompt,
                        config=types.GenerateContentConfig(
                            response_mime_type="application/json"
                        )
                    )
                    
                    raw_text = response.text.strip()
                    if "```" in raw_text:
                        raw_text = re.sub(r"^```(?:json)?\s*", "", raw_text, flags=re.MULTILINE)
                        raw_text = re.sub(r"\s*```$", "", raw_text, flags=re.MULTILINE)
                    
                    parsed_days = json.loads(raw_text)
                    
                    # 엑셀 워크북 생성
                    wb = openpyxl.Workbook()
                    ws = wb.active
                    ws.title = "일정표"
                    ws.views.sheetView[0].showGridLines = True
                    
                    font_title = Font(name="맑은 고딕", size=15, bold=True)
                    font_header = Font(name="맑은 고딕", size=10, bold=True)
                    font_body = Font(name="맑은 고딕", size=9)
                    font_bold = Font(name="맑은 고딕", size=9, bold=True)
                    
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
                    col_widths = {1: 9.25, 2: 12.75, 3: 7.5, 4: 9.25, 5: 64.25, 6: 3.63, 7: 10.0}
                    for c_idx, width in col_widths.items():
                        ws.column_dimensions[get_column_letter(c_idx)].width = width
                    
                    days_korean = ["월요일", "화요일", "수요일", "목요일", "금요일", "토요일", "일요일"]
                    cur_row = 4
                    
                    for item in parsed_days:
                        day_num = item.get("day", 1)
                        curr_date = start_date + timedelta(days=day_num - 1)
                        day_str = f"제{day_num:02d}일"
                        date_str = curr_date.strftime("%m월 %d일")
                        weekday_str = days_korean[curr_date.weekday()]
                        
                        rows_data = item.get("rows", [])
                        hotel_str = item.get("hotel", "")
                        meals = item.get("meals", ["-", "-", "-"])
                        
                        has_hotel = bool(hotel_str.strip())
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
                                cell_city = ws.cell(r, 2, r_data.get("city"))
                                cell_city.alignment = align_center
                            if r_data.get("transport"):
                                cell_tr = ws.cell(r, 3, r_data.get("transport"))
                                cell_tr.alignment = align_center
                            if r_data.get("time"):
                                cell_tm = ws.cell(r, 4, r_data.get("time"))
                                cell_tm.alignment = align_center
                            cell_sc = ws.cell(r, 5, r_data.get("schedule", ""))
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
                            ws.cell(cur_row + m_idx, 7, meals[m_idx] if m_idx < len(meals) else "-").alignment = align_center
                        
                        cur_row += total_rows
                    
                    output = io.BytesIO()
                    wb.save(output)
                    output.seek(0)
                    
                    render_assistant_ui("success", "선배님! 사내 표준 엑셀 DB와 완벽 동기화하여 제작 완료했습니다! 🙇‍♂️✨")
                    
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
                    render_assistant_ui("working", f"헉 선배님 오류가 났습니다 ㅠㅠ: {str(e)}")
                    st.error(f"오류 상세: {str(e)}")
