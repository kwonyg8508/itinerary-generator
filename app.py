import streamlit as st
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta
import io
import json
import os
import re
import base64
from google import genai
from google.genai import types

# 페이지 기본 설정
st.set_page_config(page_title="유럽 랜드사 일정표 생성기", layout="wide", page_icon="🇪🇺")

# 프리미엄 픽셀 UI 커스텀 스타일
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
    padding: 18px;
    box-shadow: 0 4px 20px rgba(0, 0, 0, 0.05);
    display: flex;
    align-items: center;
    gap: 20px;
    margin-bottom: 20px;
}

.gif-wrapper {
    width: 140px;
    height: 120px;
    flex-shrink: 0;
    border-radius: 12px;
    overflow: hidden;
    border: 2px solid #cbd5e1;
    background: #0f172a;
    display: flex;
    align-items: center;
    justify-content: center;
    box-shadow: inset 0 2px 4px rgba(0,0,0,0.2);
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
    padding: 14px 18px;
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
    padding: 2px 8px;
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
</style>
""", unsafe_allow_html=True)

st.title("🇪🇺 유럽 일정표 자동 생성 시스템")
st.caption("출발일과 루팅만 입력하면 사내 표준 서식 엑셀 파일로 바로 변환됩니다.")

# Gemini 클라이언트 초기화
api_key = st.secrets.get("GEMINI_API_KEY")
client = genai.Client(api_key=api_key) if api_key else None

col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("1. 기본 정보 입력")
    start_date = st.date_input("출발일자 선택", datetime.today())
    tour_title = st.text_input("행사명", placeholder="예: 2026-12-05 SM투어) 북유럽 4국 10일 SK")
    
    routing_text = st.text_area(
        "루팅 및 일정 입력 (복사-붙여넣기)",
        height=320,
        placeholder="""1. 인천/코펜하겐 (SK988, 23:35출발)
2. 코펜하겐(관광, 뉘하운 유람선 탑승 필수)-DFDS탑승 - 오슬로 이동
3. 오슬로(관광, 시청사 내부 필수)-게일로
4. 게일로-베르겐(관광, 플뢰엔산 전망대 필수)
5. 베르겐-구드방겐(송네피요르드 유람선 탑승)-플람하선(플람열차 탑승)-플라로 이동
6. 플라-오슬로(관광, 뭉크미술관 필수) - 칼스타드
7. 칼스타드-스톡홀름(관광, 드로트닝홀름 궁전, 바사박물관 필수)
8. 스톡홀름(관광, 스톡홀름 시청사 및 왕실 무기고 필수) - 실자라인 탑승 - 헬싱키로 이동
9. 헬싱키(하선 후 관광) - 공항이동 - 헬싱키/코펜하겐/인천 (SK1717 / SK987 스케쥴)
10. 인천도착"""
    )

with col2:
    st.subheader("2. AI 수배팀 작업 현황")
    
    # GIF/이미지 파일을 직접 인코딩하여 캐싱 없이 무한 재생 보장
    found_media = None
    for c in ["assistant.gif", "assistant.png", "assitant.png", "assistant.jpeg", "assistant.jpg"]:
        if os.path.exists(c):
            found_media = c
            break

    media_b64 = ""
    if found_media:
        with open(found_media, "rb") as f:
            mime = "image/gif" if found_media.endswith(".gif") else "image/png"
            media_b64 = f"data:{mime};base64,{base64.b64encode(f.read()).decode()}"

    ui_placeholder = st.empty()

    def render_assistant_ui(status="idle", message="선배님, 루팅 던져주시면 바로 랜드사 규격 엑셀 뽑아오겠습니다! ☕"):
        if status == "working":
            badge_text = "🔥 폭풍 작업 중"
            badge_class = "badge-working"
        elif status == "success":
            badge_text = "🎉 제작 완료!"
            badge_class = "badge-success"
        else:
            badge_text = "☕ 대기 중"
            badge_class = "badge-idle"

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

    render_assistant_ui("idle")

    generate_btn = st.button("🚀 엑셀 일정표 생성하기", type="primary", use_container_width=True)

    if generate_btn:
        if not api_key:
            st.error("설정(Secrets)에서 GEMINI_API_KEY가 등록되지 않았습니다.")
        elif not routing_text.strip():
            st.warning("루팅 정보를 입력해 주세요.")
        else:
            render_assistant_ui("working", "선배님 접수했습니다! 안경 딱 고쳐 쓰고 바로 엑셀 뽑아옵니다!! 두다다다⌨️🔥")
            
            with st.spinner("김똘망 사원이 정밀 규격으로 고속 제작 중..."):
                try:
                    prompt = f"""
당신은 15년 차 유럽 전문 랜드사의 수석 오퍼레이터입니다.
아래 루팅을 분석하여 한국 대형 여행사용 정밀 일정표 데이터를 JSON 배열로 출력하세요.

[입력된 루팅]
{routing_text}

[세부 정제 및 절대 규칙]
1. 도시명(city) 표기 규칙 (매우 중요):
   - 같은 도시의 일정이 계속 이어지는 동안에는 city 칸을 무조건 빈칸("")으로 둡니다.
   - 도시가 새롭게 바뀔 때만 그 행의 city 칸에 최초 1회만 적으세요.
2. 시간란(time) 표기 규칙 (매우 엄격):
   - '시간' 칸에는 오직 항공편 출발/도착 시간(HH:MM)과 관광 시작 행의 '전  일'만 기재합니다.
   - 페리 탑승, 열차, 버스 이동, 관광지 등에는 일체 시간을 적지 말고 무조건 빈칸("")으로 비워두세요.
   - '오전', '오후', '출발', '도착' 등의 한글 텍스트는 시간란에 절대 금지입니다.
3. 교통편('전용차량') 규칙:
   - 일반 육로 관광일의 경우 2행(조식 후 첫 관광/이동 시작 행)에만 transport="전용차량", time="전  일"을 단 1회 표기합니다.
   - 이후의 모든 행(관광지, 식사, 중간이동, 공항이동 등)의 transport는 무조건 빈칸("")입니다.
4. 일정 문구 작성 규칙:
   - '핵심관광', '집중투어' 같은 수식어 절대 금지. 오직 '[도시명] 시내 관광' 형태로 작성하세요.
   - 필수 명소를 포함하여 해당 도시의 대표 명소(▶)를 최소 4개 이상 풍성하게 구성하세요.
   - 관광지는 반드시 '~한 [명소명]' 형태의 개별 행으로 분리 작성하세요.
5. 식사(meals) 규칙:
   - 기본적으로 "호텔식", "현지식"으로 통일하고 기내식 또는 미제공만 "-"로 처리합니다.
6. 숙박 표기:
   - 호텔: "HOTEL : ANY 4* HOTEL CLASS" (마지막 행="석식 후", 그 다음 행="호텔 투숙 및 휴식")
   - 크루즈/페리: "CRUISE : DFDS SEAWAYS (SEAVIEW 2-BERTH)" 등
   - 기내박: "HOTEL : 기내 숙박"
   - 귀국일: hotel="" (마지막 행="인천 도착 후 해산")

[출력 포맷: JSON 배열만 출력]
[
  {{
    "day": 7,
    "rows": [
      {{"city": "", "transport": "", "time": "", "schedule": "호텔 조식 후"}},
      {{"city": "스톡홀름", "transport": "전용차량", "time": "전  일", "schedule": "스톡홀름 시내 관광"}},
      {{"city": "", "transport": "", "time": "", "schedule": " ▶ 호화로운 스웨덴 왕실의 별궁 드로트닝홀름 궁전(내부)"}},
      {{"city": "", "transport": "", "time": "", "schedule": " ▶ 현존하는 가장 오래된 전함 바사호가 전시된 바사 박물관(내부)"}},
      {{"city": "", "transport": "", "time": "", "schedule": " ▶ 중세의 정취가 그대로 보존된 매력적인 감라스탄 구시가지"}},
      {{"city": "", "transport": "", "time": "", "schedule": " ▶ 웅장한 바로크 양식의 스톡홀름 왕궁(외관)"}},
      {{"city": "", "transport": "", "time": "", "schedule": "석식 후"}},
      {{"city": "", "transport": "", "time": "", "schedule": "호텔 투숙 및 휴식"}}
    ],
    "meals": ["호텔식", "현지식", "현지식"],
    "hotel": "HOTEL : ANY 4* HOTEL CLASS"
  }}
]
"""
                    response = client.models.generate_content(
                        model='gemini-3.5-flash-lite',
                        contents=prompt,
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
                    
                    render_assistant_ui("success", "선배님! 요청하신 유럽 일정표 사내 표준 엑셀로 완성했습니다! 아래 다운로드 버튼을 눌러주세요! 🎉")
                    
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
                    render_assistant_ui("idle", f"헉 선배님 오류가 났습니다 ㅠㅠ: {str(e)}")
                    st.error(f"오류 상세: {str(e)}")
