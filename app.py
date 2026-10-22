import streamlit as st
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta
import io
import json
from google import genai
from google.genai import types

# 페이지 기본 설정
st.set_page_config(page_title="유럽 랜드사 일정표 생성기", layout="wide")
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
    st.subheader("2. 일정표 생성")
    generate_btn = st.button("🚀 엑셀 일정표 생성하기", type="primary")

    if generate_btn:
        if not api_key:
            st.error("설정(Secrets)에서 GEMINI_API_KEY가 등록되지 않았습니다.")
        elif not routing_text.strip():
            st.warning("루팅 정보를 입력해 주세요.")
        else:
            with st.spinner("AI가 랜드사 정밀 규격에 맞춰 풍성한 일정표를 생성 중입니다..."):
                try:
                    prompt = f"""
당신은 15년 차 유럽 전문 랜드사의 수석 오퍼레이터입니다.
아래 루팅을 분석하여 한국 대형 여행사용 정밀 일정표 데이터를 JSON 배열로 출력하세요.

[입력된 루팅]
{routing_text}

[관광지 구성 절대 원칙 - 엄격 준수]
1. 관광지 수량 할당제 (중요):
   - 관광이 포함된 날은 ' ▶ '로 시작하는 관광지 행이 **반드시 최소 4개 이상** 포함되어야 합니다.
   - 사용자가 '필수'라고 명시한 곳은 1순위로 반드시 넣되, '필수' 스팟만 달랑 넣는 것은 엄격히 금지됩니다.
   - 반드시 해당 도시의 한국인 패키지 대표 명소(광장, 대성당, 구시가지, 궁전 외관 등)를 최소 3~4개 추가하여 하루 전체 일정을 풍성하게 채우세요.
   - 예시: 오슬로(시청사 내부 필수) ➔ 시청사(내부) 뿐만 아니라 '비겔란 조각공원', '아케르스후스 요새', '칼 요한 거리' 등을 반드시 함께 포함하여 최소 4개 이상의 명소를 구성할 것.
2. 관광지 문체 및 행 분리:
   - 1행당 1개 관광지만 작성 (한 행에 몰아넣기 금지).
   - 반드시 '~한 [명소명]' 문체 적용 (예: " ▶ 바이킹의 역사를 담은 오슬로 시청사(내부)", " ▶ 인간의 희로애락을 조각한 비겔란 조각공원").
3. 시간란 규칙:
   - '오전', '오후', '출발', '도착' 일체 금지.
   - 오직 'HH:MM' 형태의 숫자 시간(예: "23:35", "06:30") 또는 '전  일'만 표기.
4. 교통편('전용차량') 규칙:
   - 일반 관광일: 2행에만 transport="전용차량", time="전  일" 단 1회 표기.
   - 이후 관광지, 식사, 호텔, "공항으로 이동" 등의 행에는 transport 칸을 무조건 빈칸("")으로 유지.
5. 항공편 및 경유지 분리:
   - 도시 이름은 무조건 city 칸에만 기재 (transport 칸에 도시명 절대 금지).
   - transport는 순수 항공 편명(예: "SK1717")만 기재.
   - 출발 시간과 도착 시간을 각각의 행에 모두 기재.
6. 숙박 표기:
   - 호텔 투숙일: hotel="HOTEL : ANY 4* HOTEL CLASS" (마지막 행은 "호텔 투숙 및 휴식")
   - 페리/크루즈일: hotel="CRUISE : DFDS SEAWAYS (SEAVIEW 2-BERTH)" 등 (마지막 행은 "선내 투숙 및 휴식")
   - 기내 숙박일: hotel="HOTEL : 기내 숙박" (별도 투숙 문구 없음)
   - 귀국 당일: hotel="" (마지막 행은 "인천 도착 후 해산")

[반드시 아래 예시 포맷의 풍성한 JSON 구조로 응답]
[
  {{
    "day": 3,
    "rows": [
      {{"city": "", "transport": "", "time": "", "schedule": "호텔 조식 후"}},
      {{"city": "오슬로", "transport": "전용차량", "time": "전  일", "schedule": "오슬로 시내 관광"}},
      {{"city": "", "transport": "", "time": "", "schedule": " ▶ 노벨 평화상 시상식이 열리는 상징적인 오슬로 시청사(내부)"}},
      {{"city": "", "transport": "", "time": "", "schedule": " ▶ 인간의 삶을 생생하게 표현한 비겔란 조각공원"}},
      {{"city": "", "transport": "", "time": "", "schedule": " ▶ 오슬로 최고의 번화가이자 활기찬 칼 요한 거리"}},
      {{"city": "", "transport": "", "time": "", "schedule": " ▶ 바다를 굽어보는 중세의 웅장한 아케르스후스 요새(조망)"}},
      {{"city": "", "transport": "", "time": "", "schedule": "게일로로 이동"}},
      {{"city": "", "transport": "", "time": "", "schedule": "석식 후"}},
      {{"city": "", "transport": "", "time": "", "schedule": "호텔 투숙 및 휴식"}}
    ],
    "meals": ["호텔식", "현지식", "호텔식"],
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
                    
                    parsed_days = json.loads(response.text)
                    
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
                    
                    # 1행: 타이틀바
                    full_title = tour_title.strip() if tour_title.strip() else f"{start_date.strftime('%Y-%m-%d')} 유럽투어 일정표"
                    ws.row_dimensions[1].height = 28.0
                    ws.cell(1, 1, full_title).font = font_title
                    ws.cell(1, 1).alignment = align_center
                    
                    for c in range(1, 8):
                        cell = ws.cell(1, c)
                        cell.fill = fill_title
                        cell.border = Border(top=med_side, bottom=med_side, left=med_side if c==1 else None, right=med_side if c==7 else None)
                    ws.merge_cells("A1:G1")
                    
                    # 2행: 간격행 (높이 13.0)
                    ws.row_dimensions[2].height = 13.0
                    
                    # 3행: 헤더행
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
                        
                        # 2~5열 세부 일정
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
                            
                        # HOTEL 행
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
                    
                    st.success("✅ 사내 표준 인쇄 서식으로 일정표가 완성되었습니다!")
                    final_title = tour_title.strip() if tour_title.strip() else "유럽투어"
                    file_name = f"{start_date.strftime('%Y-%m-%d')} {final_title} - 일정표.xlsx"
                    
                    st.download_button(
                        label="📥 완성된 엑셀 일정표 다운로드",
                        data=output,
                        file_name=file_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"오류가 발생했습니다: {str(e)}")
