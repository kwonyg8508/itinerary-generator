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
    tour_title = st.text_input("행사명", placeholder="예: 2026-12-05 SM투어) 동유럽 3국 9일 KE")
    
    routing_text = st.text_area(
        "루팅 및 일정 입력 (복사-붙여넣기)",
        height=320,
        placeholder="""1. ICN/VIE (KE)
2. 비엔나-잘츠부르크(호엔성 및 트래킹)
3. 잘츠부르크-할슈타트 or 잘츠카머굿/트래킹 및 파이브핑거스 전망대-블레드
4. 블레드 (블레드성, 호수, 와이너리)
5. 트리글라브 국립공원 (점심 피크닉 느낌 도시락) - 류블랴냐
6. 류블랴냐(관광) - 비엔나
7. 비엔나(관광) 후 OUT
8. 인천"""
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
            with st.spinner("AI가 랜드사 표준 문체와 서식에 맞춰 일정표를 제작 중입니다..."):
                try:
                    # 1. 랜드사 전용 프롬프트 구성
                    prompt = f"""
당신은 15년 차 유럽 전문 랜드사의 수석 오퍼레이터입니다.
아래 루팅을 분석하여 한국 대형 여행사용 정밀 일정표 데이터를 JSON 배열로 출력하세요.

[입력된 루팅]
{routing_text}

[작성 및 출력 원칙]
1. 각 일차(Day)는 row_items 리스트를 가집니다. 각 row는 [도시, 교통편, 시간, 일정내용]으로 구성됩니다.
2. 일정 설명 문체:
   - 주요 관광지 앞에는 반드시 ' ▶ ' 기호를 붙이고, 절대 '입장지: 설명' 형식을 쓰지 마세요.
   - 반드시 '~한 [명소명]' 형태로 작성하세요. (예: " ▶ 아름다운 프라하의 휴식처 카를교", " ▶ 독특한 철제 간판이 늘어선 번화가 게트라이데 거리", " ▶ 깍아지는 절벽위에 세워진 고성 블레드성")
   - 내부관람 시에는 명소 뒤에 '(내부)' 표기.
3. 일정 라인 구성 순서:
   - 1일차(출발일):
     1행: [도시: "인  천", 교통편: "", 시간: "", 일정: "인천 국제공항 미팅 및 출국 수속"]
     2행: [도시: "", 교통편: 항공편명(예: KE969), 시간: 출발시간(예: "10:50"), 일정: "인천 국제공항 출발, [목적지] 향발"]
     3행: [도시: 목적지도착도시, 교통편: "", 시간: 도착시간(예: "16:20"), 일정: "[목적지] 도착 후 가이드 미팅"]
     4행: [도시: "", 교통편: "", 시간: "", 일정: "석식 후"]
     5행: [도시: "", 교통편: "", 시간: "", 일정: "호텔 투숙 및 휴식"]
   - 2일차 이후(관광일):
     1행: [도시: "", 교통편: "", 시간: "", 일정: "호텔 조식 후"]
     2행: [도시: "", 교통편: "", 시간: "", 일정: "[이동지 또는 첫 집결 안내]"] (조식 후 이동이 있으면 조식 아랫 칸에 기재)
     3행: [도시: 관광도시명, 교통편: "전용차량", 시간: "전  일", 일정: " ▶ 첫번째 관광지..."]
     이후 행: 관광지들을 ' ▶ ~한 명소' 형태로 순차 기재
     마지막 전 행: [도시: "", 교통편: "", 시간: "", 일정: "석식 후"]
     마지막 행: [도시: "", 교통편: "", 시간: "", 일정: "호텔 투숙 및 휴식"]
4. 도시 기재 규칙: 해당 도시 일정이 실제로 시작하는 라인의 도시 칸에만 적으세요.
5. 식사: meals는 [조식, 중식, 석식] 3개 요소 배열 (예: ["호텔식", "현지식", "한  식"], 비행기 이동은 "-", 없으면 "-"). 특식 표시는 제외.
6. 호텔: hotel은 "HOTEL : ANY 4* HOTEL CLASS" 또는 기내숙박 시 "HOTEL : 기내 숙박".

[반드시 아래 JSON 포맷으로만 응답]
[
  {{
    "day": 1,
    "rows": [
      {{"city": "인  천", "transport": "", "time": "", "schedule": "인천 국제공항 제2터미널 미팅"}},
      {{"city": "", "transport": "KE969", "time": "10:50", "schedule": "인천 국제공항 출발, 프라하 향발"}},
      {{"city": "프라하", "transport": "", "time": "16:20", "schedule": "프라하 도착"}},
      {{"city": "", "transport": "", "time": "", "schedule": "석식 후"}},
      {{"city": "", "transport": "", "time": "", "schedule": "호텔 투숙 및 휴식"}}
    ],
    "meals": ["-", "-", "현지식"],
    "hotel": "HOTEL : ANY 4* HOTEL CLASS"
  }}
]
"""
                    response = client.models.generate_content(
                        model='gemini-3.8-flash',
                        contents=prompt,
                        config=types.GenerateContentConfig(
                            response_mime_type="application/json"
                        )
                    )
                    
                    parsed_days = json.loads(response.text)
                    
                    # 2. 엑셀 워크북 직접 생성 (사내 정밀 규격 적용)
                    wb = openpyxl.Workbook()
                    ws = wb.active
                    ws.title = "일정표"
                    ws.views.sheetView[0].showGridLines = True
                    
                    # 스타일 정의
                    font_title = Font(name="맑은 고딕", size=15, bold=True)
                    font_header = Font(name="맑은 고딕", size=10, bold=True)
                    font_body = Font(name="맑은 고딕", size=9)
                    font_bold = Font(name="맑은 고딕", size=9, bold=True)
                    
                    fill_title = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")
                    fill_hotel = PatternFill(start_color="EFEFEF", end_color="EFEFEF", fill_type="solid")
                    
                    align_center = Alignment(horizontal="center", vertical="center")
                    align_top_center = Alignment(horizontal="center", vertical="top", wrap_text=True)
                    align_top_left = Alignment(horizontal="left", vertical="top", wrap_text=True)
                    
                    thin_side = Side(style="thin", color="000000")
                    med_side = Side(style="medium", color="000000")
                    
                    # 1행: 최상단 제목 (A1:G1 병합 및 노란 배경)
                    full_title = tour_title.strip() if tour_title.strip() else f"{start_date.strftime('%Y-%m-%d')} 유럽투어 일정표"
                    ws.merge_cells("A1:G1")
                    ws.row_dimensions[1].height = 28.0
                    for c in range(1, 8):
                        cell = ws.cell(1, c)
                        cell.fill = fill_title
                        cell.border = Border(top=med_side, bottom=thin_side, left=med_side if c==1 else None, right=med_side if c==7 else None)
                    t_cell = ws.cell(1, 1, full_title)
                    t_cell.font = font_title
                    t_cell.alignment = align_center
                    
                    # 2행: 헤더행
                    headers = ["날짜", "도시", "교통편", "시간", "일정", "식사", ""]
                    ws.row_dimensions[2].height = 22.0
                    ws.merge_cells("F2:G2")
                    for c in range(1, 8):
                        cell = ws.cell(2, c, headers[c-1])
                        cell.font = font_header
                        cell.alignment = align_center
                        cell.border = Border(top=thin_side, bottom=med_side, left=med_side if c==1 else thin_side, right=med_side if c==7 else thin_side)
                    
                    # 열 너비 설정 (요청 규격 100% 반영)
                    col_widths = {1: 9.25, 2: 12.75, 3: 7.5, 4: 9.25, 5: 64.25, 6: 3.63, 7: 10.0}
                    for c_idx, width in col_widths.items():
                        ws.column_dimensions[get_column_letter(c_idx)].width = width
                    
                    days_korean = ["월요일", "화요일", "수요일", "목요일", "금요일", "토요일", "일요일"]
                    cur_row = 3
                    
                    # 일자별 데이터 작성 루프
                    for item in parsed_days:
                        day_num = item.get("day", 1)
                        curr_date = start_date + timedelta(days=day_num - 1)
                        day_str = f"제{day_num:02d}일"
                        date_str = curr_date.strftime("%m월 %d일")
                        weekday_str = days_korean[curr_date.weekday()]
                        
                        rows_data = item.get("rows", [])
                        hotel_str = item.get("hotel", "HOTEL : ANY 4* HOTEL CLASS")
                        meals = item.get("meals", ["-", "-", "-"])
                        
                        # 최소 3줄(날짜/식사용) 및 행 계산
                        total_rows = max(len(rows_data), 3) + 1 # +1은 HOTEL 행
                        
                        # 일자 블록 기본 셀 서식 (가로 내부선 제거, 세로선만 적용)
                        for r_idx in range(total_rows):
                            r = cur_row + r_idx
                            ws.row_dimensions[r].height = 18.75
                            is_hotel_row = (r_idx == total_rows - 1)
                            
                            for c in range(1, 8):
                                cell = ws.cell(r, c)
                                cell.font = font_body
                                
                                # 테두리 처리: 내부 가로선 없음, 블록 맨 아래(HOTEL)에만 굵은 바닥선
                                b_top = None
                                b_bottom = med_side if is_hotel_row else None
                                b_left = med_side if c == 1 else thin_side
                                b_right = med_side if c == 7 else (thin_side if c != 6 else thin_side)
                                
                                # HOTEL 행 바로 위에는 얇은 윗선 추가
                                if is_hotel_row:
                                    b_top = thin_side
                                    cell.fill = fill_hotel
                                
                                cell.border = Border(top=b_top, bottom=b_bottom, left=b_left, right=b_right)
                        
                        # 1열: 날짜 (제01일 / 12월 05일 / 요일)
                        c1 = ws.cell(cur_row, 1, day_str)
                        c1.font = font_bold
                        c1.alignment = align_top_center
                        
                        c2 = ws.cell(cur_row + 1, 1, date_str)
                        c2.alignment = align_top_center
                        
                        c3 = ws.cell(cur_row + 2, 1, weekday_str)
                        c3.alignment = align_top_center
                        
                        # 2~5열: 세부 행 데이터 기입
                        for idx, r_data in enumerate(rows_data):
                            r = cur_row + idx
                            # 도시
                            if r_data.get("city"):
                                cell_city = ws.cell(r, 2, r_data.get("city"))
                                cell_city.alignment = align_top_center
                            # 교통편
                            if r_data.get("transport"):
                                cell_tr = ws.cell(r, 3, r_data.get("transport"))
                                cell_tr.alignment = align_top_center
                            # 시간
                            if r_data.get("time"):
                                cell_tm = ws.cell(r, 4, r_data.get("time"))
                                cell_tm.alignment = align_top_center
                            # 일정
                            cell_sc = ws.cell(r, 5, r_data.get("schedule", ""))
                            cell_sc.alignment = align_top_left
                            
                        # HOTEL 행 기입 (일정 5열에 볼드 표기)
                        h_row = cur_row + total_rows - 1
                        h_cell = ws.cell(h_row, 5, hotel_str)
                        h_cell.font = font_bold
                        h_cell.alignment = align_top_left
                        
                        # 6, 7열: 조/중/석 식사 기입 (내부 가로선 없이 3행에 깔끔 배치)
                        meal_labels = ["조", "중", "석"]
                        for m_idx in range(3):
                            l_cell = ws.cell(cur_row + m_idx, 6, meal_labels[m_idx])
                            l_cell.alignment = align_top_center
                            
                            v_cell = ws.cell(cur_row + m_idx, 7, meals[m_idx] if m_idx < len(meals) else "-")
                            v_cell.alignment = align_top_center
                            
                        cur_row += total_rows
                        
                    output = io.BytesIO()
                    wb.save(output)
                    output.seek(0)
                    
                    st.success("✅ 사내 표준 인쇄 서식으로 일정표가 생성되었습니다!")
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
