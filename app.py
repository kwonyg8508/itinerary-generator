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
            with st.spinner("AI가 랜드사 정밀 규격에 맞춰 일정표를 생성 중입니다..."):
                try:
                    prompt = f"""
당신은 15년 차 유럽 전문 랜드사의 수석 오퍼레이터입니다.
아래 루팅을 분석하여 한국 대형 여행사용 정밀 일정표 데이터를 JSON 배열로 출력하세요.

[입력된 루팅]
{routing_text}

[작성 및 출력 절대 원칙]
1. 각 일차(Day)는 rows 리스트를 가집니다.
   - 절대 한 schedule 칸 안에 줄바꿈(\\n)으로 여러 관광지를 몰아넣지 마세요.
   - 관광지 하나당 무조건 rows의 개별 객체(행) 하나로 분리하세요.

2. 관광지 표기 및 문체 규칙:
   - 사용자가 '필수'라고 명시한 스팟은 절대 누락 없이 포함하세요.
   - 필수 스팟 외에도 해당 도시의 대표 명소를 3~5개 함께 추가하여 일정을 풍성하게 구성하세요.
   - 관광지 앞에는 반드시 ' ▶ ' 기호를 붙이고, 절대 '입장지: 설명' 형식을 쓰지 마세요.
   - 반드시 '~한 [명소명]' 형태로 작성하세요. (예: " ▶ 알록달록한 건물들이 늘어선 동화 같은 항구 뉘하운 거리")
   - '오전', '오후', '저녁'과 같은 시간대 수식어는 일정 텍스트에 넣지 마세요.

3. 교통편('전용차량') 표기 규칙 (매우 엄격):
   - 일반 육로 관광일:
     * 1행: [일정: "호텔 조식 후"] (교통편/시간 비움)
     * 2행: 첫 이동 또는 첫 관광지 라인에만 [교통편: "전용차량", 시간: "전  일"]을 단 1회 기재.
     * 3행 이후의 모든 관광지, 석식, 호텔 투숙 행의 교통편/시간 칸은 무조건 빈칸("")이어야 합니다.
   - 항공, 크루즈(페리), 산악열차 등 특수 이동이 있는 경우에만 해당 행에 편명과 시간을 기재하세요.

4. 하루 일정의 마지막 라인(일정 마무리) 규칙:
   - 일반 호텔 투숙일: 마지막 전 행은 "석식 후", 마지막 행은 "호텔 투숙 및 휴식"
   - 야간 크루즈/페리 탑승일: 마지막 전 행은 "선내 뷔페 석식 후", 마지막 행은 "선내 투숙 및 휴식"
   - 기내박(야간 비행) 또는 귀국일: 해당 항공 스케줄 및 "인천 도착 후 해산"으로 마무리 (투숙 문구 불필요)

5. 호텔 및 숙박 표기 (영문 기준):
   - 일반 4성 호텔: "HOTEL : ANY 4* HOTEL CLASS"
   - 크루즈/페리: "CRUISE : DFDS SEAWAYS (SEAVIEW 2-BERTH)" 또는 "CRUISE : SILJA LINE (SEAVIEW 2-BERTH)"
   - 기내 숙박: "HOTEL : IN-FLIGHT"

6. 경유 항공편 분리:
   - 도시가 바뀔 때마다(환승지, 도착지) rows를 분리하고 city 칸에 변경된 도시명을 적으세요.
   - 출발 시간만 있을 경우, 통상적인 비행시간을 반영하여 도착 시간을 쌍으로 적으세요.

7. 식사(meals): [조식, 중식, 석식] 3개 요소 배열 (예: ["호텔식", "현지식", "선상식"], 비행기/미제공은 "-")

[반드시 아래 JSON 포맷으로만 응답]
[
  {{
    "day": 1,
    "rows": [
      {{"city": "인  천", "transport": "", "time": "", "schedule": "인천 국제공항 미팅 및 출국 수속"}},
      {{"city": "", "transport": "SK988", "time": "23:35", "schedule": "인천 국제공항 출발, 코펜하겐 향발"}}
    ],
    "meals": ["-", "-", "기내식"],
    "hotel": "HOTEL : IN-FLIGHT"
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
                    
                    # 1행: 최상단 타이틀 (A1:G1)
                    full_title = tour_title.strip() if tour_title.strip() else f"{start_date.strftime('%Y-%m-%d')} 유럽투어 일정표"
                    ws.row_dimensions[1].height = 28.0
                    ws.cell(1, 1, full_title).font = font_title
                    ws.cell(1, 1).alignment = align_center
                    
                    for c in range(1, 8):
                        cell = ws.cell(1, c)
                        cell.fill = fill_title
                        cell.border = Border(top=med_side, bottom=med_side, left=med_side if c==1 else None, right=med_side if c==7 else None)
                    ws.merge_cells("A1:G1")
                    
                    # 2행: 요청하신 높이 13짜리 빈 구분행 (테두리 없음)
                    ws.row_dimensions[2].height = 13.0
                    
                    # 3행: 헤더행 (A3:G3)
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
                    cur_row = 4  # 본문 데이터는 4행부터 시작
                    
                    for item in parsed_days:
                        day_num = item.get("day", 1)
                        curr_date = start_date + timedelta(days=day_num - 1)
                        day_str = f"제{day_num:02d}일"
                        date_str = curr_date.strftime("%m월 %d일")
                        weekday_str = days_korean[curr_date.weekday()]
                        
                        rows_data = item.get("rows", [])
                        hotel_str = item.get("hotel", "HOTEL : ANY 4* HOTEL CLASS")
                        meals = item.get("meals", ["-", "-", "-"])
                        
                        total_rows = max(len(rows_data), 3) + 1
                        
                        # 셀 서식 및 테두리 설정 (내부 가로선 없음)
                        for r_idx in range(total_rows):
                            r = cur_row + r_idx
                            ws.row_dimensions[r].height = 18.75
                            is_hotel_row = (r_idx == total_rows - 1)
                            
                            for c in range(1, 8):
                                cell = ws.cell(r, c)
                                cell.font = font_body
                                
                                b_top = thin_side if is_hotel_row else None
                                b_bottom = med_side if is_hotel_row else None
                                b_left = med_side if c == 1 else thin_side
                                b_right = med_side if c == 7 else thin_side
                                
                                if is_hotel_row:
                                    cell.fill = fill_hotel
                                cell.border = Border(top=b_top, bottom=b_bottom, left=b_left, right=b_right)
                        
                        # 1열 날짜
                        ws.cell(cur_row, 1, day_str).font = font_bold
                        ws.cell(cur_row, 1).alignment = align_center
                        ws.cell(cur_row + 1, 1, date_str).alignment = align_center
                        ws.cell(cur_row + 2, 1, weekday_str).alignment = align_center
                        
                        # 2~5열 세부 일정 기입 (각 행별로 깔끔하게 분리)
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
