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

[절대 금지 및 필수 작성 규칙]
1. 공항 명칭 금지: '국제공항', '터미널', '반타', '카스트룹' 등 공항 명칭 일체 쓰지 마세요. 무조건 '인천', '코펜하겐' 등 도시명만 씁니다.
2. 시간란 규칙: '오전', '오후', '저녁', '출발', '도착' 등의 한글을 절대 넣지 마세요. 오직 'HH:MM' 형태의 순수 숫자 시간이나 '전  일'만 허용됩니다.
3. 교통편('전용차량') 표기 규칙:
   - 일반 육로 관광일의 경우: 1행(호텔 조식 후) 다음 줄인 2행에만 transport="전용차량", time="전  일"을 단 1회 기재합니다.
   - 그 이후 중간 이동('~로 이동')이나 관광지, 식사 행에는 transport를 절대 넣지 말고 무조건 빈칸("")으로 두세요.
4. 항공편 표기 규칙:
   - 편명(transport)은 비행기가 출발하는 행에만 적고, 도착 행의 transport는 빈칸("")으로 둡니다.
   - 출발 시간과 도착 시간은 각각 해당 행의 time 칸에 반드시 둘 다 숫자로 적으세요 (출발 시간만 있으면 예상 도착 시간을 계산해서 채울 것).
   - 비행기 출발 일정 문구는 간결하게 '[목적지] 향발'로 작성하세요 (예: "인천 출발, 코펜하겐 향발").
5. 관광지 행 분리:
   - 한 행에 여러 관광지를 줄바꿈으로 몰아넣지 마세요. 무조건 1개 관광지당 rows의 1개 객체로 분리하세요.
   - 관광지 앞에는 ' ▶ '를 붙이고 '~한 [명소명]' 문체를 유지하세요. 사용자의 필수 스팟을 포함하여 3~5개 명소를 채우세요.
6. 일정 마무리:
   - 호텔 투숙일: 마지막 전 행="석식 후", 마지막 행="호텔 투숙 및 휴식"
   - 크루즈 투숙일: 마지막 전 행="선내 뷔페 석식 후", 마지막 행="선내 투숙 및 휴식"
   - 귀국/기내박일: 별도 투숙 문구 없이 항공 일정 및 "인천 도착 후 해산"으로 마무리.
7. 호텔 표기: "HOTEL : ANY 4* HOTEL CLASS", "CRUISE : DFDS SEAWAYS (SEAVIEW 2-BERTH)", "HOTEL : IN-FLIGHT" 등 영문 표기.

[반드시 아래 예시와 정확히 동일한 구조의 JSON 포맷으로만 응답]
[
  {{
    "day": 1,
    "rows": [
      {{"city": "인  천", "transport": "", "time": "", "schedule": "인천 미팅 및 출국 수속"}},
      {{"city": "", "transport": "SK988", "time": "23:35", "schedule": "인천 출발, 코펜하겐 향발"}}
    ],
    "meals": ["-", "-", "기내식"],
    "hotel": "HOTEL : IN-FLIGHT"
  }},
  {{
    "day": 2,
    "rows": [
      {{"city": "코펜하겐", "transport": "", "time": "06:30", "schedule": "코펜하겐 도착 후 가이드 미팅"}},
      {{"city": "", "transport": "전용차량", "time": "전  일", "schedule": "코펜하겐 시내로 이동"}},
      {{"city": "", "transport": "", "time": "", "schedule": " ▶ 알록달록한 건물이 늘어선 동화 같은 항구 뉘하운 유람선 탑승"}},
      {{"city": "", "transport": "", "time": "", "schedule": " ▶ 덴마크 왕실의 주거지 아말리엔보르 궁전(외관)"}},
      {{"city": "", "transport": "", "time": "", "schedule": "DFDS 터미널로 이동하여 승선 수속"}},
      {{"city": "", "transport": "DFDS", "time": "16:30", "schedule": "코펜하겐 출발, 오슬로 향발"}},
      {{"city": "", "transport": "", "time": "", "schedule": "선내 뷔페 석식 후"}},
      {{"city": "", "transport": "", "time": "", "schedule": "선내 투숙 및 휴식"}}
    ],
    "meals": ["기내식", "현지식", "선상식"],
    "hotel": "CRUISE : DFDS SEAWAYS (SEAVIEW 2-BERTH)"
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
                        hotel_str = item.get("hotel", "HOTEL : ANY 4* HOTEL CLASS")
                        meals = item.get("meals", ["-", "-", "-"])
                        
                        total_rows = max(len(rows_data), 3) + 1
                        
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
                            cell_sc = ws.cell(r,
