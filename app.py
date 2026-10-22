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
    tour_title = st.text_input("행사명", placeholder="예: 북유럽 4국 10일 SK")
    
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
            with st.spinner("AI가 랜드사 정밀 규격에 맞춰 일정표를 재구성 중입니다..."):
                try:
                    prompt = f"""
당신은 15년 차 유럽 전문 랜드사의 수석 오퍼레이터입니다.
아래 루팅을 분석하여 한국 대형 여행사용 정밀 일정표 데이터를 JSON 배열로 출력하세요.

[입력된 루팅]
{routing_text}

[작성 및 출력 절대 원칙]
1. 각 일차(Day)는 rows 리스트를 가집니다. 각 row는 {{"city": "...", "transport": "...", "time": "...", "schedule": "..."}} 형태입니다.

2. 관광지 구성 및 문체 규칙:
   - 사용자가 '필수'라고 명시한 스팟은 절대 빠뜨리지 말고 반드시 포함하세요.
   - 단, '필수' 명소만 적지 말고, 해당 도시에서 한국 패키지 관광객이 둘러보는 대표 관광 명소들을 3~5개 함께 채워 풍성하게 구성하세요.
   - 주요 관광지 앞에는 반드시 ' ▶ ' 기호를 붙이고, 절대 '입장지: 설명' 형식을 쓰지 마세요.
   - 반드시 '~한 [명소명]' 형태로 작성하세요. (예: " ▶ 알록달록한 건물들이 늘어선 동화 같은 항구 뉘하운 거리", " ▶ 뭉크의 절규를 감상할 수 있는 뭉크 미술관(내부)")
   - '오전', '오후', '저녁'과 같은 시간대 수식어는 일정 텍스트에 넣지 마세요.

3. 교통편 및 이동 규칙 (매우 중요):
   - 특별한 이동 수단이 없는 일반 육로 관광일:
     * 1행(호텔 조식 후)에는 교통편과 시간을 비워둡니다.
     * 조식 후 첫 이동/출발이 시작되는 행(2행)에만 교통편: "전용차량", 시간: "전  일"을 단 한 번 기재합니다.
     * 하루 일정의 중간이나 마지막에는 절대 "전용차량"을 다시 넣지 마세요.
   - 항공 / 야간 페리(크루즈) / 산악열차:
     * 이동편이 시작되는 라인에 편명/교통수단명과 시간을 적습니다.
     * 출발 시간만 적혀 있는 경우, 도착 라인에 통상적인 운항/이동 시간을 계산하여 예상 도착 시간을 적어주세요.

4. 경유 항공편(환승 귀국편) 분리 및 도시명 전환 규칙:
   - 예: 헬싱키/코펜하겐/인천 (SK1717 / SK987)으로 귀국하는 경우
     * 헬싱키 출발 라인: city="헬싱키", transport="SK1717", time="출발시간", schedule="헬싱키 반타 국제공항 출발, 코펜하겐 향발"
     * 코펜하겐 도착 라인: city="코펜하겐", transport="", time="도착시간", schedule="코펜하겐 공항 도착 후 환승 대기"
     * 코펜하겐 재출발 라인: city="", transport="SK987", time="출발시간", schedule="코펜하겐 출발, 인천 향발 (기내박)"
     * 다음 날 인천 도착 라인: city="인  천", transport="", time="도착시간", schedule="인천 국제공항 도착 후 해산"
   - 도시가 바뀔 때는 반드시 해당 라인의 city 칸에 변경된 도시명을 적어주세요.

5. 도시 기재 규칙: 해당 도시 일정이 실제로 시작하는 라인의 도시 칸에만 적으세요.
6. 식사: meals는 [조식, 중식, 석식] 3개 요소 배열. 선박내 뷔페식은 '선상식', 기내박은 '기내식'으로 기재.
7. 호텔: hotel은 "HOTEL : ANY 4* HOTEL CLASS" 또는 "HOTEL : DFDS SEAWAYS (씨뷰 2인 1실)" 또는 "HOTEL : 기내 숙박" 등으로 기재.

[반드시 아래 JSON 포맷으로만 응답]
[
  {{
    "day": 1,
    "rows": [
      {{"city": "인  천", "transport": "", "time": "", "schedule": "인천 국제공항 미팅 및 출국 수속"}},
      {{"city": "", "transport": "SK988", "time": "23:35", "schedule": "인천 국제공항 출발, 코펜하겐 향발"}},
      {{"city": "코펜하겐", "transport": "", "time": "06:30", "schedule": "코펜하겐 카스트룹 공항 도착"}}
    ],
    "meals": ["-", "-", "기내식"],
    "hotel": "HOTEL : 기내 숙박"
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
                    
                    # 엑셀 생성
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
                    
                    # 1행 타이틀
                    full_title = tour_title.strip() if tour_title.strip() else f"{start_date.strftime('%Y-%m-%d')} 유럽투어 일정표"
                    ws.row_dimensions[1].height = 28.0
                    ws.cell(1, 1, full_title).font = font_title
                    ws.cell(1, 1).alignment = align_center
                    
                    for c in range(1, 8):
                        cell = ws.cell(1, c)
                        cell.fill = fill_title
                        cell.border = Border(top=med_side, bottom=thin_side, left=med_side if c==1 else None, right=med_side if c==7 else None)
                    ws.merge_cells("A1:G1")
                    
                    # 2행 헤더
                    headers = ["날짜", "도시", "교통편", "시간", "일정", "식사", ""]
                    ws.row_dimensions[2].height = 22.0
                    for c in range(1, 8):
                        cell = ws.cell(2, c, headers[c-1])
                        cell.font = font_header
                        cell.alignment = align_center
                        cell.border = Border(top=thin_side, bottom=med_side, left=med_side if c==1 else thin_side, right=med_side if c==7 else thin_side)
                    ws.merge_cells("F2:G2")
                    
                    col_widths = {1: 9.25, 2: 12.75, 3: 7.5, 4: 9.25, 5: 64.25, 6: 3.63, 7: 10.0}
                    for c_idx, width in col_widths.items():
                        ws.column_dimensions[get_column_letter(c_idx)].width = width
                    
                    days_korean = ["월요일", "화요일", "수요일", "목요일", "금요일", "토요일", "일요일"]
                    cur_row = 3
                    
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
                        
                        # 세부 일정
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
                    
                    st.success("✅ 정밀 규칙이 반영된 일정표가 생성되었습니다!")
                    final_title = tour_title.strip() if tour_title.strip() else "북유럽일정표"
                    file_name = f"{start_date.strftime('%Y-%m-%d')} {final_title} - 일정표.xlsx"
                    
                    st.download_button(
                        label="📥 완성된 엑셀 일정표 다운로드",
                        data=output,
                        file_name=file_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"오류가 발생했습니다: {str(e)}")
