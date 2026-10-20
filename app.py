import streamlit as st
import openpyxl
from copy import copy
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
    tour_title = st.text_input("행사명", placeholder="예: 그리스 산토리니/미코노스 9일")
    
    routing_text = st.text_area(
        "루팅 및 일정 입력 (복사-붙여넣기)",
        height=320,
        placeholder="""1. ICN/DXB
2. DXB/ATH (Arr. 1500) - 간단시내관광(1시간) - 석식 - 호텔투숙
3. ATH 풀데이 관광 (포세이돈 신전포함) - 루프탑 석식
4. ATH - 페리탑승 - 미코노스(관광)
5. 미코노스 - 페리탑승 - 낙소스(관광)
6. 낙소스(관광) - 오후 17시 페리탑승 - 산토리니
7. 산토리니(풀데이, 석식 특식)
8. 산토리니 (오전 관광 및 중식 후) 공항이동
9. 인천도착"""
    )

with col2:
    st.subheader("2. 일정표 생성")
    generate_btn = st.button("🚀 엑셀 일정표 생성하기", type="primary")

    if generate_btn:
        if not api_key:
            st.error("설정에서 GEMINI_API_KEY가 등록되지 않았습니다.")
        elif not routing_text.strip():
            st.warning("루팅 정보를 입력해 주세요.")
        else:
            with st.spinner("AI가 일정을 분석하고 사내 서식 엑셀을 제작 중입니다..."):
                try:
                    # 1. Gemini로 루팅 텍스트 구조화 파싱
                    prompt = f"""
                    당신은 15년 차 유럽 전문 랜드사 오퍼레이터입니다.
                    아래 고객의 루팅 텍스트를 분석하여, 유럽 랜드사 표준 일정표에 들어갈 일자별 데이터를 JSON 형식으로 변환하세요.

                    [입력된 루팅]
                    {routing_text}

                    [작성 원칙]
                    1. city: 방문 도시명 (예: "인   천", "두   바이\\n아 테 네", "아 테 네\\n수 니 온").
                    2. transport: 항공편명, 페리, 전용차량 등 (예: "EK323", "전용차량", "쾌속페리").
                    3. time: 시간 정보 (예: "23:55", "오  전", "전  일", "17:00").
                    4. schedule: 해당 일자의 세부 일정 리스트 (문자열 배열).
                       - 주요 관광지 앞에는 반드시 ' ▶ ' 기호를 붙이고 핵심 명칭과 짧은 설명을 적으세요.
                       - 특식이나 하이라이트가 있으면 '★ 특식 : ' 형태로 기재하세요.
                       - 하루 일정의 마지막은 '호텔 투숙 및 휴식' 형태로 마무리하세요.
                    5. meals: [조식, 중식, 석식] 3개 요소 배열 (예: ["호텔식", "현지식", "특식"], 비행기 이동은 "기내식", 없으면 "-").
                    6. hotel: "HOTEL : ANY 4* HOTEL CLASS" 또는 기내숙박 시 "HOTEL : 기내 숙박", 크루즈 시 "HOTEL : DFDS 크루즈 - SEA SIDE 기준" 등.
                    
                    반드시 유효한 JSON 배열 형태로만 응답하세요. (마크다운 코드블록 제외, 순수 JSON만)
                    예시:
                    [
                      {{
                        "day": 1,
                        "city": "인   천",
                        "transport": "EK323",
                        "time": "23:55",
                        "schedule": ["인천 국제공항 미팅", "인천 출발"],
                        "meals": ["-", "-", "기내식"],
                        "hotel": "HOTEL : 기내 숙박"
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
                    
                    # 2. 엑셀 템플릿 로드 및 데이터 입력
                    src_wb = openpyxl.load_workbook("template.xlsx")
                    src_ws = src_wb.active
                    
                    wb = openpyxl.Workbook()
                    ws = wb.active
                    ws.title = "일정표"
                    
                    # 서식 및 열 너비 복사
                    for col in range(1, 8):
                        col_letter = openpyxl.utils.get_column_letter(col)
                        ws.column_dimensions[col_letter].width = src_ws.column_dimensions[col_letter].width
                        
                    for r in range(1, 4):
                        for c in range(1, 8):
                            src_cell = src_ws.cell(r, c)
                            tgt_cell = ws.cell(r, c, src_cell.value)
                            if src_cell.has_style:
                                tgt_cell.font = copy(src_cell.font)
                                tgt_cell.border = copy(src_cell.border)
                                tgt_cell.fill = copy(src_cell.fill)
                                tgt_cell.alignment = copy(src_cell.alignment)
                                
                    sample_border = copy(src_ws.cell(4, 1).border)
                    sample_font = copy(src_ws.cell(4, 5).font)
                    sample_font_bold = copy(src_ws.cell(4, 1).font)
                    sample_align_center = copy(src_ws.cell(4, 1).alignment)
                    sample_align_left = copy(src_ws.cell(4, 5).alignment)
                    
                    days_korean = ["월요일", "화요일", "수요일", "목요일", "금요일", "토요일", "일요일"]
                    cur_row = 4
                    
                    for item in parsed_days:
                        day_num = item.get("day", 1)
                        curr_date = start_date + timedelta(days=day_num - 1)
                        day_str = f"제{day_num:02d}일"
                        date_str = curr_date.strftime("%Y-%m-%d 00:00:00")
                        weekday_str = days_korean[curr_date.weekday()]
                        
                        sch_list = item.get("schedule", [])
                        hotel_str = item.get("hotel", "")
                        total_lines = len(sch_list) + (1 if hotel_str else 0)
                        total_rows = max(3, total_lines)
                        
                        for r_idx in range(total_rows):
                            r = cur_row + r_idx
                            for c in range(1, 8):
                                cell = ws.cell(r, c)
                                cell.font = copy(sample_font)
                                cell.border = copy(sample_border)
                                cell.alignment = copy(sample_align_left)
                                
                        # 1열: 날짜
                        c1_1 = ws.cell(cur_row, 1, day_str)
                        c1_1.font = copy(sample_font_bold)
                        c1_1.alignment = copy(sample_align_center)
                        
                        c1_2 = ws.cell(cur_row + 1, 1, date_str)
                        c1_2.alignment = copy(sample_align_center)
                        
                        c1_3 = ws.cell(cur_row + 2, 1, weekday_str)
                        c1_3.alignment = copy(sample_align_center)
                        
                        # 2열: 도시
                        c2 = ws.cell(cur_row, 2, item.get("city", ""))
                        c2.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='top', wrap_text=True)
                        
                        # 3열: 교통편
                        c3 = ws.cell(cur_row, 3, item.get("transport", ""))
                        c3.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='top', wrap_text=True)
                        
                        # 4열: 시간
                        c4 = ws.cell(cur_row, 4, item.get("time", ""))
                        c4.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='top', wrap_text=True)
                        
                        # 5열: 세부일정
                        for s_idx, sch in enumerate(sch_list):
                            ws.cell(cur_row + s_idx, 5, sch)
                        if hotel_str:
                            h_cell = ws.cell(cur_row + total_lines - 1, 5, hotel_str)
                            h_cell.font = copy(sample_font_bold)
                            
                        # 6, 7열: 조/중/석 식사
                        meals = item.get("meals", ["-", "-", "-"])
                        meal_labels = ["조", "중", "석"]
                        for m_idx in range(3):
                            lbl_cell = ws.cell(cur_row + m_idx, 6, meal_labels[m_idx])
                            lbl_cell.alignment = copy(sample_align_center)
                            lbl_cell.font = copy(sample_font_bold)
                            
                            val_cell = ws.cell(cur_row + m_idx, 7, meals[m_idx] if m_idx < len(meals) else "-")
                            val_cell.alignment = copy(sample_align_center)
                            
                        cur_row += total_rows
                        
                    output = io.BytesIO()
                    wb.save(output)
                    output.seek(0)
                    
                    st.success("✅ 일정표 생성이 완료되었습니다!")
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
