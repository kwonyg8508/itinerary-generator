import streamlit as st
import os
import json
import io
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime, timedelta
from google import genai
from google.genai import types

# 항공 검색 애드온 모듈 임포트
try:
    from flight_search import render_flight_search_ui
except ImportError:
    render_flight_search_ui = None

# -------------------------------------------------------------
# 1. 페이지 기본 설정 및 스타일
# -------------------------------------------------------------
st.set_page_config(page_title="유럽 일정표 자동 생성 시스템", layout="wide")

st.markdown("""
<style>
    .main-title { font-size: 1.8rem; font-weight: 800; color: #1E3A8A; margin-bottom: 0.5rem; }
    .sub-desc { color: #4B5563; font-size: 0.95rem; margin-bottom: 1.5rem; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">🇪🇺 유럽 일정표 자동 생성 시스템 (Master Edition)</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-desc">랜드사 실무 완판본 규격 | 항공 스케줄 실시간 연동 및 엑셀 일정표 자동화</div>', unsafe_allow_html=True)

# -------------------------------------------------------------
# 2. 항공 스케줄 검색 Add-on 렌더링
# -------------------------------------------------------------
if render_flight_search_ui:
    render_flight_search_ui()

# -------------------------------------------------------------
# 3. 엑셀 완판본 서식 생성 헬퍼 함수
# -------------------------------------------------------------
def build_excel_itinerary(schedule_data):
    """
    LLM이 생성한 일정 데이터를 랜드사 완판본 엑셀 서식으로 변환
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "행사일정표"
    ws.views.sheetView[0].showGridLines = True

    # 폰트 및 스타일 정의
    font_header = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
    font_bold = Font(name="맑은 고딕", size=9, bold=True)
    font_regular = Font(name="맑은 고딕", size=9)
    fill_navy = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    fill_day = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")

    thin_border_side = Side(style='thin', color='D9D9D9')
    border_all = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

    # 컬럼 헤더 설정
    headers = ["일차", "지역", "교통편", "시간", "일정 내용", "식사", "호텔"]
    ws.append(headers)

    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = font_header
        cell.fill = fill_navy
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_all
    ws.row_dimensions[1].height = 24

    # 데이터 입력
    days = schedule_data.get("days", [])
    for row_idx, d in enumerate(days, start=2):
        day_str = f"{d.get('day', row_idx-1)}일차"
        region = d.get("region", "")
        trans = d.get("transport", "")
        time_str = d.get("time", "")
        content = d.get("content", "")
        
        # 식사 조립
        meals = d.get("meals", {})
        meal_str = f"조: {meals.get('breakfast', '호텔식')}\n중: {meals.get('lunch', '현지식')}\n석: {meals.get('dinner', '한식')}"
        hotel = d.get("hotel", "투숙 호텔 미정 (동급)")

        ws.append([day_str, region, trans, time_str, content, meal_str, hotel])

        # 셀 서식 적용
        for col_idx in range(1, 8):
            c = ws.cell(row=row_idx, column=col_idx)
            c.border = border_all
            c.font = font_regular
            c.alignment = Alignment(vertical="center", wrap_text=True)

        ws.cell(row=row_idx, column=1).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=row_idx, column=1).fill = fill_day
        ws.cell(row=row_idx, column=2).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=row_idx, column=3).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=row_idx, column=4).alignment = Alignment(horizontal="center", vertical="center")

        ws.row_dimensions[row_idx].height = 65

    # 열 너비 설정
    col_widths = {1: 8, 2: 12, 3: 12, 4: 10, 5: 50, 6: 18, 7: 24}
    for col_idx, width in col_widths.items():
        col_letter = openpyxl.utils.get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = width

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# -------------------------------------------------------------
# 4. [옵션 B] 직접 루팅 입력 및 Gemini 생성 영역
# -------------------------------------------------------------
st.write("---")
st.markdown("### 📝 [옵션 B] 행사 루팅 입력 및 완판본 엑셀 생성")

routing_text = st.text_area(
    "일자별 루팅 및 항공 스케줄 입력",
    value=st.session_state.get("routing_text_injected", ""),
    height=260,
    placeholder="""1. 인천 출발 / [터키항공] TK021 ICN 10:45 ➔ IST 16:45 / 프라하 도착 후 호텔 투숙
2. 프라하 전일 시내 관광 (프라하 성, 카를교, 구시가지 광장 등)
3. 체스키크룸로프 이동 후 관광 / 잘츠부르크 이동
4. 잘츠부르크 시내 관광 / 비엔나 이동
5. 비엔나 전일 관광 (쉔부른 궁전, 성 슈테판 대성당)
6. 부다페스트 이동 / 야경 다뉴브 유람선 투어
7. 부다페스트 시내 전일 관광 (어부의 요새, 영웅광장)
8. 부다페스트 공항 이동 / [터키항공] TK1036 BUD 14:00 ➔ IST 17:15 / 기내 숙박
9. 인천 국제공항 도착 후 해산"""
)

col_act1, col_act2 = st.columns([1, 4])
with col_act1:
    generate_btn = st.button("🚀 엑셀 완판본 생성하기", type="primary")

# -------------------------------------------------------------
# 5. Gemini API 호출 및 엑셀 다운로드 처리
# -------------------------------------------------------------
if generate_btn:
    if not routing_text.strip():
        st.warning("⚠️ 루팅 텍스트를 입력하거나 상단 검색기에서 [👉 일정표 주입]을 눌러주세요.")
    else:
        api_key = st.secrets.get("GEMINI_API_KEY", os.environ.get("GEMINI_API_KEY", ""))
        if not api_key:
            st.error("❌ Secrets에 'GEMINI_API_KEY'가 설정되어 있지 않습니다.")
        else:
            with st.spinner("Gemini가 실무 완판본 규격으로 일정을 조립하고 엑셀을 렌더링 중입니다..."):
                try:
                    client = genai.Client(api_key=api_key)
                    
                    prompt = f"""
당신은 20년 경력의 유럽 패키지 전문 랜드사 수배 총괄 OP 팀장입니다.
아래 제공된 [일자별 루팅 및 항공 스케줄]을 분석하여 실무 일정표 엑셀에 들어갈 JSON 데이터 규격으로 작성하세요.

[필수 작성 규칙]
1. 관광지 명칭은 반드시 핵심 포인트 위주로 간결하고 정확하게 작성할 것.
2. 1일차 출국 및 귀국일(기내숙박 포함) 항공 시간과 편명을 누락 없이 정확히 반영할 것.
3. 식사는 조식(호텔식), 중식(현지식/특식), 석식(한식/현지식)을 자연스럽게 배정할 것.
4. 반드시 아래 JSON 포맷으로만 응답할 것 (추가 설명 금지):

{{
  "days": [
    {{
      "day": 1,
      "region": "인천/프라하",
      "transport": "항공/전용차량",
      "time": "10:45",
      "content": "인천 국제공항 출발\\n프라하 바츨라프 하벨 공항 도착 후 가이드 미팅\\n호텔 이동 및 체크인",
      "meals": {{ "breakfast": "기내식", "lunch": "기내식", "dinner": "현지식" }},
      "hotel": "일급(4성) 호텔 투숙"
    }}
  ]
}}

[일자별 루팅 및 항공 스케줄]
{routing_text}
"""
                    response = client.models.generate_content(
                        model='gemini-2.5-flash',
                        contents=prompt,
                        config=types.GenerateContentConfig(
                            response_mime_type="application/json"
                        )
                    )
                    
                    result_json = json.loads(response.text)
                    excel_data = build_excel_itinerary(result_json)
                    
                    st.success("🎉 완판본 엑셀 일정표 생성이 완료되었습니다!")
                    
                    st.download_button(
                        label="📥 엑셀 일정표 파일 다운로드 (.xlsx)",
                        data=excel_data,
                        file_name=f"유럽_행사일정표_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
                    
                except Exception as e:
                    st.error(f"❌ 일정표 생성 중 오류가 발생했습니다: {e}")
