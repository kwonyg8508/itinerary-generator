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
1. 공항 명칭 금지: '국제공항', '터미널', '반타', '카스트룹' 등 공항 명칭 일체 쓰지 마세요. 무조건 '인천', '코펜하겐', '헬싱키' 등 도시명만 씁니다.
2. 시간란 규칙: '오전', '오후', '저녁', '출발', '도착' 등의 한글을 절대 넣지 마세요. 오직 'HH:MM' 형태의 순수 숫자 시간이나 '전  일'만 허용됩니다.
3. 교통편('전용차량') 표기 엄격 제한:
   - 일반 육로 관광일의 경우: 1행(호텔 조식 후 또는 하선 후) 바로 다음 줄인 2행에만 transport="전용차량", time="전  일"을 단 1회 기재합니다.
   - 관광지 이동, 호텔 이동, 특히 "공항으로 이동" 같은 마무리 이동 행에는 transport를 절대 넣지 말고 무조건 빈칸("")으로 두세요.
4. 항공편 및 경유지 분리 규칙 (중요):
   - 도시 이름은 무조건 city 칸에만 들어가야 합니다. transport 칸에 도시 이름을 넣는 것은 절대 금지입니다.
   - transport 칸에는 오직 순수 항공 편명(예: "SK1717", "SK987")만 들어갑니다.
   - 항공 출발/도착 시간은 각각 해당 행의 time 칸에 반드시 둘 다 숫자로 적으세요 (출발 시간만 있으면 예상 도착 시간을 계산해서 채울 것).
   - 경유 항공편 처리 방식 예시 (헬싱키/코펜하겐/인천 귀국일):
     * 헬싱키 공항 출발 행: city="헬싱키", transport="SK1717", time="21:05", schedule="헬싱키 출발, 코펜하겐 향발"
     * 코펜하겐 도착 및 환승 행: city="코펜하겐", transport="", time="21:45", schedule="코펜하겐 도착 후 환승"
     * 코펜하겐 출발 행: city="", transport="SK987", time="23:55", schedule="코펜하겐 출발, 인천 향발"
     * 다음 날 인천 도착 행: city="인  천", transport="", time="18:35", schedule="인천 도착 후 해산"
5. 관광지 행 분리:
   - 한 행에 여러 관광지를 몰아넣지 마세요. 무조건 1개 관광지당 rows의 1개 객체로 분리하세요.
   - 관광지 앞에는 ' ▶ '를 붙이고 '~한 [명소명]' 문체를 유지하세요. 사용자의 필수 스팟을 포함하여 3~5개 명소를 채우세요.
6. 일정 마무리 및 호텔 표기:
   - 호텔 투숙일: 마지막 전 행="석식 후", 마지막 행="호텔 투숙 및 휴식", hotel="HOTEL : ANY 4* HOTEL CLASS"
   - 크루즈 투숙일: 마지막 전 행="선내 뷔페 석식 후", 마지막 행="선내 투숙 및 휴식", hotel="CRUISE : DFDS SEAWAYS (SEAVIEW 2-BERTH)" 또는 "CRUISE : SILJA LINE (SEAVIEW 2-BERTH)"
   - 기내 숙박일(야간 비행): 별도 투숙 문구 없이 항공 일정으로 끝나며, hotel="HOTEL : 기내 숙박"
   - 마지막 날 귀국: "인천 도착 후 해산"으로 마무리하며, hotel="" (빈칸)

[반드시 아래 예시와 정확히 동일한 구조의 JSON 포맷으로만 응답]
[
  {{
    "day": 9,
    "rows": [
      {{"city": "", "transport": "", "time": "", "schedule": "호텔 조식 후"}},
      {{"city": "헬싱키", "transport": "전용차량", "time": "전  일", "schedule": "헬싱키 시내 관광"}},
      {{"city": "", "transport": "", "time": "", "schedule": " ▶ 웅장한 백색의 랜드마크 헬싱키 대성당"}},
      {{"city": "", "transport": "", "time": "", "schedule": "공항으로 이동"}},
      {{"city": "헬싱키", "transport": "SK1717", "time": "21:05", "schedule": "헬싱키 출발, 코펜하겐 향발"}},
      {{"city": "코펜하겐", "transport": "", "time": "21:45", "schedule": "코펜하겐 도착 후 환승"}},
      {{"city": "", "transport": "SK987", "time": "23:55", "schedule": "코펜하겐 출발, 인천 향발"}}
    ],
    "meals": ["호텔식", "현지식", "기내식"],
    "hotel": "HOTEL : 기내 숙박"
  }},
  {{
    "day": 10,
    "rows": [
      {{"city": "인  천", "transport": "", "time": "18:35", "schedule": "인천 도착 후 해산"}}
    ],
    "meals": ["기내식", "-", "-"],
