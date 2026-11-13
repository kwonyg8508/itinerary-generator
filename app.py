import streamlit as st
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta
import io
import json
import os
import re
import glob
import base64
import requests
from bs4 import BeautifulSoup
from google import genai
from google.genai import types

try:
    from flight_search import render_flight_search_ui
except ImportError:
    render_flight_search_ui = None

st.set_page_config(page_title="유럽 일정표 자동 생성 시스템 (Master Edition)", layout="wide", page_icon="🇪🇺")

if "generated_file" not in st.session_state: st.session_state.generated_file = None
if "generated_filename" not in st.session_state: st.session_state.generated_filename = None
if "ui_status" not in st.session_state: st.session_state.ui_status = "idle"
if "ui_msg" not in st.session_state: st.session_state.ui_msg = "선배님! 실시간 항공 검색 및 주입 기능이 다시 완벽하게 부활했습니다! ✈️☕"
if "routing_text_injected" not in st.session_state: st.session_state.routing_text_injected = ""
if "selected_input_mode" not in st.session_state: st.session_state.selected_input_mode = "[옵션 A] 간편 단문 요청"

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Pretendard:wght@400;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Pretendard', sans-serif; }
.assistant-card {
    background: #ffffff; border: 2px solid #e2e8f0; border-radius: 16px; padding: 16px;
    box-shadow: 0 4px 15px rgba(0, 0, 0, 0.04); display: flex; align-items: center; gap: 16px; margin-bottom: 20px;
}
.gif-wrapper {
    width: 135px; height: 125px; flex-shrink: 0; border-radius: 12px; overflow: hidden;
    border: 2px solid #cbd5e1; background: #0f172a; display: flex; align-items: center; justify-content: center;
}
.gif-wrapper img { width: 100%; height: 100%; object-fit: cover; image-rendering: pixelated; }
.speech-bubble {
    flex-grow: 1; background: #f8fafc; border: 1.5px solid #e2e8f0; border-radius: 12px;
    padding: 14px 16px; position: relative;
}
.speech-bubble::before {
    content: ''; position: absolute; left: -10px; top: 50%; transform: translateY(-50%);
    border-width: 6px 10px 6px 0; border-style: solid; border-color: transparent #e2e8f0 transparent transparent;
}
.assistant-header { display: flex; align-items: center; justify-content: space-between; margin-bottom: 6px; }
.assistant-name { font-size: 14px; font-weight: 700; color: #1e293b; display: flex; align-items: center; gap: 6px; }
.assistant-badge { font-size: 11px; font-weight: 600; padding: 3px 8px; border-radius: 20px; }
.badge-idle { background: #e2e8f0; color: #475569; }
.badge-working { background: #fee2e2; color: #ef4444; }
.badge-success { background: #dcfce7; color: #15803d; }
.assistant-message { font-size: 13.5px; font-weight: 500; color: #334155; line-height: 1.5; }
</style>
""", unsafe_allow_html=True)

st.title("🇪🇺 유럽 일정표 자동 생성 시스템 (Master Edition)")
st.caption("실무 완판 규격: 실시간 항공 검색 & 주입 | 편명 교통편 칸 배치 | 관광지 볼드 | 호텔 4성/5성 분기")

if render_flight_search_ui:
    render_flight_search_ui()

def get_available_db_files():
    db_folder = "database"
    if not os.path.exists(db_folder): return []
    return [os.path.basename(f) for f in glob.glob(os.path.join(db_folder, "*.xlsx"))]

def get_reference_db_context(user_query: str):
    db_folder = "database"
    excel_files = glob.glob(os.path.join(db_folder, "*.xlsx"))
    if not excel_files: return "", ""
    q_lower = user_query.lower()
    best_file = next((f for f in excel_files if any(kw in q_lower for kw in ["노르웨이", "발틱", "아이슬란드", "북유럽", "그리스", "로마", "이탈리아"])), excel_files[0])
    matched_filename = os.path.basename(best_file)
    try:
        wb = openpyxl.load_workbook(best_file, data_only=True)
        ws = wb.active
        content_summary = []
        for r in range(4, min(ws.max_row + 1, 40)):
            schedule = str(ws.cell(r, 5).value or "").strip()
            if schedule.startswith("▶"):
                content_summary.append(schedule[:50])
            if len(content_summary) >= 10: break
        return matched_filename, "\n".join(content_summary)
    except Exception:
        return "", ""

api_key = st.secrets.get("GEMINI_API_KEY")
client = genai.Client(api_key=api_key) if api_key else None

col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("1. 입력 방식 선택")
    input_mode_options = ["[옵션 A] 간편 단문 요청", "[옵션 B] 직접 루팅 입력 (정밀 항공/일정 파싱)"]
    selected_mode_idx = 1 if st.session_state.selected_input_mode == input_mode_options[1] else 0

    input_mode = st.radio("작성 방식을 선택하세요:", input_mode_options, index=selected_mode_idx, horizontal=True, key="radio_input_mode")
    st.session_state.selected_input_mode = input_mode

    start_date = st.date_input("출발일자 선택", datetime.today())
    tour_title = st.text_input("행사명 (선택)", placeholder="예: 아이슬란드 오로라 일주 7일 SAS")
    
    if "[옵션 A]" in input_mode:
        quick_prompt = st.text_input("간단 요청 / 웹 URL 입력", placeholder="예: 아이슬란드 오로라 일주 7일 SAS")
    else:
        current_injected_val = st.session_state.get("routing_text_injected", "")
        routing_text = st.text_area("루팅 및 일정 입력 / 웹 URL (상단 항공 검색에서 [일정표 주입] 시 자동 채워짐)", value=current_injected_val, height=280)

with col2:
    st.subheader("2. AI 수배팀 작업 현황")
    ui_placeholder = st.empty()

    def render_assistant_ui(status="idle", message="선배님, 항공 검색 후 주입하시거나 요청해 주시면 즉각 빌드하겠습니다! ☕"):
        badge_text, badge_class = ("🎉 제작 완료!", "badge-success") if status=="success" else (("🔥 폭풍 작업 중", "badge-working") if status=="working" else ("☕ 대기 중", "badge-idle"))
        cands = ["done.gif", "working.gif", "assistant.gif"]
        media_file = next((c for c in cands if os.path.exists(c)), None)
        media_b64 = f"data:image/gif;base64,{base64.b64encode(open(media_file,'rb').read()).decode()}" if media_file else ""
        img_tag = f'<img src="{media_b64}">' if media_b64 else '👓'
        
        ui_placeholder.markdown(f"""
        <div class="assistant-card">
            <div class="gif-wrapper">{img_tag}</div>
            <div class="speech-bubble">
                <div class="assistant-header">
                    <span class="assistant-name">👓 김똘망 사원 <span style="font-size:12px;color:#64748b;">(유럽 OP팀 막내)</span></span>
                    <span class="assistant-badge {badge_class}">{badge_text}</span>
                </div>
                <div class="assistant-message">{message}</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

    render_assistant_ui(st.session_state.ui_status, st.session_state.ui_msg)
    generate_btn = st.button("🚀 엑셀 일정표 생성하기", type="primary", use_container_width=True)

    if generate_btn:
        if not api_key: st.error("GEMINI_API_KEY가 등록되지 않았습니다.")
        else:
            st.session_state.generated_file = None
            raw_text_input = quick_prompt if "[옵션 A]" in input_mode else routing_text
            combined_text = f"{tour_title} {raw_text_input}".strip()
            
            ref_filename, ref_context = get_reference_db_context(combined_text)
            render_assistant_ui("working", "선택하신 항공 스케줄과 루팅을 정밀하게 조합하여 엑셀을 빌드 중입니다! ⚡")
            
            with st.spinner("김똘망 사원이 일정표를 신속하게 제작하고 있습니다..."):
                try:
                    system_prompt = f"""
당신은 유럽 단체 패키지 수배 15년 차 수석 오퍼레이터 김똘망입니다.
아래 [사용자 입력 및 항공 루팅]을 절대 법전으로 삼아 고객 전달용 완벽한 엑셀 일정표 JSON을 생성하세요.

[사용자 입력 및 항공 루팅 (1순위 절대 반영)]
행사명: {tour_title}
요청내용: {raw_text_input}

[참고용 가이드북 스니펫 (표현 참고용)]
{ref_context if ref_context else "(없음)"}

[🚨 랜드사 실무 절대 규칙 - 위반 절대 금지]
1. **항공 시간/편명 정밀 반영**:
   - 사용자가 주입하거나 입력한 항공편명과 시간을 그대로 반영하세요.
   - 편명은 교통편(`transport`) 칸에 단독으로 기재하고, 시간(`time`) 칸에는 정확한 `HH:MM`을 기재하세요.
2. **도보 이동수단 전면 삭제**:
   - 시내 관광 중 걷는 것은 교통편에 적지 않으며, 교통편 칸을 빈칸("")으로 비워두세요.
3. **'전용차량 / 전  일' 2행 고정 및 반복 금지**:
   - 2일차부터 차량 투어 시작 시 **2번째 행(index 1)에만** `transport="전용차량"`, `time="전  일"`을 고정하고 반복 표기하지 마세요.
4. **관광 명소 부분 볼드체 서식 (매우 중요)**:
   - 관광지 설명 행은 반드시 **`▶ [설명] **[핵심명소]**`** 형식을 준수하여 핵심 명소 이름에 마크다운 볼드(`**`)를 적용하세요.
5. **호텔 명기 규칙**:
   - 튀르키예 지역은 `HOTEL : ANY 5* HOTEL CLASS`, 그 외 유럽 지역은 무조건 **`HOTEL : ANY 4* HOTEL CLASS`**로 표기하세요.
6. **이동 양식 통일**:
   - 도시 간 이동은 반드시 **'[도착도시명]으로 이동 (약 O시간 O분 소요)'** 형식으로만 작성하세요.
7. **식사 내용 일정 칸 기재 원천 차단**:
   - 일정 텍스트 안에 식사 종류를 적지 말고 우측 식사 칸(조/중/석)에서만 관리하세요.
"""

                    json_schema = {
                        "type": "ARRAY",
                        "items": {
                            "type": "OBJECT",
                            "properties": {
                                "day": {"type": "INTEGER"},
                                "rows": {
                                    "type": "ARRAY",
                                    "items": {
                                        "type": "OBJECT",
                                        "properties": {
                                            "city": {"type": "STRING"},
                                            "transport": {"type": "STRING"},
                                            "time": {"type": "STRING"},
                                            "schedule": {"type": "STRING"}
                                        },
                                        "required": ["schedule"]
                                    }
                                },
                                "meals": {"type": "ARRAY", "items": {"type": "STRING"}},
                                "hotel": {"type": "STRING"}
                            },
                            "required": ["day", "rows"]
                        }
                    }

                    response = client.models.generate_content(
                        model='gemini-3.8-flash',
                        contents=system_prompt,
                        config=types.GenerateContentConfig(response_mime_type="application/json", response_schema=json_schema)
                    )

                    raw_text = response.text.strip()
                    if "```" in raw_text:
                        raw_text = re.sub(r"^```(?:json)?\s*", "", raw_text, flags=re.MULTILINE)
                        raw_text = re.sub(r"\s*```$", "", raw_text, flags=re.MULTILINE)
                    
                    parsed_raw = json.loads(raw_text)
                    day_dict = {}
                    for item in parsed_raw:
                        d_num = item.get("day", 1)
                        if d_num not in day_dict:
                            day_dict[d_num] = {"day": d_num, "rows": [], "meals": item.get("meals") or ["-", "-", "-"], "hotel": item.get("hotel") or ""}
                        for r_item in item.get("rows", []):
                            sc = str(r_item.get("schedule") or "").strip()
                            if sc in ["석식 후", "석식후"]: continue
                            r_item["schedule"] = sc.replace("[", "").replace("]", "")
                            day_dict[d_num]["rows"].append(r_item)
                        if item.get("hotel"): day_dict[d_num]["hotel"] = str(item.get("hotel")).replace("[", "").replace("]", "").strip()
                        if item.get("meals"): day_dict[d_num]["meals"] = item.get("meals")

                    parsed_days = [day_dict[k] for k in sorted(day_dict.keys())]
                    total_days_cnt = len(parsed_days)

                    turkey_keywords = ["이스탄불", "카파도키아", "파묵칼레", "터키", "튀르키예", "안탈리아"]

                    for item in parsed_days:
                        d_num = item["day"]
                        rows = item["rows"]
                        
                        cleaned_rows = []
                        for r in rows:
                            sc = str(r.get("schedule") or "").strip()
                            tr = str(r.get("transport") or "").strip()
                            if tr in ["도보", "도보이동"]: continue
                            if any(k in sc for k in ["중식(", "석식(", "조식(", "중식 (", "석식 ("]): continue
                            cleaned_rows.append(r)
                        rows = cleaned_rows

                        seen_cities_in_day = set()
                        for r in rows:
                            c_val = str(r.get("city") or "").strip()
                            if c_val:
                                if c_val in seen_cities_in_day:
                                    r["city"] = ""
                                else:
                                    seen_cities_in_day.add(c_val)

                        if 1 < d_num < total_days_cnt and len(rows) > 1:
                            for idx, r in enumerate(rows):
                                if idx > 1:
                                    if r.get("transport") == "전용차량": r["transport"] = ""
                                    if r.get("time") == "전  일": r["time"] = ""
                                    if str(r.get("schedule", "")).startswith("▶"): r["time"] = ""
                            if not rows[1].get("time") or rows[1].get("time") == "전  일":
                                rows[1]["transport"] = "전용차량"
                                rows[1]["time"] = "전  일"

                        item["rows"] = rows

                        if d_num == total_days_cnt - 1:
                            item["hotel"] = "HOTEL : 기내 숙박"
                        elif d_num == total_days_cnt:
                            item["hotel"] = ""
                        else:
                            all_text = " ".join([r.get("schedule", "") + " " + str(r.get("city", "")) for r in rows])
                            is_turkey = any(tk in all_text for tk in turkey_keywords)
                            item["hotel"] = "HOTEL : ANY 5* HOTEL CLASS" if is_turkey else "HOTEL : ANY 4* HOTEL CLASS"

                    wb = openpyxl.Workbook()
                    ws = wb.active
                    ws.title = "일정표"
                    ws.views.sheetView[0].showGridLines = True
                    
                    font_title = Font(name="맑은 고딕", size=15, bold=True)
                    font_header = Font(name="맑은 고딕", size=10, bold=True)
                    font_body = Font(name="맑은 고딕", size=9)
                    font_bold = Font(name="맑은 고딕", size=9, bold=True)
                    font_notice_red = Font(name="맑은 고딕", size=11, bold=True, color="FF0000")
                    
                    in_font_regular = InlineFont(rFont="맑은 고딕", sz=9.0, b=False)
                    in_font_bold = InlineFont(rFont="맑은 고딕", sz=9.0, b=True)
                    
                    fill_title = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")
                    fill_hotel = PatternFill(start_color="EFEFEF", end_color="EFEFEF", fill_type="solid")
                    align_center = Alignment(horizontal="center", vertical="center")
                    align_mid_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
                    thin_side = Side(style="thin", color="000000")
                    med_side = Side(style="medium", color="000000")
                    
                    full_title = tour_title.strip() if tour_title.strip() else f"{start_date.strftime('%Y-%m-%d')} 유럽투어 일정표"
                    ws.row_dimensions[1].height = 28.0
                    ws.cell(1, 1, full_title).font = font_title
                    ws.cell(1, 1).alignment = align_center
                    
                    for c in range(1, 8):
                        cell = ws.cell(1, c)
                        cell.fill = fill_title
                        cell.border = Border(top=med_side, bottom=med_side, left=med_side if c == 1 else None, right=med_side if c == 7 else None)
                    ws.merge_cells("A1:G1")
                    ws.row_dimensions[2].height = 13.0
                    
                    headers = ["날짜", "도시", "교통편", "시간", "일정", "식사", ""]
                    ws.row_dimensions[3].height = 22.0
                    for c in range(1, 8):
                        cell = ws.cell(3, c, headers[c-1])
                        cell.font = font_header
                        cell.alignment = align_center
                        cell.border = Border(top=med_side, bottom=med_side, left=med_side if c==1 else thin_side, right=med_side if c==7 else thin_side)
                    ws.merge_cells("F3:G3")
                    
                    col_widths = {1: 10.0, 2: 12.75, 3: 11.0, 4: 9.25, 5: 64.25, 6: 3.63, 7: 10.0}
                    for c_idx, width in col_widths.items(): ws.column_dimensions[get_column_letter(c_idx)].width = width
                    
                    days_korean = ["월요일", "화요일", "수요일", "목요일", "금요일", "토요일", "일요일"]
                    cur_row = 4
                    
                    for item_idx, item in enumerate(parsed_days):
                        day_num = item.get("day", item_idx + 1)
                        curr_date = start_date + timedelta(days=day_num - 1)
                        day_str, date_str, weekday_str = f"제{day_num:02d}일", curr_date.strftime("%m월 %d일"), days_korean[curr_date.weekday()]
                        
                        rows_data, hotel_str, meals = item.get("rows", []), str(item.get("hotel") or "").strip(), item.get("meals") or ["-", "-", "-"]
                        is_last_day = (item_idx == len(parsed_days) - 1)
                        has_hotel = bool(hotel_str) and not is_last_day
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
                                if is_last_row and has_hotel: cell.fill = fill_hotel
                                cell.border = Border(top=b_top, bottom=b_bottom, left=b_left, right=b_right)
                        
                        ws.cell(cur_row, 1, day_str).font = font_bold
                        ws.cell(cur_row, 1).alignment = align_center
                        ws.cell(cur_row + 1, 1, date_str).alignment = align_center
                        ws.cell(cur_row + 2, 1, weekday_str).alignment = align_center
                        
                        for idx, r_data in enumerate(rows_data):
                            r = cur_row + idx
                            sc_text = str(r_data.get("schedule") or "")
                            if r_data.get("city"): ws.cell(r, 2, str(r_data.get("city"))).alignment = align_center
                            if r_data.get("transport"): ws.cell(r, 3, str(r_data.get("transport"))).alignment = align_center
                            if r_data.get("time"): ws.cell(r, 4, str(r_data.get("time"))).alignment = align_center
                                
                            cell_sc = ws.cell(r, 5)
                            cell_sc.alignment = align_mid_left
                            
                            if sc_text.strip().startswith("▶"):
                                bold_match = re.search(r"^(.*?)\*\*(.+?)\*\*(.*)$", sc_text)
                                if bold_match:
                                    p_txt, b_txt, s_txt = bold_match.group(1), bold_match.group(2), bold_match.group(3)
                                    rich_elems = []
                                    if p_txt: rich_elems.append(TextBlock(in_font_regular, p_txt))
                                    if b_txt: rich_elems.append(TextBlock(in_font_bold, b_txt))
                                    if s_txt: rich_elems.append(TextBlock(in_font_regular, s_txt))
                                    cell_sc.value = CellRichText(rich_elems)
                                else:
                                    cell_sc.value = sc_text
                                    cell_sc.font = font_body
                            else:
                                cell_sc.value = sc_text
                                cell_sc.font = font_body
                        
                        if has_hotel:
                            h_cell = ws.cell(cur_row + total_rows - 1, 5, hotel_str)
                            h_cell.font = font_bold
                            h_cell.alignment = align_mid_left
                        
                        for m_idx in range(3):
                            ws.cell(cur_row + m_idx, 6, ["조", "중", "석"][m_idx]).alignment = align_center
                            ws.cell(cur_row + m_idx, 7, str(meals[m_idx] if m_idx < len(meals) else "-")).alignment = align_center
                        
                        cur_row += total_rows
                    
                    notice_row = cur_row + 1
                    ws.row_dimensions[notice_row].height = 22.0
                    notice_cell = ws.cell(notice_row, 1, "* 상기 일정은 현지 사정에 따라 변경될 수 있습니다 *")
                    notice_cell.font = font_notice_red
                    notice_cell.alignment = align_center
                    ws.merge_cells(start_row=notice_row, start_column=1, end_row=notice_row, end_column=7)
                    
                    output = io.BytesIO()
                    wb.save(output)
                    st.session_state.generated_file = output.getvalue()
                    st.session_state.generated_filename = f"{start_date.strftime('%Y-%m-%d')} {tour_title.strip() or '유럽투어'} - 일정표.xlsx"
                    st.session_state.ui_status = "success"
                    st.session_state.ui_msg = "선배님! 실시간 검색된 항공 스케줄이 반영된 최종 마스터 엑셀이 완성되었습니다! 🙇‍♂️✨"
                    st.rerun()

                except Exception as e:
                    st.session_state.ui_status = "working"
                    st.session_state.ui_msg = f"오류 발생: {str(e)}"
                    st.error(f"오류 상세: {str(e)}")

    if st.session_state.generated_file:
        st.success("✅ 파일 준비 완료!")
        st.download_button(
            label="📥 완성된 실무 마스터 엑셀 일정표 다운로드",
            data=st.session_state.generated_file,
            file_name=st.session_state.generated_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
