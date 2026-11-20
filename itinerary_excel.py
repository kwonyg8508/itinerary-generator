from copy import copy
from datetime import datetime, timedelta
from io import BytesIO
from pathlib import Path
import re

from openpyxl import load_workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.worksheet.pagebreak import Break


TEMPLATE = Path(__file__).with_name("itinerary-template.xlsx")
BODY_ROW_HEIGHT = 25
WEEKDAYS = ("월요일", "화요일", "수요일", "목요일", "금요일", "토요일", "일요일")


def _set_text(cell, value):
    cell.value = value
    if cell.data_type == "f":
        cell.data_type = "s"


def _write_schedule(cell, value):
    parts = re.split(r"\*\*(.+?)\*\*", str(value or ""))
    if len(parts) == 1:
        _set_text(cell, value or None)
        return
    font = InlineFont(rFont=cell.font.name, sz=cell.font.sz, b=True)
    cell.value = CellRichText([
        TextBlock(font, part) if index % 2 else part
        for index, part in enumerate(parts) if part
    ])


def _display_transport(row):
    transport = str(row.get("transport") or "").strip()
    time = row.get("time") or ""
    schedule = str(row.get("schedule") or "")
    if transport == "전용버스":
        return transport, "전일"

    # A row that also describes boarding the next service may carry its name.
    arrival = re.search(r"도착|하선|하차|착륙|입항", schedule)
    departure = re.search(r"출발|출항|이륙|향발|탑승|승선", schedule)
    if arrival and not departure:
        return None, time
    return transport or None, time


def _order_departure_day(rows):
    first_departure = next((i for i, row in enumerate(rows) if row.get("flight_event") == "dep"), None)
    if first_departure is None:
        return rows
    meetings = []
    remaining = []
    for row in rows:
        schedule = str(row.get("schedule") or "")
        if not row.get("flight_event") and re.search(r"인천|ICN", schedule) and re.search(r"미팅|집결|출국 수속|출국수속", schedule):
            meetings.append({**row, "time": ""})
        else:
            remaining.append(row)
    index = next(i for i, row in enumerate(remaining) if row.get("flight_event") == "dep")
    return remaining[:index] + meetings + remaining[index:]


def render_itinerary(days, start_date, title):
    if not days:
        raise ValueError("출력할 일정이 없습니다.")
    workbook = load_workbook(TEMPLATE, rich_text=True, keep_links=False)
    # The itinerary template uses static cells; legacy names can still point outside it.
    workbook.defined_names.clear()
    sheet = workbook.active
    sheet.defined_names.clear()
    row_styles = {
        number: tuple(copy(cell._style) for cell in sheet[number][:7])
        for number in (4, 5, 6, 7, 8, 10, 13, 22, 23, 55, 72, 75, 76)
    }
    header = [cell.value for cell in sheet[4][:7]]
    notice = sheet["A76"].value
    date_format = sheet["A6"].number_format
    sheet.unmerge_cells("A76:G76")
    sheet.delete_rows(5, 73)
    for number in list(sheet.row_dimensions):
        if number >= 5:
            del sheet.row_dimensions[number]
    sheet.row_breaks.brk.clear()
    _set_text(sheet["A2"], title)

    def style_row(number, source, height=BODY_ROW_HEIGHT):
        sheet.row_dimensions[number].height = height
        for col, style in enumerate(row_styles[source], 1):
            sheet.cell(number, col)._style = copy(style)

    current = 5
    used_height = 0
    for item in days:
        day = item["day"]
        rows = _order_departure_day(item["rows"]) if day == 1 else item["rows"]
        meals = item.get("meals") or []
        hotel = item.get("hotel")
        size = max(3, len(rows))
        block_height = BODY_ROW_HEIGHT * (size + bool(hotel))
        # Repeat the sample's spacer and header after approximately one page of rows.
        if used_height and used_height + block_height > 900:
            style_row(current, 55, 11.25)
            sheet.row_breaks.append(Break(id=current, max=6))
            current += 1
            style_row(current, 4, 15)
            for column, value in enumerate(header, 1):
                sheet.cell(current, column).value = value
            current += 1
            used_height = 0

        for offset in range(size):
            source = (5, 6, 7)[offset] if offset < 3 else 8
            if offset == size - 1 and not hotel and offset >= 2:
                source = 75
            row_number = current + offset
            style_row(row_number, source)
            if offset >= len(rows):
                continue
            row = rows[offset]
            schedule = row.get("schedule") or ""
            if schedule.lstrip().startswith(("▶", "※")):
                sheet.cell(row_number, 5)._style = copy(row_styles[22 if schedule.lstrip().startswith("▶") else 23][4])
            _set_text(sheet.cell(row_number, 2), row.get("city") or None)
            transport, time = _display_transport(row)
            for column in (3, 4):
                cell = sheet.cell(row_number, column)
                font = copy(cell.font)
                font.bold = False
                cell.font = font
            _set_text(sheet.cell(row_number, 3), transport)
            if re.fullmatch(r"(?:[01]?\d|2[0-3]):[0-5]\d", time):
                sheet.cell(row_number, 4).value = datetime.strptime(time, "%H:%M").time()
                sheet.cell(row_number, 4).number_format = "hh:mm"
            else:
                _set_text(sheet.cell(row_number, 4), time or None)
            _write_schedule(sheet.cell(row_number, 5), schedule)

        day_date = start_date + timedelta(days=day - 1)
        sheet.cell(current, 1).value = f"제{day:02d}일"
        sheet.cell(current + 1, 1).value = day_date
        sheet.cell(current + 1, 1).number_format = date_format
        sheet.cell(current + 2, 1).value = WEEKDAYS[day_date.weekday()]
        visible_meals = list(zip("조중석", meals))
        visible_meals = [(label, value) for label, value in visible_meals if value and value != "-"]
        for offset, (label, value) in enumerate(visible_meals):
            sheet.cell(current + offset, 6).value = label
            _set_text(sheet.cell(current + offset, 7), value)
        current += size
        if hotel:
            style_row(current, 72 if "기내" in hotel else 10)
            _set_text(sheet.cell(current, 5), hotel)
            current += 1
        used_height += block_height

    style_row(current, 76, 15)
    sheet.cell(current, 1).value = notice
    sheet.merge_cells(start_row=current, start_column=1, end_row=current, end_column=7)
    sheet.print_area = f"A1:G{current}"
    output = BytesIO()
    workbook.save(output)
    return output.getvalue()
