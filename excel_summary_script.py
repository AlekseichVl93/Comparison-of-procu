
import openpyxl
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment
import re

def normalize_string(s):
    if not isinstance(s, str):
        return ""
    return re.sub(r'\s+', ' ', s.strip().lower())

def get_sheet_data(sheet):
    data = []
    for row in sheet.iter_rows(values_only=True):
        data.append(list(row))
    return data

def find_max_row(sheet):
    for row in reversed(range(1, sheet.max_row + 1)):
        if any(cell.value is not None for cell in sheet[row]):
            return row
    return 0

def build_summary_table(filename):
    wb = load_workbook(filename)
    sheet_names = wb.sheetnames[1:]  # начиная со второй вкладки

    summary_wb = Workbook()
    summary_ws = summary_wb.active
    summary_ws.title = "Сводная таблица"

    # Заголовки строк 1 и 2
    headers_row_1 = ["Наименование", "Количество запрошенное"]
    headers_row_2 = ["", ""]
    for name in sheet_names:
        headers_row_1 += [name, None, None, None]
        headers_row_2 += [
            "Количество предложенное",
            "Цена без НДС за шт",
            "Сроки поставки",
            "Комментарий поставщика"
        ]

    summary_ws.append(headers_row_1)
    summary_ws.append(headers_row_2)

    for col in range(3, len(headers_row_1) + 1, 4):
        summary_ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col+3)
        summary_ws.cell(row=1, column=col).alignment = Alignment(horizontal="center", vertical="center")

    all_positions = {}

    for idx, sheet_name in enumerate(sheet_names):
        sheet = wb[sheet_name]
        max_row = find_max_row(sheet)
        for row in range(2, max_row + 1):
            name = normalize_string(sheet[f"A{row}"].value)
            if not name:
                continue

            key = name
            quantity = sheet[f"B{row}"].value
            values = [sheet[f"{col}{row}"].value for col in "CDEF"]

            if key not in all_positions:
                all_positions[key] = {
                    "Наименование": sheet[f"A{row}"].value,
                    "Количество": sheet[f"B{row}"].value,
                    "data": [[] for _ in range(len(sheet_names))]
                }
            if idx == 0:
                all_positions[key]["Количество"] = quantity
            all_positions[key]["data"][idx] = values

    written = {}
    for position in all_positions.values():
        row = [
            position["Наименование"],
            position["Количество"]
        ]
        for block in position["data"]:
            row += block if block else [""] * 4

        key = normalize_string(position["Наименование"])
        if key in written:
            for i in range(2, len(row)):
                if row[i]:
                    summary_ws.cell(row=written[key], column=i+1, value=row[i])
        else:
            summary_ws.append(row)
            written[key] = summary_ws.max_row

    return summary_wb

# Сохранение итогового файла
input_path = "выгрузка под скрипт 6 кп.xlsx"
output_path = "свод_таблица_итог.xlsx"
summary_wb = build_summary_table(input_path)
summary_wb.save(output_path)
print(f"Готово! Итоговый файл сохранен как: {output_path}")
