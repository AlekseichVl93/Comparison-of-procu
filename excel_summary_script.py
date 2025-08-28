import openpyxl
from openpyxl.styles import Alignment


def build_summary_table(filename: str) -> openpyxl.Workbook:
    """
    Строит сводную книгу по правилам упрощённого варианта:
    - Заголовки в 2 строки.
    - A1:A2 и B1:B2 объединены.
    - Для каждой вкладки со 2-й по последнюю создаётся группа из 4 колонок:
      (Количество предложенное, Цена без НДС за шт, Сроки поставки, Комментарий поставщика),
      а её заголовок (имя вкладки) объединяется на первой строке.
    - Данные собираются по полному совпадению наименования (ключ — колонка A на вкладках-поставщиках).
    """

    # Загружаем исходный Excel
    wb_src = openpyxl.load_workbook(filename)

    # Создаём новую книгу для свода
    wb_out = openpyxl.Workbook()
    ws = wb_out.active
    ws.title = "Свод"

    # ------- Заголовки -------
    headers_row_1 = ["Наименование", "Количество запрошенное"]
    headers_row_2 = ["", ""]

    # Берём вкладки начиная со второй
    sheet_names = wb_src.sheetnames[1:]

    # Для каждой вкладки добавляем её «шапку» (объединяем потом по 4 колонки)
    for sheet_name in sheet_names:
        headers_row_1.extend([sheet_name, "", "", ""])
        headers_row_2.extend([
            "Количество предложенное",
            "Цена без НДС за шт",
            "Сроки поставки",
            "Комментарий поставщика",
        ])

    ws.append(headers_row_1)
    ws.append(headers_row_2)

    # Объединяем A1:A2 и B1:B2 и центрируем
    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)
    ws.merge_cells(start_row=1, start_column=2, end_row=2, end_column=2)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws["B1"].alignment = Alignment(horizontal="center", vertical="center")

    # Объединяем заголовки по поставщикам (каждые 4 колонки)
    for col in range(3, len(headers_row_1) + 1, 4):
        ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + 3)
        ws.cell(row=1, column=col).alignment = Alignment(horizontal="center", vertical="center")

    # ------- Сбор данных по всем вкладкам со 2-й -------
    # data = { "наименование": { "Количество запрошенное": X, "поставщики": {sheet_name: (C,D,E,F)} } }
    data = {}

    for sheet_name in sheet_names:
        ws_src = wb_src[sheet_name]
        # ожидается, что строки начинаются со 2-й (первая — заголовки)
        for row in ws_src.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]:
                continue

            key = row[0].strip() if isinstance(row[0], str) else row[0]
            qty_requested = row[1]  # что в колонке B на вкладке-поставщике
            values_cdef = row[2:6]  # C..F: (Кол-во предлож., Цена, Сроки, Комментарий)

            if key not in data:
                # Берём "Количество запрошенное" из первой встреченной вкладки
                data[key] = {"Количество запрошенное": qty_requested, "поставщики": {}}

            data[key]["поставщики"][sheet_name] = values_cdef

    # ------- Заполнение свода -------
    out_row = 3
    for name, payload in data.items():
        ws.cell(row=out_row, column=1, value=name)
        ws.cell(row=out_row, column=2, value=payload["Количество запрошенное"])

        col = 3
        for sheet_name in sheet_names:
            vals = payload["поставщики"].get(sheet_name)
            if vals:
                for i in range(4):
                    ws.cell(row=out_row, column=col + i, value=vals[i])
            col += 4

        out_row += 1

    return wb_out


# -------- локальный запуск (на Hugging Face не выполняется) --------
if __name__ == "__main__":
    # Пример локального запуска: подставьте свои пути
    input_path = "выгрузка под скрипт 6 кп.xlsx"
    output_path = "свод_таблица_итог.xlsx"

    wb_result = build_summary_table(input_path)
    wb_result.save(output_path)
    print(f"Готово! Итоговый файл сохранён как: {output_path}")
