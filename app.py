# app.py
import os
import tempfile
import gradio as gr
from excel_summary_script import build_summary_table

# Храним загруженный файл в состоянии (после нажатия "Загрузить .xlsx")
def keep_file(file_obj):
    # UploadButton отдаёт объект с полем .name (путь к временному файлу)
    return file_obj

def make_summary(file_obj):
    if not file_obj:
        return None, "⚠️ Сначала загрузите .xlsx"

    # На всякий случай — поддержим и список (если вдруг file_count="multiple")
    path = getattr(file_obj, "name", None)
    if path is None and isinstance(file_obj, list) and file_obj:
        path = getattr(file_obj[0], "name", None)

    if not path:
        return None, "❌ Не удалось прочитать путь к файлу."

    try:
        wb = build_summary_table(path)
        fd, out_path = tempfile.mkstemp(suffix=".xlsx")
        os.close(fd)
        wb.save(out_path)
        return out_path, "✅ Готово! Скачайте итоговый файл."
    except Exception as e:
        return None, f"❌ Ошибка: {e}"

CSS = """
.gradio-container { max-width: 820px !important; margin: 0 auto !important; }
"""

with gr.Blocks(css=CSS, title="Comparison — свод КП") as demo:
    gr.Markdown("## 📊 Свод КП\nШаг 1 — загрузите Excel (.xlsx). Шаг 2 — нажмите **Сформировать свод**.")

    file_state = gr.State(None)  # сюда положим загруженный файл

    with gr.Row():
        # Кнопка загрузки: никаких дроп-зон, просто диалог выбора файла
        upload_btn = gr.UploadButton("📁 Загрузить .xlsx", file_types=[".xlsx"], file_count="single")
        run_btn = gr.Button("🚀 Сформировать свод", variant="primary")

    with gr.Row():
        out_file = gr.File(label="Скачать результат", interactive=False, type="filepath")
        status = gr.Textbox(label="Статус", interactive=False, lines=2)

    # После выбора файла — положим его в состояние
    upload_btn.upload(fn=keep_file, inputs=upload_btn, outputs=file_state)

    # По кнопке — берём файл из состояния и формируем свод
    run_btn.click(fn=make_summary, inputs=file_state, outputs=[out_file, status])

if __name__ == "__main__":
    demo.launch()
