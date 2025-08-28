# app.py
import os
import tempfile
import gradio as gr
import excel_summary_script as ess

def process_uploaded(file_obj):
    if not file_obj:
        return None, "⚠️ Сначала выберите .xlsx"
    try:
        wb = ess.build_summary_table(file_obj.name)
        fd, out_path = tempfile.mkstemp(suffix=".xlsx")
        os.close(fd)
        wb.save(out_path)
        return out_path, "✅ Готово! Скачайте итоговый файл."
    except Exception as e:
        return None, f"❌ Ошибка: {e}"

CSS = """
.gradio-container {max-width: 820px !important; margin: 0 auto !important;}
/* делаем элементы компактными и убираем любые растягивания */
#wrap {padding-top: 12px;}
"""

with gr.Blocks(css=CSS, title="Comparison — свод КП") as demo:
    gr.Markdown("## 📊 Свод КП\nНажмите кнопку ниже и выберите Excel (.xlsx).")

    with gr.Column(elem_id="wrap"):
        # ВАЖНО: используем UploadButton вместо File — никакой большой дроп-зоны
        upload = gr.UploadButton("📁 Загрузить .xlsx", file_types=[".xlsx"], file_count="single")
        file_out = gr.File(label="Скачать результат", interactive=False)
        status = gr.Textbox(label="Статус", interactive=False, lines=2)

    # Обрабатываем сразу после выбора файла
    upload.upload(process_uploaded, inputs=upload, outputs=[file_out, status])

if __name__ == "__main__":
    demo.launch()
