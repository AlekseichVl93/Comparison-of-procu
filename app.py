# app.py
import os
import tempfile
import gradio as gr
import excel_summary_script as ess  # ваш модуль

def run(file_obj):
    if not file_obj:
        return None, "⚠️ Сначала выберите .xlsx"
    try:
        wb = ess.build_summary_table(file_obj.name)
        fd, out_path = tempfile.mkstemp(suffix=".xlsx")
        os.close(fd)
        wb.save(out_path)
        return out_path, "✅ Готово! Нажмите на файл выше, чтобы скачать."
    except Exception as e:
        return None, f"❌ Ошибка: {e}"

CSS = """
.gradio-container {max-width: 900px !important; margin: 0 auto !important;}
#upload-col {min-width: 420px;}
"""

with gr.Blocks(css=CSS) as demo:
    gr.Markdown("## 📊 Свод КП\nЗагрузите Excel (.xlsx) и нажмите **Собрать свод**.")
    with gr.Row():
        with gr.Column(elem_id="upload-col"):
            file_in = gr.File(
                label="Загрузите Excel (.xlsx)",
                file_types=[".xlsx"],
                type="filepath",
                height=200
            )
            btn = gr.Button("Собрать свод", variant="primary")
        with gr.Column():
            file_out = gr.File(label="Скачать результат")
            status = gr.Textbox(label="Статус", interactive=False, lines=2)

    btn.click(run, inputs=file_in, outputs=[file_out, status])

if __name__ == "__main__":
    demo.launch()
