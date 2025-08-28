# app.py
import os
import tempfile
import gradio as gr
import excel_summary_script as ess  # ваш файл

def run_build(input_file):
    if input_file is None:
        return None, "⚠️ Файл не загружен."

    try:
        wb = ess.build_summary_table(input_file.name)

        # сохраняем во временный .xlsx для скачивания
        fd, out_path = tempfile.mkstemp(suffix=".xlsx")
        os.close(fd)
        wb.save(out_path)

        return out_path, "✅ Готово! Нажмите, чтобы скачать итоговый файл."
    except Exception as e:
        return None, f"❌ Ошибка: {e}"

with gr.Blocks(title="Свод КП") as demo:
    gr.Markdown("## 📊 Свод КП из нескольких вкладок")
    file_in = gr.File(label="Загрузите Excel (.xlsx)", file_types=[".xlsx"])
    run_btn = gr.Button("Собрать свод")
    file_out = gr.File(label="Скачать результат")
    status = gr.Textbox(label="Статус", interactive=False)

    run_btn.click(run_build, inputs=file_in, outputs=[file_out, status])

if __name__ == "__main__":
    demo.launch()
