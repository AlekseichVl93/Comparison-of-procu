# app.py
import os
import tempfile
import gradio as gr
import excel_summary_script as ess  # ваш модуль

def run_build(file_obj):
    """Принимает загруженный .xlsx из Gradio, возвращает путь к итоговому файлу + статус."""
    if not file_obj:
        return None, "⚠️ Сначала загрузите .xlsx"

    try:
        # file_obj — это TempFile, путь лежит в .name
        wb = ess.build_summary_table(file_obj.name)

        # Сохраняем во временный .xlsx, чтобы отдать на скачивание
        fd, out_path = tempfile.mkstemp(suffix=".xlsx")
        os.close(fd)
        wb.save(out_path)

        return out_path, "✅ Готово! Нажмите на файл, чтобы скачать."
    except Exception as e:
        return None, f"❌ Ошибка обработки: {e}"

with gr.Blocks(
    title="Comparison — свод КП",
    theme=gr.themes.Soft(),  # приятная светлая тема
    fill_height=True
) as demo:
    gr.Markdown(
        "## 📊 Свод КП из нескольких вкладок\n"
        "Загрузите исходный Excel (.xlsx), затем нажмите **Собрать свод**."
    )

    with gr.Row():
        with gr.Column(scale=1):
            file_in = gr.File(
                label="Загрузите Excel (.xlsx)",
                file_types=[".xlsx"],
                type="filepath",
                height=200
            )
            run_btn = gr.Button("🚀 Собрать свод", variant="primary", scale=1)
        with gr.Column(scale=1):
            file_out = gr.File(label="Скачать результат", interactive=False)
            status = gr.Textbox(label="Статус", interactive=False, lines=2)

    run_btn.click(run_build, inputs=file_in, outputs=[file_out, status])

# В Spaces можно оставить этот блок — локально тоже запустится.
if __name__ == "__main__":
    demo.launch(server_name="0.0.0.0", server_port=7860)
