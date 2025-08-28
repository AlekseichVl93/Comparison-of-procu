import gradio as gr
from excel_summary_script import build_summary_table

def process_file(file):
    if file is None:
        return None
    # Формируем сводный файл
    summary_wb = build_summary_table(file.name)
    output_path = "summary_output.xlsx"
    summary_wb.save(output_path)
    return output_path

with gr.Blocks() as demo:
    gr.Markdown("## 📊 Сравнитель КП")
    gr.Markdown("Загрузите Excel-файл и получите сводный отчёт")

    with gr.Row():
        input_file = gr.File(label="Загрузите Excel (.xlsx)", file_types=[".xlsx"], type="file")
        output_file = gr.File(label="Скачать результат", type="file")

    run_btn = gr.Button("Сформировать свод")

    run_btn.click(fn=process_file, inputs=input_file, outputs=output_file)

if __name__ == "__main__":
    demo.launch()
