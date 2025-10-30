# app.py
import os
import tempfile
import gradio as gr
import excel_summary_script as ess  # ваш файл

def run_build(input_file):
    if input_file is None:
        return None, "⚠️ Файл не загружен.", gr.update(visible=False, value=None)

    try:
        wb = ess.build_summary_table(input_file.name)

        # Сохраняем с понятным именем
        original_name = os.path.splitext(os.path.basename(input_file.name))[0]
        out_path = os.path.join(tempfile.gettempdir(), f"{original_name}_свод.xlsx")
        wb.save(out_path)

        return out_path, "✅ Готово! Нажмите кнопку ниже для скачивания.", gr.update(visible=True, value=out_path)
    except Exception as e:
        return None, f"❌ Ошибка: {e}", gr.update(visible=False, value=None)

with gr.Blocks(title="Свод КП", css="""
    .yellow-button {background-color: #FFD700 !important; color: black !important; font-weight: bold !important;}
    .input-section {border: 2px solid #4CAF50; border-radius: 10px; padding: 20px; background-color: #f0f8f0;}
    .output-section {border: 2px solid #2196F3; border-radius: 10px; padding: 20px; background-color: #f0f4ff;}
    .info-section {border: 2px solid #FF9800; border-radius: 10px; padding: 20px; background-color: #fff8e1;}
""") as demo:
    gr.Markdown("## 📊 Свод КП из выгрузки ЯЗакупок (YP)")
    
    with gr.Group(elem_classes="info-section"):
        gr.Markdown("""
        ### ⚠️ Важная информация
        
        1. **Загружайте Excel, выгруженный из ЯЗакупок БЕЗ изменений в нем**
        2. Программа переформатирует только выгруженный Excel. Если КП не попали в Excel, то их и не будет в своде
        3. Программа показывает цены в рублях, однако вы можете загружать в нее любую валюту. Необходимо будет вручную изменить валюту в Excel
        4. **Проверяйте наличие всех позиций в переформатированном своде**
        """)
    
    with gr.Group(elem_classes="input-section"):
        gr.Markdown("### 📥 Шаг 1: Загрузите файл")
        file_in = gr.File(label="Выберите Excel файл (.xlsx)", file_types=[".xlsx"])
        run_btn = gr.Button("▶️ Собрать свод", elem_classes="yellow-button", size="lg")
    
    status = gr.Textbox(label="Статус обработки", interactive=False)
    
    with gr.Group(elem_classes="output-section"):
        gr.Markdown("### 📤 Шаг 2: Скачайте результат")
        file_out = gr.File(label="Готовый файл", elem_id="file_out")
        download_btn = gr.DownloadButton("⬇️ Скачать результат", elem_classes="yellow-button", size="lg", visible=False)

    run_btn.click(
        run_build,
        inputs=file_in,
        outputs=[file_out, status, download_btn]
    )

if __name__ == "__main__":
    demo.launch()
