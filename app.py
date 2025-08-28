# app.py
import os
import tempfile
import gradio as gr
from excel_summary_script import build_summary_table

# ---------- helpers ----------
def _extract_path(file_obj):
    """Достаём путь к временному файлу из UploadButton/State (учитываем разные форматы Gradio v4)."""
    if not file_obj:
        return None
    # Если вдруг пришёл список — берём первый
    if isinstance(file_obj, (list, tuple)) and file_obj:
        file_obj = file_obj[0]
    # dict {"name": "...", "path": "..."} или объект с .path / .name
    if isinstance(file_obj, dict):
        return file_obj.get("path") or file_obj.get("name")
    return getattr(file_obj, "path", None) or getattr(file_obj, "name", None)

def _display_name(file_obj):
    p = _extract_path(file_obj)
    if not p:
        return "Файл не выбран"
    return f"Файл: {os.path.basename(p)}"

# ---------- handlers ----------
def keep_file(file_obj):
    """Сохраняем выбранный файл в state и показываем имя."""
    return file_obj, _display_name(file_obj)

def make_summary(file_obj):
    """Формируем свод и возвращаем кнопку скачивания + статус."""
    path = _extract_path(file_obj)
    if not path:
        return gr.update(visible=False, value=None), "⚠️ Сначала выберите .xlsx"

    try:
        wb = build_summary_table(path)
        fd, out_path = tempfile.mkstemp(suffix=".xlsx")
        os.close(fd)
        wb.save(out_path)
        return gr.update(value=out_path, visible=True), "✅ Готово! Можно скачивать."
    except Exception as e:
        return gr.update(visible=False, value=None), f"❌ Ошибка: {e}"

# ---------- UI ----------
CSS = """
.gradio-container { max-width: 820px !important; margin: 0 auto !important; }
"""

with gr.Blocks(css=CSS, title="Свод КП") as demo:
    gr.Markdown("## 📊 Свод КП\nШаг 1 — **выберите Excel (.xlsx)**. Шаг 2 — **нажмите «Сформировать свод»**.")

    file_state = gr.State(None)

    with gr.Row():
        choose_btn = gr.UploadButton("📁 Выбрать файл (.xlsx)", file_types=[".xlsx"], file_count="single")
        run_btn = gr.Button("🚀 Сформировать свод", variant="primary")

    file_info = gr.Markdown("Файл не выбран", elem_id="fileinfo")
    download_btn = gr.DownloadButton("Скачать результат", visible=False)
    status = gr.Textbox(label="Статус", interactive=False, lines=2)

    # 1) сохраняем выбранный файл в state + показываем имя
    choose_btn.upload(fn=keep_file, inputs=choose_btn, outputs=[file_state, file_info])
    # 2) по клику формируем свод и выдаём кнопку скачивания
    run_btn.click(fn=make_summary, inputs=file_state, outputs=[download_btn, status])

if __name__ == "__main__":
    demo.launch()
