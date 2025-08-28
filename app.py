# app.py
import os
import tempfile
import gradio as gr
from excel_summary_script import build_summary_table

# ---------- utils ----------
def _path_from(file_obj):
    if not file_obj:
        return None
    # UploadButton может вернуть объект или dict/список — обработаем всё
    if isinstance(file_obj, (list, tuple)) and file_obj:
        file_obj = file_obj[0]
    if isinstance(file_obj, dict):
        return file_obj.get("path") or file_obj.get("name")
    return getattr(file_obj, "path", None) or getattr(file_obj, "name", None)

# ---------- handlers ----------
def keep_file(file_obj):
    """Сохраняем выбранный файл в state и показываем имя."""
    path = _path_from(file_obj)
    if not path:
        return None, "Файл не выбран"
    name = os.path.basename(path)
    return file_obj, f"Файл: **{name}**"

def make_summary(file_obj):
    """Формируем свод и даём скачать результат."""
    path = _path_from(file_obj)
    if not path:
        return gr.update(visible=False), "⚠️ Сначала выберите .xlsx"

    try:
        wb = build_summary_table(path)
        fd, out_path = tempfile.mkstemp(suffix=".xlsx")
        os.close(fd)
        wb.save(out_path)
        # передаём путь в DownloadButton
        return gr.update(visible=True, value=out_path), "✅ Готово! Можно скачивать."
    except Exception as e:
        return gr.update(visible=False), f"❌ Ошибка: {e}"

# ---------- UI ----------
CSS = """
.gradio-container { max-width: 820px !important; margin: 0 auto !important; }
#wrap { gap: 12px; }
"""

with gr.Blocks(css=CSS, title="Свод КП") as demo:
    gr.Markdown("## 📊 Свод КП\nШаг 1 — **выберите Excel (.xlsx)**. Шаг 2 — **нажмите «Сформировать свод»**.")

    file_state = gr.State(None)

    with gr.Row(elem_id="wrap"):
        choose_btn = gr.UploadButton("📁 Выбрать файл (.xlsx)", file_types=[".xlsx"], file_count="single")
        run_btn = gr.Button("🚀 Сформировать свод", variant="primary")

    file_info = gr.Markdown("Файл не выбран")
    download_btn = gr.DownloadButton("⬇️ Скачать результат", visible=False)
    status = gr.Textbox(label="Статус", interactive=False, lines=2)

    # 1) сохраняем выбранный файл и показываем имя
    choose_btn.upload(fn=keep_file, inputs=choose_btn, outputs=[file_state, file_info])
    # 2) формируем свод
    run_btn.click(fn=make_summary, inputs=file_state, outputs=[download_btn, status])

if __name__ == "__main__":
    demo.launch()
