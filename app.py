# app.py
import os
import sys
import tempfile
import traceback
import gradio as gr
from excel_summary_script import build_summary_table

# ---------- перенаправляем stdout/stderr в файл ----------
sys.stdout = open("log.txt", "w")
sys.stderr = sys.stdout


# ---------- helpers ----------
def _path_from(file_obj):
    """Достаём путь к временному файлу (Gradio v4 может отдавать объект/dict/список)."""
    if not file_obj:
        return None
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
        print("[warn] файл не выбран")
        return None, "Файл не выбран"
    print(f"[ok] выбран файл: {path}")
    return file_obj, f"Файл: **{os.path.basename(path)}**"

def make_summary(file_obj):
    """Формируем свод и возвращаем (кнопка скачивания, статус)."""
    path = _path_from(file_obj)
    print(f"[info] запускаем свод для: {path}")

    if not path:
        return gr.update(visible=False), "⚠️ Сначала выберите .xlsx"

    try:
        wb = build_summary_table(path)
        fd, out_path = tempfile.mkstemp(suffix=".xlsx")
        os.close(fd)
        wb.save(out_path)
        print(f"[ok] итоговый файл сохранён: {out_path}")
        return gr.update(visible=True, value=out_path), "✅ Готово! Скачайте результат."
    except Exception as e:
        tb = traceback.format_exc()
        print(f"[ERR] {tb}")
        return gr.update(visible=False), f"❌ Ошибка: {e}"


# ---------- UI ----------
CSS = """
.gradio-container { max-width: 820px !important; margin: 0 auto !important; }
#controls { gap: 12px; }
"""

with gr.Blocks(css=CSS, title="Свод КП") as demo:
    gr.Markdown("## 📊 Свод КП\nШаг 1 — **выберите Excel (.xlsx)**. Шаг 2 — **нажмите «Сформировать свод»**.")

    file_state = gr.State(None)

    # две кнопки
    with gr.Row(elem_id="controls"):
        choose_btn = gr.UploadButton("📁 Выбрать файл (.xlsx)", file_types=[".xlsx"], file_count="single")
        run_btn = gr.Button("🚀 Сформировать свод", variant="primary")

    file_info = gr.Markdown("Файл не выбран")
    download_btn = gr.DownloadButton("⬇️ Скачать результат", visible=False)
    status = gr.Textbox(label="Статус", lines=2, interactive=False)

    choose_btn.upload(fn=keep_file, inputs=choose_btn, outputs=[file_state, file_info])
    run_btn.click(fn=make_summary, inputs=file_state, outputs=[download_btn, status])


if __name__ == "__main__":
    demo.launch()
