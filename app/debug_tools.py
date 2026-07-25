import os
import re
from datetime import datetime
from app.config import SCREENSHOT_DIR, HTML_DIR, TEXT_DIR
from app.logger import get_logger

log = get_logger("debug_tools")

def slugify(value: str) -> str:
    value = str(value).strip().lower()
    value = re.sub(r"[^a-zA-Z0-9]+", "_", value)
    return value.strip("_") or "step"

def timestamp():
    return datetime.now().strftime("%Y%m%d_%H%M%S")

def save_screenshot(page, step_name: str):
    name = f"{timestamp()}_{slugify(step_name)}.png"
    path = os.path.join(SCREENSHOT_DIR, name)
    page.screenshot(path=path, full_page=True)
    log.info(f"Saved screenshot: {path}")
    return path

def save_html(page, step_name: str):
    name = f"{timestamp()}_{slugify(step_name)}.html"
    path = os.path.join(HTML_DIR, name)
    html = page.content()
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
    log.info(f"Saved HTML: {path}")
    return path

def save_visible_text(page, step_name: str):
    name = f"{timestamp()}_{slugify(step_name)}.txt"
    path = os.path.join(TEXT_DIR, name)
    try:
        txt = page.locator("body").inner_text()
    except Exception:
        txt = ""
    with open(path, "w", encoding="utf-8") as f:
        f.write(txt)
    log.info(f"Saved visible text: {path}")
    return path

def save_all(page, step_name: str):
    shot = save_screenshot(page, step_name)
    html = save_html(page, step_name)
    text = save_visible_text(page, step_name)
    return {"screenshot": shot, "html": html, "text": text}

def save_all_frames(page, step_name: str):
    results = []
    frames = page.frames
    for idx, frame in enumerate(frames):
        try:
            frame_html = frame.content()
        except Exception:
            frame_html = ""
        try:
            frame_text = frame.locator("body").inner_text()
        except Exception:
            frame_text = ""

        html_name = f"{timestamp()}_{slugify(step_name)}_frame_{idx}.html"
        txt_name = f"{timestamp()}_{slugify(step_name)}_frame_{idx}.txt"

        html_path = os.path.join(HTML_DIR, html_name)
        txt_path = os.path.join(TEXT_DIR, txt_name)

        with open(html_path, "w", encoding="utf-8") as f:
            f.write(frame_html)

        with open(txt_path, "w", encoding="utf-8") as f:
            f.write(frame_text)

        results.append({
            "frame_index": idx,
            "html": html_path,
            "text": txt_path
        })

    log.info(f"Saved {len(results)} frame dumps for step: {step_name}")
    return results