import os
import json
import re
import traceback
import logging
import requests
from flask import Flask, request, jsonify
from docxtpl import DocxTemplate
from pptx import Presentation
from pptx.util import Inches
from drive_utils import upload_to_drive

# Templates directory
TEMPLATES_DIR = os.path.join(os.path.dirname(__file__), "templates")
DOCX_TEMPLATE = os.path.join(TEMPLATES_DIR, "Market_Gap_Analysis_Template.docx")
PPTX_TEMPLATE = os.path.join(TEMPLATES_DIR, "Market_Gap_Analysis_Template.pptx")

app = Flask(__name__)
logging.basicConfig(level=logging.INFO)


def insert_chart_into_slide(slide, chart_path, left=Inches(1), top=Inches(2), width=Inches(4), height=Inches(3)):
    """Insert a PNG chart onto the given slide at the specified position."""
    slide.shapes.add_picture(chart_path, left, top, width, height)


def replace_placeholder(slide, key, text):
    """
    Replace all {{ key }} tokens (case-insensitive, optional spaces)
    in any text frame on the given slide.
    """
    pattern = re.compile(r"\{\{\s*" + re.escape(key) + r"\s*\}\}", flags=re.IGNORECASE)
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            if not para.runs:
                para.text = pattern.sub(text, para.text)
            else:
                for run in para.runs:
                    run.text = pattern.sub(text, run.text)


def to_direct_drive_url(url: str) -> str:
    """Convert a Google Drive share URL into a direct-download URL."""
    m = re.search(r"[?&]id=([^&]+)", url)
    if m:
        return f"https://drive.google.com/uc?export=download&id={m.group(1)}"
    m = re.search(r"/d/([^/]+)", url)
    if m:
        return f"https://drive.google.com/uc?export=download&id={m.group(1)}"
    return url


def download_chart(url: str, local_path: str) -> str:
    """Download a chart PNG from Google Drive into local_path."""
    direct_url = to_direct_drive_url(url)
    os.makedirs(os.path.dirname(local_path), exist_ok=True)
    resp = requests.get(direct_url)
    resp.raise_for_status()
    with open(local_path, "wb") as f:
        f.write(resp.content)
    return local_path


@app.route("/generate_market_reports", methods=["POST"])
def generate_market_reports_api():
    data = request.get_json(force=True)
    logging.info("📦 Incoming payload for market reports:\n%s", json.dumps(data, indent=2))

    session_id = data.get("session_id")
    if not session_id:
        return jsonify({"error": "Missing session_id"}), 400

    local_path = os.path.join("temp_sessions", session_id)
    os.makedirs(local_path, exist_ok=True)

    try:
        result = generate_market_reports(
            session_id=session_id,
            email=data.get("email", ""),
            folder_id=data.get("folder_id", ""),
            payload=data,
            local_path=local_path
        )
        return jsonify(result), 200

    except Exception as e:
        logging.error(f"🔥 Market Reports generation failed: {e}")
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


def generate_market_reports(session_id: str,
                            email: str,
                            folder_id: str,
                            payload: dict,
                            local_path: str) -> dict:
    # 1. Render DOCX
    doc = DocxTemplate(DOCX_TEMPLATE)
    context = payload.get("content", {}).copy()
    context["date"] = payload.get("date", "")
    context["organization_name"] = payload.get("organization_name", "")

    docx_path = os.path.join(local_path, f"market_gap_analysis_report_{session_id}.docx")
    doc.render(context)
    doc.save(docx_path)
    docx_url = upload_to_drive(docx_path, session_id, folder_id)

    # 2. Render PPTX
    pres = Presentation(PPTX_TEMPLATE)
    content = payload.get("content", {})
    charts = payload.get("charts", {})

    # Map zero-based slide index → content key
    content_slide_map = {
        1: "executive_summary",       # slide 2 in deck
        2: "current_state_overview",  # slide 3
        3: "hardware_gap_analysis",   # slide 4
        4: "software_gap_analysis",   # slide 5
        5: "market_benchmarking",     # slide 6 (if present)
    }
    for idx, key in content_slide_map.items():
        if idx < len(pres.slides):
            replace_placeholder(pres.slides[idx], key, content.get(key, ""))

    # Map zero-based slide index → chart key
    chart_slide_map = {
        3: "hardware_insights_tier",  # slide 4
        4: "software_insights_tier",  # slide 5
    }
    for slide_idx, chart_key in chart_slide_map.items():
        if slide_idx < len(pres.slides) and chart_key in charts:
            local_png = os.path.join(local_path, f"{chart_key}.png")
            png_path = download_chart(charts[chart_key], local_png)
            insert_chart_into_slide(
                pres.slides[slide_idx],
                png_path,
                left=Inches(1),
                top=Inches(2),
                width=Inches(4),
                height=Inches(3)
            )

    # 3. Save and upload PPTX
    pptx_path = os.path.join(local_path, f"market_gap_analysis_executive_report_{session_id}.pptx")
    pres.save(pptx_path)
    pptx_url = upload_to_drive(pptx_path, session_id, folder_id)

    # 4. Build result
    return {
        "session_id": session_id,
        "gpt_module": "gap_market",
        "status": "complete",
        "content": content,
        "charts": charts,
        "files": [
            {"file_name": fn}
            for fn in payload.get("appendices", [])
        ],
        "file_1_name": os.path.basename(docx_path),
        "file_1_url": docx_url,
        "file_2_name": os.path.basename(pptx_path),
        "file_2_url": pptx_url
    }


if __name__ == "__main__":
    port = int(os.getenv("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
