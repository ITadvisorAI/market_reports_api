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



from pptx.util import Inches
from pptx.enum.shapes import MSO_SHAPE_TYPE

def insert_chart_into_slide(slide, chart_path, left=Inches(5.5), top=Inches(2), width=Inches(4), height=Inches(3)):
    slide.shapes.add_picture(chart_path, left, top, width, height)

def place_charts(presentation, chart_paths_dict):
    for slide in presentation.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                text = shape.text_frame.text.strip().lower()
                if "hardware tier distribution" in text and "hardware_insights_tier" in chart_paths_dict:
                    insert_chart_into_slide(slide, chart_paths_dict["hardware_insights_tier"])
                elif "software tier distribution" in text and "software_insights_tier" in chart_paths_dict:
                    insert_chart_into_slide(slide, chart_paths_dict["software_insights_tier"])
def to_direct_drive_url(url: str) -> str:
    """
    Convert a Google Drive share URL into a direct-download URL.
    """
    m = re.search(r"[?&]id=([^&]+)", url)
    if m:
        file_id = m.group(1)
        return f"https://drive.google.com/uc?export=download&id={file_id}"
    m = re.search(r"/d/([^/]+)", url)
    if m:
        file_id = m.group(1)
        return f"https://drive.google.com/uc?export=download&id={file_id}"
    return url


def download_chart(url: str, local_path: str) -> str:
    """
    Download a chart PNG from Drive into local_path as a file.
    """
    direct_url = to_direct_drive_url(url)
    os.makedirs(os.path.dirname(local_path), exist_ok=True)
    resp = requests.get(direct_url)
    resp.raise_for_status()
    with open(local_path, "wb") as f:
        f.write(resp.content)
    return local_path


def replace_placeholder(slide, key, text):
    """
    Replace {{ key }} placeholders across all shape types, including text boxes, placeholders, and content.
    """
    pattern = re.compile(r"\{\{\s*" + re.escape(key) + r"\s*\}\}")
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        text_frame = shape.text_frame
        for paragraph in text_frame.paragraphs:
            if not paragraph.runs:
                paragraph.text = pattern.sub(text, paragraph.text)
            for run in paragraph.runs:
                run.text = pattern.sub(text, run.text)

def start_market_gap():
    data = request.get_json(force=True)
    logging.info("📦 Incoming payload:\n%s", json.dumps(data, indent=2))

    session_id = data.get("session_id")
    if not session_id:
        return jsonify({"error": "Missing session_id"}), 400

    # Create local session folder
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
        next_webhook = data.get("next_action_webhook")
        if next_webhook and result:
            try:
                requests.post(next_webhook, json=result, timeout=30)
            except Exception:
                pass
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
    """
    Renders DOCX and PPTX using templates, uploads to Drive, and constructs result.
    """
    try:
        # 1. Generate DOCX report
        doc = DocxTemplate(DOCX_TEMPLATE)
        context = {}
        # Fill in all content sections
        context.update(payload.get("content", {}))
        # Add metadata
        context["date"] = payload.get("date", "")
        context["organization_name"] = payload.get("organization_name", "")

        docx_filename = f"market_gap_analysis_report_{session_id}.docx"
        docx_path = os.path.join(local_path, docx_filename)
        doc.render(context)
        doc.save(docx_path)
        docx_url = upload_to_drive(docx_path, session_id, folder_id)

        # 2. Generate PPTX executive report
        pres = Presentation(PPTX_TEMPLATE)
        content = payload.get("content", {})
        charts = payload.get("charts", {})

        # Slide 1: Executive Summary
        # Slide 2–6: Main content
        content_slide_map = {
            2: "executive_summary",
            3: "current_state_overview",
            4: "hardware_gap_analysis",
            5: "software_gap_analysis",
            6: "market_benchmarking"
        }

        for slide_index, key in content_slide_map.items():
            if len(pres.slides) > slide_index:
                replace_placeholder(pres.slides[slide_index], key, content.get(key, ""))

        # Optional: Add chart images to Slide 0 (Agenda)
        chart_slide_index = 0
        if len(pres.slides) > chart_slide_index:
            hw_url = charts.get("hardware_tier_distribution") or charts.get("hardware_insights_tier")
            sw_url = charts.get("software_tier_distribution") or charts.get("software_insights_tier")

            if hw_url:
                chart_local = os.path.join(local_path, "hardware_tier.png")
                chart_path = download_chart(hw_url, chart_local)
                if os.path.exists(chart_path):
                    pres.slides[chart_slide_index].shapes.add_picture(
                        chart_path,
                        Inches(0.5), Inches(1.8),
                        width=Inches(4), height=Inches(3)
                    )
            if sw_url:
                chart_local = os.path.join(local_path, "software_tier.png")
                chart_path = download_chart(sw_url, chart_local)
                if os.path.exists(chart_path):
                    pres.slides[chart_slide_index].shapes.add_picture(
                        chart_path,
                        Inches(5), Inches(1.8),
                        width=Inches(4), height=Inches(3)
                    )

        # Save PPTX
        pptx_filename = f"market_gap_analysis_executive_report_{session_id}.pptx"
        pptx_path = os.path.join(local_path, pptx_filename)
        pres.save(pptx_path)
        pptx_url = upload_to_drive(pptx_path, session_id, folder_id)

        # 3. Construct result payload
        result = {
            "session_id": session_id,
            "gpt_module": "gap_market",
            "status": "complete",
            "content": content,
            "charts": charts,
            "files": [
                {"file_name": f["file_name"], "file_url": f["file_url"]}
                for f in payload.get("files", [])
            ],
            "file_1_name": os.path.basename(docx_path),
            "file_1_url": docx_url,
            "file_2_name": os.path.basename(pptx_path),
            "file_2_url": pptx_url
        }

        # 4. Downstream callback
        next_webhook = payload.get("next_action_webhook") or (
            os.getenv("IT_STRATEGY_API_URL", "") + "/start_it_strategy"
        )
        if next_webhook:
            try:
                requests.post(next_webhook, json=result, timeout=30)
            except Exception:
                pass

        return result

    except Exception as e:
        logging.error(f"🔥 Market Reports generation failed: {e}")
        traceback.print_exc()
        return None


if __name__ == "__main__":
    port = int(os.getenv("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
