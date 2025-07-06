import os
import json
import re
import traceback
import logging
import requests
from flask import Flask, request, jsonify
from docxtpl import DocxTemplate
from pptx import Presentation
from pptx.util import Inches, Pt
from drive_utils import upload_to_drive

# Ensure staging directory exists
os.makedirs("temp_sessions", exist_ok=True)

# Templates directory
TEMPLATES_DIR = os.path.join(os.path.dirname(__file__), "templates")
DOCX_TEMPLATE = os.path.join(TEMPLATES_DIR, "Market_Gap_Analysis_Template.docx")
PPTX_TEMPLATE = os.path.join(TEMPLATES_DIR, "Market_Gap_Analysis_Template.pptx")

app = Flask(__name__)
logging.basicConfig(level=logging.INFO)


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
    Download a chart PNG from Drive into local_path.
    """
    direct_url = to_direct_drive_url(url)
    os.makedirs(os.path.dirname(local_path), exist_ok=True)
    resp = requests.get(direct_url)
    resp.raise_for_status()
    with open(local_path, "wb") as f:
        f.write(resp.content)
    logging.info(f"✅ Downloaded chart to %s", local_path)
    return local_path


def replace_placeholder(slide, key, text):
    """
    Replace any {{ key }} tags in text frames of the slide.
    """
    pattern = re.compile(r"\{\{\s*" + re.escape(key) + r"\s*\}\}")
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                if pattern.search(paragraph.text):
                    paragraph.text = pattern.sub(text, paragraph.text)
                    logging.info(f"✅ Replaced placeholder %s on slide", key)

@app.route("/generate_market_reports", methods=["POST"])
def start_market_gap():
    data = request.get_json(force=True)
    logging.info("📦 Incoming payload for market reports:\n%s", json.dumps(data, indent=2))

    session_id = data.get("session_id")
    if not session_id:
        return jsonify({"error": "Missing session_id"}), 400

    # Create local session folder
    local_path = os.path.join("temp_sessions", session_id)
    os.makedirs(local_path, exist_ok=True)

    try:
        result = generate_market_reports(
            session_id=session_id,
            email=data.get("organization_name", ""),
            folder_id=data.get("folder_id", ""),
            payload=data,
            local_path=local_path
        )
        if not result:
            logging.error("❌ generate_market_reports returned no result")
            return jsonify({"error": "Report generation failed"}), 500
        logging.info("✅ Completed market reports for %s", session_id)
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
    # Verify template files exist
    logging.info(f"Using DOCX template: {DOCX_TEMPLATE}")
    if not os.path.exists(DOCX_TEMPLATE):
        raise FileNotFoundError(f"DOCX template not found at {DOCX_TEMPLATE}")
    logging.info(f"Using PPTX template: {PPTX_TEMPLATE}")
    if not os.path.exists(PPTX_TEMPLATE):
        raise FileNotFoundError(f"PPTX template not found at {PPTX_TEMPLATE}")

    try:
        # 1. Generate DOCX
        logging.info("🔄 Starting DOCX generation...")(session_id: str,
                            email: str,
                            folder_id: str,
                            payload: dict,
                            local_path: str) -> dict:
    """
    Renders DOCX and PPTX using templates, uploads to Drive, and constructs result.
    """
    try:
        # 1. Generate DOCX
        doc = DocxTemplate(DOCX_TEMPLATE)
        context = payload.get("content", {}).copy()
        context["date"] = payload.get("date", "")
        context["organization_name"] = email

        docx_filename = f"market_gap_analysis_report_{session_id}.docx"
        docx_path = os.path.join(local_path, docx_filename)
        doc.render(context)
        doc.save(docx_path)
        logging.info("✅ DOCX saved to %s", docx_path)
        docx_url = upload_to_drive(docx_path, folder_id)
        logging.info("✅ DOCX uploaded to %s", docx_url)

        # 2. Generate PPTX
        pres = Presentation(PPTX_TEMPLATE)
        content = payload.get("content", {})
        charts = payload.get("charts", {})

        # Populate all placeholders on template slides
        for key, text in content.items():
            for slide in pres.slides:
                replace_placeholder(slide, key, text)

        # Insert charts for slides 1 and 2
        hw_url = charts.get("hardware_insights_tier")
        if hw_url:
            chart_local = os.path.join(local_path, "hardware_tier.png")
            chart_path = download_chart(hw_url, chart_local)
            pres.slides[1].shapes.add_picture(
                chart_path, Inches(1), Inches(1), width=Inches(8), height=Inches(4.5)
            )
            logging.info("✅ Hardware chart inserted on slide 1")

        sw_url = charts.get("software_insights_tier")
        if sw_url:
            chart_local = os.path.join(local_path, "software_tier.png")
            chart_path = download_chart(sw_url, chart_local)
            pres.slides[2].shapes.add_picture(
                chart_path, Inches(1), Inches(1), width=Inches(8), height=Inches(4.5)
            )
            logging.info("✅ Software chart inserted on slide 2")

        # Save PPTX
        pptx_filename = f"market_gap_analysis_executive_report_{session_id}.pptx"
        pptx_path = os.path.join(local_path, pptx_filename)
        pres.save(pptx_path)
        logging.info("✅ PPTX saved to %s", pptx_path)
        pptx_url = upload_to_drive(pptx_path, folder_id)
        logging.info("✅ PPTX uploaded to %s", pptx_url)

        # 3. Construct result
        result = {
            "session_id": session_id,
            "status": "complete",
            "content": content,
            "charts": charts,
            "report_urls": [docx_url, pptx_url],
            "file_1_name": os.path.basename(docx_path),
            "file_1_url": docx_url,
            "file_2_name": os.path.basename(pptx_path),
            "file_2_url": pptx_url
        }
        return result

    except Exception as e:
        logging.error(f"🔥 Error in generate_market_reports: {e}")
        traceback.print_exc()
        return None

if __name__ == "__main__":
    port = int(os.getenv("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
