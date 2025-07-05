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
    Replace ANY {{ key }} tag in all text frames of the slide.
    """
    pattern = re.compile(r"\{\{\s*" + re.escape(key) + r"\s*\}\}")
    for shape in slide.shapes:
        if shape.has_text_frame:
            new_txt = pattern.sub(text, shape.text)
            if new_txt != shape.text:
                shape.text = new_txt


@app.route("/start_market_gap", methods=["POST"])
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

        # Slide 0: Executive Summary
        replace_placeholder(pres.slides[0],
                            "executive_summary",
                            content.get("executive_summary", ""))

        # Slide 1: Hardware Tier Distribution Chart
        slide = pres.slides[1]
        hw_url = charts.get("hardware_tier_distribution") or charts.get("hardware_insights_tier")
        if hw_url:
            chart_local = os.path.join(local_path, "hardware_tier.png")
            chart_path = download_chart(hw_url, chart_local)
            slide.shapes.add_picture(
                chart_path,
                Inches(1), Inches(1),
                width=Inches(8), height=Inches(4.5)
            )

        # Slide 2: Software Tier Distribution Chart
        slide = pres.slides[2]
        sw_url = charts.get("software_tier_distribution") or charts.get("software_insights_tier")
        if sw_url:
            chart_local = os.path.join(local_path, "software_tier.png")
            chart_path = download_chart(sw_url, chart_local)
            slide.shapes.add_picture(
                chart_path,
                Inches(1), Inches(1),
                width=Inches(8), height=Inches(4.5)
            )

        # Slide 3: Current State Overview
        slide = prs.slides.add_slide(layout)
        # … add title/chart …
        textbox = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = textbox.text_frame
        for line in payload['content']['current_state_overview'].split("\n"):
        tf.add_paragraph().text = line

        # Slide 4: Hardware Gap Analysis
        slide = prs.slides.add_slide(layout)
        # … add title/chart …
        textbox = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = textbox.text_frame
        for line in payload['content']['hardware_gap_analysis'].split("\n"):
        tf.add_paragraph().text = line

        # Slide 5: Software Gap Analysis
        slide = prs.slides.add_slide(layout)
        # … add title/chart …
        textbox = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = textbox.text_frame
        for line in payload['content']['software_gap_analysis'].split("\n"):
        tf.add_paragraph().text = line

        # Slide 6: Market Benchmarking
        slide = prs.slides.add_slide(layout)
        # … add title/chart …
        textbox = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = textbox.text_frame
        for line in payload['content']['market_benchmarking'].split("\n"):
        tf.add_paragraph().text = line

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
