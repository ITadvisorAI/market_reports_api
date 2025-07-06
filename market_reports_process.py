
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
    doc = DocxTemplate(DOCX_TEMPLATE)
    context = payload.get("content", {})
    context["date"] = payload.get("date", "")
    context["organization_name"] = payload.get("organization_name", "")

    docx_path = os.path.join(local_path, f"market_gap_analysis_report_{session_id}.docx")
    doc.render(context)
    doc.save(docx_path)
    docx_url = upload_to_drive(docx_path, session_id, folder_id)

    pres = Presentation(PPTX_TEMPLATE)
    content = payload.get("content", {})
    charts = payload.get("charts", {})

    if len(pres.slides) > 1:
        try:
            pres.slides[1].placeholders[1].text = content.get("executive_summary", "")
        except Exception:
            pass

    slide_text_map = {
        4: "current_state_overview",
        5: "hardware_gap_analysis",
        6: "software_gap_analysis",
        7: "market_benchmarking"
    }
    for slide_index, key in slide_text_map.items():
        try:
            if len(pres.slides) > slide_index:
                pres.slides[slide_index].placeholders[1].text = content.get(key, "")
        except Exception:
            pass

    chart_slide_index = 0
    if len(pres.slides) > chart_slide_index:
        hw_url = charts.get("hardware_tier_distribution") or charts.get("hardware_insights_tier")
        sw_url = charts.get("software_tier_distribution") or charts.get("software_insights_tier")

        if hw_url:
            chart_local = os.path.join(local_path, "hardware_tier.png")
            chart_path = download_chart(hw_url, chart_local)
            if os.path.exists(chart_path):
                pres.slides[chart_slide_index].shapes.add_picture(chart_path, Inches(0.5), Inches(1.8), width=Inches(4), height=Inches(3))
        if sw_url:
            chart_local = os.path.join(local_path, "software_tier.png")
            chart_path = download_chart(sw_url, chart_local)
            if os.path.exists(chart_path):
                pres.slides[chart_slide_index].shapes.add_picture(chart_path, Inches(5), Inches(1.8), width=Inches(4), height=Inches(3))

    pptx_path = os.path.join(local_path, f"market_gap_analysis_executive_report_{session_id}.pptx")
    pres.save(pptx_path)
    pptx_url = upload_to_drive(pptx_path, session_id, folder_id)

    return {
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


if __name__ == "__main__":
    port = int(os.getenv("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
