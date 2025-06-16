import os
import json
import logging
import threading
import re
from flask import Flask, request, jsonify
from market_reports_process import generate_market_reports
from drive_utils import upload_to_drive

app = Flask(__name__)
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

BASE_DIR = "temp_sessions"
os.makedirs(BASE_DIR, exist_ok=True)

@app.route("/", methods=["GET"])
def health():
    return "✅ Market Reports API is live", 200

@app.route("/generate_market_reports", methods=["POST"])
def generate_reports():
    try:
        data = request.get_json(force=True)
        session_id = data.get("session_id")
        email = data.get("email", "")
        folder_id = data.get("folder_id")  # Reuse existing Drive folder

        logging.info("📦 Incoming payload for market reports:\n%s", json.dumps(data, indent=2))

        # Validate required fields
        missing = []
        if not session_id:
            missing.append("session_id")
        if not data.get("content"):
            missing.append("content")
        if not data.get("charts"):
            missing.append("charts")
        if missing:
            return jsonify({"error": f"Missing required fields: {', '.join(missing)}"}), 400

        # Prepare local session folder
        folder_name = session_id if session_id.startswith("Temp_") else f"Temp_{session_id}"
        local_path = os.path.join(BASE_DIR, folder_name)
        os.makedirs(local_path, exist_ok=True)

        # Background processing
        def runner():
            try:
                generate_market_reports(session_id, email, folder_id, data, local_path)
            except Exception:
                logging.exception("🔥 Error generating market reports")

        threading.Thread(target=runner, daemon=True).start()
        return jsonify({"message": "Market reports generation started"}), 200

    except Exception:
        logging.exception("🔥 Failed to initiate market reports generation")
        return jsonify({"error": "Internal server error"}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    logging.info(f"🚦 Starting Market Reports API on port {port}")
    app.run(host="0.0.0.0", port=port)
