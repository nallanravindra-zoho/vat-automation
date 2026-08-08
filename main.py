"""
Webhook behind the "Run VAT Return" Zoho Books widget, deployed on Cloud Run.
Pulls Invoice/Bill/Expense/Credit Note data directly from the Zoho Books API
for the requested period (see zoho_books_api.py), builds the return, and
writes the result to OneDrive for Business via Microsoft Graph (see
onedrive_utils.py for the auth setup that requires in Azure AD).

Flow:
  widget (browser) --POST {year, month}--> Cloud Run --pulls the period's
  data live from Zoho Books, runs the pipeline, uploads results to
  OneDrive--> returns {folder_path, files, net_vat_due} --> widget renders

Environment variables (set at deploy time — see deploy.sh):
  ZOHO_CLIENT_ID / ZOHO_CLIENT_SECRET / ZOHO_REFRESH_TOKEN / ZOHO_ORG_ID
  ZOHO_API_DOMAIN / ZOHO_ACCOUNTS_DOMAIN   (data-center specific — see
                                            zoho_books_api.py)
  AZURE_TENANT_ID / AZURE_CLIENT_ID / AZURE_CLIENT_SECRET
                         Azure AD app registration (AZURE_CLIENT_SECRET via
                         Secret Manager)
  ONEDRIVE_DRIVE_ID      OR ONEDRIVE_USER_UPN (email whose OneDrive to use)
  ONEDRIVE_OUTPUT_ROOT   e.g. "/VAT Automation"
                         — the "VAT - <Month>'<YY>" folder is created under here
  ALLOWED_ORIGINS        comma-separated, e.g. https://books.zoho.com
  WEBHOOK_SHARED_SECRET  see TODO below
"""
import os
import json
import logging
import tempfile
import shutil
from datetime import datetime, timezone

from flask import Flask, request, jsonify
from flask_cors import CORS

import vat_pipeline
import onedrive_utils

app = Flask(__name__)

ALLOWED_ORIGINS = os.environ.get("ALLOWED_ORIGINS", "https://books.zoho.com").split(",")
CORS(app, origins=ALLOWED_ORIGINS)

ONEDRIVE_OUTPUT_ROOT = os.environ.get("ONEDRIVE_OUTPUT_ROOT", "/VAT Automation")

# TODO: set WEBHOOK_SHARED_SECRET as a Cloud Run env var (ideally sourced from
# Secret Manager) once you flip the service to --allow-unauthenticated, and
# set the same value as a header in widgets/index.html's fetch() call. CORS
# origin restriction alone isn't a real access control — anyone can spoof an
# Origin header from outside a browser context.
WEBHOOK_SHARED_SECRET = os.environ.get("WEBHOOK_SHARED_SECRET")

logging.basicConfig(level=logging.INFO)
log = logging.getLogger("vat-webhook")


@app.route("/healthz", methods=["GET"])
def healthz():
    return jsonify({"status": "ok"})


@app.route("/run-vat-return", methods=["POST"])
def run_vat_return():
    if WEBHOOK_SHARED_SECRET and request.headers.get("X-Webhook-Secret") != WEBHOOK_SHARED_SECRET:
        return jsonify({"error": "unauthorized"}), 401

    payload = request.get_json(force=True)
    org_id = payload.get("organization_id")
    year = payload.get("year")
    month = payload.get("month")

    if not year or not month:
        return jsonify({"error": "year and month are required"}), 400

    run_started = datetime.now(timezone.utc).isoformat()
    log.info(f"Run requested: org={org_id} period={year}-{month}")

    work_root = tempfile.mkdtemp(prefix="vat_run_")
    output_root = os.path.join(work_root, "output")

    try:
        drive_id = onedrive_utils.resolve_drive_id()

        result = vat_pipeline.run(
            output_root=output_root,
            year=year,
            month=month,
            org_id=org_id,
        )

        local_folder_name = os.path.basename(result["folder_path"])
        dest_folder_path = f"{ONEDRIVE_OUTPUT_ROOT}/{local_folder_name}"
        uploaded = onedrive_utils.upload_folder(drive_id, result["folder_path"], dest_folder_path)

    except FileNotFoundError as e:
        log.error(f"Source data missing: {e}")
        return jsonify({"error": str(e)}), 422
    except Exception as e:
        log.exception("Pipeline failed")
        return jsonify({"error": str(e)}), 500
    finally:
        shutil.rmtree(work_root, ignore_errors=True)

    # Audit trail alongside the output — auditors requesting samples get this too
    log_entry = {
        "run_started_utc": run_started,
        "run_completed_utc": datetime.now(timezone.utc).isoformat(),
        "organization_id": org_id,
        "period": f"{year}-{month}",
        "onedrive_folder": dest_folder_path,
        "files": uploaded,
        "net_vat_due": result["net_vat_due"],
        "reconciliation_mismatches": result.get("reconciliation_mismatches"),
    }
    try:
        onedrive_utils.upload_json(
            drive_id, dest_folder_path,
            f"_run_log_{run_started.replace(':', '-')}.json",
            json.dumps(log_entry, indent=2, default=str).encode("utf-8"),
        )
    except Exception:
        log.exception("Failed to write run log (non-fatal, continuing)")

    return jsonify({
        "folder_path": dest_folder_path,
        "files": [os.path.basename(u) for u in uploaded],
        "net_vat_due": result["net_vat_due"],
        "reconciliation_mismatches": result.get("reconciliation_mismatches"),
    })


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8080)))
