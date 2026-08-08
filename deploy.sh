#!/usr/bin/env bash
# Deploys the "Run VAT Return" webhook to Cloud Run.
# Source data is pulled LIVE from the Zoho Books API on each run — no synced
# folder needed. Output goes to OneDrive for Business via Microsoft Graph.
#
# Prerequisites (one-time):
#   Zoho:
#     - A self-client OAuth app (https://api-console.zoho.com -> Self Client)
#       with scope ZohoBooks.fullaccess.READ, generate a refresh token
#     - Client ID, Client Secret, Refresh Token, and your Organization ID
#       (Zoho Books -> Settings -> Organization Profile)
#     - Confirm your data center: books.zoho.sa -> zohoapis.sa / accounts.zoho.sa,
#       books.zoho.com -> zohoapis.com / accounts.zoho.com
#   Azure AD (for OneDrive output — see onedrive_utils.py header for full steps):
#     - App registration with Files.ReadWrite.All (Application), admin-consented
#     - Tenant ID, Client ID, Client Secret
#     - Target OneDrive's drive ID, or the user's UPN (email)
#
# Usage:
#   1. Edit the variables below (or export them before running).
#   2. gcloud auth login && gcloud config set project <your-project-id>
#   3. ./deploy.sh
#
# Re-running is safe — every gcloud command here is idempotent.

set -euo pipefail

# ---- Configuration — edit these -------------------------------------------
PROJECT_ID="${PROJECT_ID:-your-gcp-project-id}"
REGION="${REGION:-me-central2}"                       # Doha — closest GCP region to KSA
SERVICE_NAME="${SERVICE_NAME:-vat-return-webhook}"
ALLOWED_ORIGINS="${ALLOWED_ORIGINS:-https://books.zoho.com,https://books.zoho.sa}"

ZOHO_CLIENT_ID="${ZOHO_CLIENT_ID:-}"
ZOHO_CLIENT_SECRET="${ZOHO_CLIENT_SECRET:-}"            # goes into Secret Manager
ZOHO_REFRESH_TOKEN="${ZOHO_REFRESH_TOKEN:-}"             # goes into Secret Manager
ZOHO_ORG_ID="${ZOHO_ORG_ID:-}"
ZOHO_API_DOMAIN="${ZOHO_API_DOMAIN:-https://www.zohoapis.sa}"
ZOHO_ACCOUNTS_DOMAIN="${ZOHO_ACCOUNTS_DOMAIN:-https://accounts.zoho.sa}"

AZURE_TENANT_ID="${AZURE_TENANT_ID:-}"
AZURE_CLIENT_ID="${AZURE_CLIENT_ID:-}"
AZURE_CLIENT_SECRET="${AZURE_CLIENT_SECRET:-}"          # goes into Secret Manager
ONEDRIVE_USER_UPN="${ONEDRIVE_USER_UPN:-}"               # e.g. finance@knightstelecom.com — leave blank if using ONEDRIVE_DRIVE_ID
ONEDRIVE_DRIVE_ID="${ONEDRIVE_DRIVE_ID:-}"               # alternative to ONEDRIVE_USER_UPN
ONEDRIVE_OUTPUT_ROOT="${ONEDRIVE_OUTPUT_ROOT:-/VAT Automation}"
WEBHOOK_SHARED_SECRET="${WEBHOOK_SHARED_SECRET:-}"       # also goes into widgets/index.html
# ----------------------------------------------------------------------------

for var in ZOHO_CLIENT_ID ZOHO_CLIENT_SECRET ZOHO_REFRESH_TOKEN ZOHO_ORG_ID \
           AZURE_TENANT_ID AZURE_CLIENT_ID AZURE_CLIENT_SECRET; do
  if [ -z "${!var}" ]; then
    echo "ERROR: $var must be set (export it, or edit the top of this script)."
    exit 1
  fi
done
if [ -z "$ONEDRIVE_USER_UPN" ] && [ -z "$ONEDRIVE_DRIVE_ID" ]; then
  echo "ERROR: set either ONEDRIVE_USER_UPN or ONEDRIVE_DRIVE_ID."
  exit 1
fi

echo "Project:   $PROJECT_ID"
echo "Region:    $REGION"
echo "Service:   $SERVICE_NAME"
echo

gcloud config set project "$PROJECT_ID"

echo "== Enabling required APIs =="
gcloud services enable \
  run.googleapis.com \
  cloudbuild.googleapis.com \
  artifactregistry.googleapis.com \
  secretmanager.googleapis.com

put_secret() {
  local name="$1" value="$2"
  if gcloud secrets describe "$name" >/dev/null 2>&1; then
    echo -n "$value" | gcloud secrets versions add "$name" --data-file=-
  else
    echo -n "$value" | gcloud secrets create "$name" --data-file=- --replication-policy=automatic
  fi
}

echo "== Storing secrets in Secret Manager =="
put_secret "vat-webhook-zoho-client-secret" "$ZOHO_CLIENT_SECRET"
put_secret "vat-webhook-zoho-refresh-token" "$ZOHO_REFRESH_TOKEN"
put_secret "vat-webhook-azure-client-secret" "$AZURE_CLIENT_SECRET"

SECRETS_ARG="ZOHO_CLIENT_SECRET=vat-webhook-zoho-client-secret:latest,ZOHO_REFRESH_TOKEN=vat-webhook-zoho-refresh-token:latest,AZURE_CLIENT_SECRET=vat-webhook-azure-client-secret:latest"
if [ -n "$WEBHOOK_SHARED_SECRET" ]; then
  put_secret "vat-webhook-shared-secret" "$WEBHOOK_SHARED_SECRET"
  SECRETS_ARG="${SECRETS_ARG},WEBHOOK_SHARED_SECRET=vat-webhook-shared-secret:latest"
fi

echo "== Building the container image via Cloud Build =="
IMAGE="gcr.io/${PROJECT_ID}/${SERVICE_NAME}:latest"
gcloud builds submit --tag "$IMAGE" .

echo "== Deploying to Cloud Run =="
ONEDRIVE_TARGET_ARG="ONEDRIVE_DRIVE_ID=${ONEDRIVE_DRIVE_ID}"
[ -n "$ONEDRIVE_USER_UPN" ] && ONEDRIVE_TARGET_ARG="ONEDRIVE_USER_UPN=${ONEDRIVE_USER_UPN}"

ENV_VARS="ZOHO_CLIENT_ID=${ZOHO_CLIENT_ID},ZOHO_ORG_ID=${ZOHO_ORG_ID},ZOHO_API_DOMAIN=${ZOHO_API_DOMAIN},ZOHO_ACCOUNTS_DOMAIN=${ZOHO_ACCOUNTS_DOMAIN}"
ENV_VARS="${ENV_VARS},AZURE_TENANT_ID=${AZURE_TENANT_ID},AZURE_CLIENT_ID=${AZURE_CLIENT_ID},${ONEDRIVE_TARGET_ARG}"
ENV_VARS="${ENV_VARS},ONEDRIVE_OUTPUT_ROOT=${ONEDRIVE_OUTPUT_ROOT},ALLOWED_ORIGINS=${ALLOWED_ORIGINS}"

gcloud run deploy "$SERVICE_NAME" \
  --image "$IMAGE" \
  --region "$REGION" \
  --platform managed \
  --allow-unauthenticated \
  --memory 1Gi \
  --cpu 1 \
  --timeout 300 \
  --max-instances 5 \
  --set-env-vars "$ENV_VARS" \
  --set-secrets "$SECRETS_ARG"

SERVICE_URL=$(gcloud run services describe "$SERVICE_NAME" --region "$REGION" --format='value(status.url)')
echo
echo "== Done =="
echo "Webhook URL: ${SERVICE_URL}/run-vat-return"
echo
echo "Next steps:"
echo "  1. Paste that URL into WEBHOOK_URL in widgets/index.html"
echo "  2. If you set WEBHOOK_SHARED_SECRET above, put the same value into"
echo "     WEBHOOK_SHARED_SECRET in widgets/index.html (sent as the"
echo "     X-Webhook-Secret header). This service is publicly reachable"
echo "     (--allow-unauthenticated) since a browser widget can't attach a"
echo "     Google identity token — the shared secret is the real gate here,"
echo "     don't skip setting it."
echo "  3. Before relying on this: call each get_*_for_period() function in"
echo "     zoho_books_api.py once against your real org and check the field"
echo "     names in the response match the ones marked 'verify' in"
echo "     vat_pipeline.py's extract_* functions — Zoho's schema has small"
echo "     regional differences and this hasn't been tested against a live"
echo "     Knights Telecom organization."
echo "  4. Confirm the Azure AD app has Files.ReadWrite.All (Application)"
echo "     granted admin consent, and the Zoho self-client has"
echo "     ZohoBooks.fullaccess.READ — without those, every call 401s."
