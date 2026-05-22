import os

BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
DATA_FOLDER = os.path.join(BASE_DIR, "data")   # Excel tracker files live in /data subfolder

TRACKER_FILES = {
    "Niyas":                "Niyas Tracker 2026 Format v2.xlsx",
    "Charlotte":            "Tracker 2026 - Charlotte_updated.xlsx",
    "Amit Gandhi":          "Tracker 2026 Format v2 (Amit Gandhi).xlsx",
    "Mitesh Sheth":         "Tracker 2026 Format v23Jan2026 (Mitesh Sheth)_updated.xlsx",
    "Mohammed Azharudheen": "Tracker_2026 Format v2_Azhar.xlsx",
}

# ── OneDrive / Microsoft Graph API config ───────────────────────────────────
# Set these as environment variables in Vercel (Project → Settings → Environment Variables)
# Leave blank to fall back to reading local Excel files bundled in the repo.
ONEDRIVE_CFG = {
    "tenant_id":     os.environ.get("AZURE_TENANT_ID", ""),
    "client_id":     os.environ.get("AZURE_CLIENT_ID", ""),
    "client_secret": os.environ.get("AZURE_CLIENT_SECRET", ""),
    "user_email":    os.environ.get("ONEDRIVE_USER_EMAIL", ""),   # e.g. amit@alphadat.ae
    "folder_path":   os.environ.get("ONEDRIVE_FOLDER_PATH", ""), # e.g. Documents/Amit Tracker/Dashboard
}

USERS = {
    "admin": "AlphaData@2026",
    "mgmt":  "Review@2026",
}

SECRET_KEY = "aD!x9Kp2mZ$vQ7nL"
HOST  = "0.0.0.0"
PORT  = 5000
DEBUG = False
