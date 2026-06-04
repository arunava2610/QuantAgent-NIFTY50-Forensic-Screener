import os
import logging
import requests
from dotenv import load_dotenv

load_dotenv()

# Centralized Logger
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger("SystemAgent")

# ---------------------------------------------------------
# CENTRALIZED API KEYS & GOOGLE DRIVE IDs
# ---------------------------------------------------------
API_KEY_DRIVE_ID = "1eaOBvwclf6xtIxfT3RnJdE3dwCKm1b_5"
EXCEL_DRIVE_ID = "1WQljak5wURGcg5izR_jPFfMwcqza97jx"
ERP_DRIVE_ID = "1Ag-8lfRVQ-vuxaIAjEL0RmLYOO5uNj5V" 


API_KEY_DRIVE_URL = f"https://drive.google.com/uc?export=download&id={API_KEY_DRIVE_ID}"
EXCEL_DRIVE_URL = f"https://drive.google.com/uc?export=download&id={EXCEL_DRIVE_ID}"
ERP_DRIVE_URL = f"https://drive.google.com/uc?export=download&id={ERP_DRIVE_ID}"

def get_api_key():
    """Fetches the live Gemini API key from Google Drive."""
    try:
        response = requests.get(API_KEY_DRIVE_URL)
        if response.status_code == 200:
            return response.text.strip()
    except Exception as e:
        logger.error(f"Failed to fetch API key from Drive: {e}")
    return None