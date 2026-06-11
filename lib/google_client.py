"""
google_client.py
Singleton para Drive v3 + gspread.
Soporta credenciales via st.secrets["GOOGLE_SERVICE_ACCOUNT"] o creds_google.json local.
"""
import streamlit as st
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import gspread

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


@st.cache_resource
def get_google_clients():
    """Retorna (drive_client, gspread_client). Cacheado como resource (1 instancia por proceso)."""
    try:
        if "GOOGLE_SERVICE_ACCOUNT" in st.secrets:
            info = {k: v for k, v in st.secrets["GOOGLE_SERVICE_ACCOUNT"].items()}
            # Streamlit guarda el private_key con \\n literales en algunos entornos
            if "private_key" in info:
                info["private_key"] = info["private_key"].replace("\\n", "\n")
            creds = Credentials.from_service_account_info(info, scopes=SCOPES)
        else:
            creds = Credentials.from_service_account_file("creds_google.json", scopes=SCOPES)

        drive = build("drive", "v3", credentials=creds)
        gc = gspread.authorize(creds)
        return drive, gc

    except Exception as e:
        st.error(f"❌ Error conectando a Google: {e}")
        st.stop()
