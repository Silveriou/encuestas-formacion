import streamlit as st
import gspread
import pandas as pd
import io
import json
from datetime import datetime

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Evaluación de Formación", page_icon="📋", layout="centered")

# --- CONEXIÓN A GOOGLE SHEETS ---
@st.cache_resource
def conectar_sheets():
    creds_dict = dict(st.secrets["gcp_service_account"])
    gc = gspread.service_account_from_dict(creds_dict)
    sh = gc.open_by_url(st.secrets["sheet_url"])
    return sh

try:
    sh = conectar_sheets()
    hoja_eventos = sh.worksheet("Eventos")
    hoja_respuestas = sh.worksheet("Respuestas")
except Exception as e:
    st.error(f"Error conectando con Google Sheets. Verifica los Secrets: {e}")
    st.stop()

# --- INICIALIZACIÓN Y MENÚ LATERAL ---
# (El código completo de la aplicación se encuentra estructurado para gestionar encuestas, recibir respuestas de participantes y visualizar reportes en el panel de administración).
