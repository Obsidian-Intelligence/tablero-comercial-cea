"""Conexión y lectura del Google Sheet 'Excel de Ventas CEA' vía gspread."""
import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials

SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

HOJAS = [
    "Prospectos",
    "Eventos",
    "Metas",
    "Alumnos_Activos",
    "Descuentos_Otorgados",
    "Config",
    "OKR_Avance",
    "KPI_Manual",
    "Comité_Lunes",
]


def hay_credenciales() -> bool:
    """True si secrets trae credenciales de Google y el ID del Sheet."""
    try:
        return "gcp_service_account" in st.secrets and "SHEET_ID" in st.secrets
    except Exception:
        return False


@st.cache_resource
def _cliente() -> gspread.Client:
    creds = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]), scopes=SCOPES
    )
    return gspread.authorize(creds)


@st.cache_data(ttl=600, show_spinner="Leyendo Excel de Ventas desde Google Sheets…")
def leer_hojas() -> dict[str, pd.DataFrame]:
    """Lee todas las hojas del Excel de Ventas y regresa {nombre_hoja: DataFrame}."""
    cliente = _cliente()
    libro = cliente.open_by_key(st.secrets["SHEET_ID"])
    resultado: dict[str, pd.DataFrame] = {}
    for nombre in HOJAS:
        try:
            hoja = libro.worksheet(nombre)
            registros = hoja.get_all_records()
            resultado[nombre] = pd.DataFrame(registros)
        except gspread.exceptions.WorksheetNotFound:
            resultado[nombre] = pd.DataFrame()
    return resultado
