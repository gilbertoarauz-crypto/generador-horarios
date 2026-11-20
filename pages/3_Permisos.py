from datetime import datetime, date, timedelta
import io
import altair as alt
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import streamlit as st
from fpdf import FPDF

# ==========================================
# ESTILOS VISUALES (CONSISTENTES CON MÓDULO BRT)
# ==========================================
st.markdown(
    """
    <style>
    @media print {
        [data-testid="stSidebar"], header, footer, .stButton, .stDownloadButton, [data-testid="stToolbar"] {
            display: none !important;
        }
        .block-container {
            padding-top: 0.5rem !important;
            padding-bottom: 0.5rem !important;
            max-width: 100% !important;
        }
    }
    .header-permisos {
        background-color: #0070c0;
        color: #ffffff;
        padding: 7px 12px;
        border-radius: 4px;
        font-weight: bold;
        font-size: 14px;
        text-align: center;
        letter-spacing: 0.5px;
        margin-bottom: 8px;
    }
    .sub-header-permisos {
        background-color: #1e293b;
        color: #f8fafc;
        padding: 5px 10px;
        border-left: 4px solid #00b050;
        font-weight: bold;
        font-size: 13px;
        margin-top: 8px;
        margin-bottom: 6px;
    }
    .boleta-preview {
        border: 2px solid #0f2952;
        padding: 14px;
        border-radius: 6px;
        background-color: #ffffff;
        color: #0f172a;
        font-family: Arial, sans-serif;
    }
    .panel-lateral-anual {
        background-color: #1e293b;
        border: 1px solid #f59e0b;
        border-left: 5px solid #f59e0b;
        color: #f8fafc;
        padding: 10px 12px;
        border-radius: 6px;
        font-size: 13px;
        line-height: 1.45;
    }
    .panel-lateral-anual-ok {
        background-color: #0f172a;
        border: 1px solid #334155;
        border-left: 5px solid #10b981;
        color: #cbd5e1;
        padding: 10px 12px;
        border-radius: 6px;
        font-size: 13px;
        line-height: 1.45;
    }
    .dias-badge {
        background-color: #0f2952;
        border: 1px solid #38bdf8;
        color: #38bdf8;
        padding: 6px 10px;
        border-radius: 5px;
        font-weight: bold;
        font-size: 14px;
        text-align: center;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ==========================================
# CATÁLOGOS OFICIALES
# ==========================================
RESPONSABLES_DEFAULT = [
    "Kevin Pitti",
    "Roberto Velarde",
    "Keily Agrazal",
    "Kathleen Batista",
    "Alexis Mojica",
    "Gilberto Arauz",
]

TIPOS_SOLICITUD = [
    "Permiso Remunerado",
    "Permiso No Remunerado",
    "Licencia Remunerada",
    "Licencia No Remunerada",
    "Compensatorio",
]

MOTIVOS_SOLICITUD_BOLETA = [
    "Enfermedad",
    "Cita Médica",
    "Fallecimiento",
    "Nacimiento hijo (a)",
    "Matrimonio",
    "Asuntos Personales",
    "Representar a la empresa, el Estado o el país",
    "Permiso Sindical",
    "Compensatorio",
]

SUCURSALES_CATALOGO = [
    "PATIO LOS PUEBLOS",
    "PATIO LA DOÑA",
    "PATIO EL CHORRILLO",
    "PATIO OJO DE AGUA",
    "PATIO LA CABIMA",
    "PATIO ALBROOK",
    "PATIO SANTA LIBRADA",
    "CORPORATIVO",
]

NOMBRE_HOJA_PERMISOS = "HISTORIAL DE PERMISOS"

COLUMNAS_PERMISOS = [
    "Key",
    "Codigo",
    "Tipo de Solicitud",
    "Fecha Salida",
    "Fecha de Ingreso",
    "Motivo Solicitud",
    "Observaciones",
    "TOTAL DIAS",
    "Hecho Por",
    "Astamp",
]


def limpiar_codigo(val):
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def limpiar_texto_pdf(texto):
    if texto is None:
        return ""
    s = str(texto)
    reemplazos = {
        "–": "-",
        "—": "-",
        "’": "'",
        "“": '"',
        "”": '"',
        "•": "*",
        "\xa0": " ",
    }
    for k, v in reemplazos.items():
        s = s.replace(k, v)
    return s.encode("latin-1", "replace").decode("latin-1")


# ==========================================
# CONEXIÓN A GOOGLE SHEETS (PERMISOS + DIRECTORIO EN DRIVE)
# ==========================================
def crear_cliente_gspread():
    if "gcp_service_account" not in st.secrets:
        return None
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]), scopes=scopes
    )
    return gspread.authorize(creds)


def obtener_hoja_permisos_gsheets():
    """Conecta con la pestaña HISTORIAL DE PERMISOS dentro del libro principal SHEET_ID."""
    client = crear_cliente_gspread()
    if client is None or "SHEET_ID" not in st.secrets:
        return None
    spreadsheet = client.open_by_key(st.secrets["SHEET_ID"])
    nombres_hojas = [ws.title for ws in spreadsheet.worksheets()]

    if NOMBRE_HOJA_PERMISOS in nombres_hojas:
        return spreadsheet.worksheet(NOMBRE_HOJA_PERMISOS)
    if "Registro" in nombres_hojas:
        return spreadsheet.worksheet("Registro")

    ws_nueva = spreadsheet.add_worksheet(
        title=NOMBRE_HOJA_PERMISOS, rows=1000, cols=len(COLUMNAS_PERMISOS)
    )
    ws_nueva.append_row(COLUMNAS_PERMISOS)
    return ws_nueva


@st.cache_data(ttl=600, show_spinner="Cargando directorio desde Google Drive...")
def cargar_directorio_desde_drive():
    """
    Lee automáticamente el directorio de colaboradores desde Google Drive:
    1. Si existe SHEET_DIRECTORIO_ID en Secrets, abre ese Google Sheet.
    2. Si no, busca la pestaña 'DIRECTORIO' o 'Directorio Consolidado' en el mismo SHEET_ID.
    Detecta automáticamente las columnas CODIGO y APELLIDO Y NOMBRE sin importar la fila de encabezado.
    """
    try:
        client = crear_cliente_gspread()
        if client is None:
            return {}, False, "Sin credenciales gcp_service_account"

        sheet_key = st.secrets.get("SHEET_DIRECTORIO_ID") or st.secrets.get("SHEET_ID")
        if not sheet_key:
            return {}, False, "Sin SHEET_ID configurado"

        spreadsheet = client.open_by_key(sheet_key)
        worksheets = spreadsheet.worksheets()
        ws_dir = None

        for ws in worksheets:
            if ws.title.strip().upper() in ["DIRECTORIO", "DIRECTORIO CONSOLIDADO"]:
                ws_dir = ws
                break
        if ws_dir is None and "SHEET_DIRECTORIO_ID" in st.secrets:
            ws_dir = worksheets[0]

        if ws_dir is None:
            return {}, False, "No se encontró pestaña 'DIRECTORIO' en Google Sheets"

        valores = ws_dir.get_all_values()
        if not valores:
            return {}, False, "La hoja de directorio está vacía"

        idx_hdr = None
        col_cod = None
        col_nom = None
        col_patio = None

        for i in range(min(15, len(valores))):
            fila_u = [str(c).strip().upper() for c in valores[i]]
            for j, cell_txt in enumerate(fila_u):
                if cell_txt in ["CODIGO", "CÓDIGO", "COD", "COD. OPERADOR"]:
                    col_cod = j
                elif any(
                    k in cell_txt
                    for k in ["APELLIDO Y NOMBRE", "NOMBRE Y APELLIDO", "NOMBRE", "COLABORADOR"]
                ):
                    col_nom = j
                elif cell_txt in ["SUCURSAL", "PATIO", "PADMIN", "PASEC", "PADIM"]:
                    col_patio = j
            if col_cod is not None and col_nom is not None:
                idx_hdr = i
                break

        if idx_hdr is None:
            return {}, False, "No se encontraron columnas CODIGO y NOMBRE en el directorio"

        mapa_patios_siglas = {
            "PAPUE": "PATIO LOS PUEBLOS",
            "PADON": "PATIO LA DOÑA",
            "PACHO": "PATIO EL CHORRILLO",
            "PAOJO": "PATIO OJO DE AGUA",
            "PACAB": "PATIO LA CABIMA",
            "PACUR": "PATIO ALBROOK",
        }

        dir_map = {}
        for fila in valores[idx_hdr + 1 :]:
            if len(fila) <= max(col_cod, col_nom):
                continue
            c = limpiar_codigo(fila[col_cod])
            nom = str(fila[col_nom]).strip()
            if not c or not nom:
                continue

            suc = "PATIO LOS PUEBLOS"
            if col_patio is not None and len(fila) > col_patio:
                p_raw = str(fila[col_patio]).strip().upper()
                for pref, nombre_suc in mapa_patios_siglas.items():
                    if pref in p_raw:
                        suc = nombre_suc
                        break
                else:
                    if p_raw and p_raw != "-":
                        suc = p_raw if "PATIO" in p_raw else f"PATIO {p_raw}"

            dir_map[c] = {"nombre": nom, "patio": suc}

        return dir_map, True, f"{ws_dir.title} ({len(dir_map):,} registros)"
    except Exception as e:
        return {}, False, str(e)


def leer_permisos_desde_gsheets():
    """Lee el historial acumulado de permisos desde Google Sheets."""
    try:
        ws = obtener_hoja_permisos_gsheets()
        if ws is None:
            return pd.DataFrame(columns=COLUMNAS_PERMISOS), False
        records = ws.get_all_records()
        if not records:
            return pd.DataFrame(columns=COLUMNAS_PERMISOS), True
        df_p = pd.DataFrame(records)
        df_p.columns = [str(c).strip() for c in df_p.columns]
        if "Codigo" in df_p.columns:
            df_p = df_p[df_p["Codigo"].astype(str).str.strip() != ""].copy()
            df_p["Codigo"] = df_p["Codigo"].apply(limpiar_codigo)
        return df_p, True
    except Exception:
        return pd.DataFrame(columns=COLUMNAS_PERMISOS), False


def agregar_permiso_en_gsheets(fila_dict: dict):
    """Inserta el nuevo permiso en la fila 2 de la pestaña HISTORIAL DE PERMISOS."""
    try:
        ws = obtener_hoja_permisos_gsheets()
        if ws is None:
            return False, "Faltan credenciales SHEET_ID o gcp_service_account en Secrets."

        encabezados = ws.row_values(1)
        if not encabezados:
            ws.append_row(COLUMNAS_PERMISOS)
            encabezados = COLUMNAS_PERMISOS

        nueva_fila_lista = []
        for col in encabezados:
            val = fila_dict.get(col, "")
            if isinstance(val, (pd.Timestamp, datetime)):
                val = val.strftime("%Y-%m-%d %H:%M:%S")
            elif isinstance(val, date):
                val = val.strftime("%Y-%m-%d")
            nueva_fila_lista.append(str(val) if val is not None else "")

        ws.insert_row(nueva_fila_lista, index=2, value_input_option="USER_ENTERED")
        return True, f"Guardado en pestaña '{ws.title}' de Google Sheets."
    except Exception as e:
        return False, f"Error al escribir en Google Sheets: {e}"


def buscar_colaborador_permiso(codigo: str):
    cod = limpiar_codigo(codigo)
    # 1. Buscar en el Directorio cargado automáticamente desde Google Drive
    if "directorio_drive" in st.session_state and cod in st.session_state.directorio_drive:
        info = st.session_state.directorio_drive[cod]
        return info["nombre"], info.get("patio", "PATIO LOS PUEBLOS"), True

    # 2. Buscar en la Asignación activa si se cargó en Cambio de Horarios
    if "operadores_dict" in st.session_state and cod in st.session_state.operadores_dict:
        info = st.session_state.operadores_dict[cod]
        patio_raw = str(info.get("patio", "LOS PUEBLOS")).upper()
        suc = patio_raw if "PATIO" in patio_raw else f"PATIO {patio_raw}"
        return info["nombre"], suc, True

    return "COLABORADOR NO ENCONTRADO EN DIRECTORIO", "PATIO LOS PUEBLOS", False


def renderizar_resumen_anual_permisos(codigo: str, df_perm: pd.DataFrame, anio_ref: int):
    if not codigo:
        st.markdown(
            '<div class="panel-lateral-anual-ok">📋 Ingresa un código de colaborador para ver su historial de permisos y licencias.</div>',
            unsafe_allow_html=True,
        )
        return

    if df_perm.empty or "Codigo" not in df_perm.columns:
        st.markdown(
            f'<div class="panel-lateral-anual-ok">✅ <b>Historial {anio_ref} — Cód. {codigo}</b><br>Sin permisos registrados aún.</div>',
            unsafe_allow_html=True,
        )
        return

    sub = df_perm[df_perm["Codigo"].astype(str).str.strip() == codigo].copy()
