from datetime import datetime, date, time, timedelta
import io
import pandas as pd
import streamlit as st
from fpdf import FPDF

# Estilo CSS para impresión directa limpia y apariencia compacta tipo módulo BRT
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
    .header-brt {
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
    .sub-header-brt {
        background-color: #1e293b;
        color: #f8fafc;
        padding: 5px 10px;
        border-left: 4px solid #0070c0;
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
    .rango-banner {
        background-color: #0f2952;
        border: 1px solid #38bdf8;
        color: #e0f2fe;
        padding: 8px 14px;
        border-radius: 6px;
        font-size: 13.5px;
        margin-bottom: 12px;
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

PATIOS_CATALOGO = [
    "Patio Los Pueblos",
    "Patio La Doña",
    "Patio Chorrillo",
    "Patio Ojo de Agua",
    "Patio La Cabima",
    "Patio Curundu",
    "Metro Pedregal",
    "Por Asignar",
]

TAREAS_CATALOGO = [
    "DISPONIBILIDAD",
    "REFUERZO",
    "ACERCAMIENTO",
    "INHABILITADO",
    "MANTENIMIENTO",
    "TANQUEO",
    "DESCANSO",
    "CAPACITACIÓN",
]

JORNADAS_CATALOGO = [
    "-",
    "02:00-10:00",
    "03:00-11:00",
    "04:00-12:00",
    "05:00-11:40",
    "05:00-13:00",
    "06:00-14:40",
    "07:00-15:40",
    "08:00-16:40",
    "09:00-17:40",
    "10:00-18:00",
    "11:00-19:00",
    "12:00-20:00",
    "13:00-21:00",
    "14:00-21:40",
    "15:00-22:40",
    "16:00-23:40",
    "17:00-24:40",
    "18:00-25:40",
    "19:00-27:00",
    "20:00-28:00",
    "21:00-29:00",
    "22:00-30:00",
    "23:00-31:00",
]

MOTIVOS_ESTANDAR = [
    "CAMBIO AUT POR JOP",
    "AUTORIZA JEOP",
    "CT-DISP SEMANALES",
    "REFUERZO",
    "REFUERZO APERTURA",
    "DISP. ARRANQUE DE OP",
    "CAMBIO POR CITA MÉDICA - AUT JOP",
    "ASUNTO PERSONAL - AUT JOP",
    "TRÁMITE JUDICIAL - AUT JOP",
    "CAPACITACIÓN",
]

MESES_ES = {
    1: "Ene",
    2: "Feb",
    3: "Mar",
    4: "Abr",
    5: "May",
    6: "Jun",
    7: "Jul",
    8: "Ago",
    9: "Sep",
    10: "Oct",
    11: "Nov",
    12: "Dic",
}

MESES_LARGO_ES = {
    1: "enero",
    2: "febrero",
    3: "marzo",
    4: "abril",
    5: "mayo",
    6: "junio",
    7: "julio",
    8: "agosto",
    9: "septiembre",
    10: "octubre",
    11: "noviembre",
    12: "diciembre",
}


# ==========================================
# UTILIDADES DE LIMPIEZA Y LECTURA UNIVERSAL
# ==========================================
def formatear_hora_celda(val):
    if pd.isna(val) or val == "":
        return ""
    if isinstance(val, timedelta):
        total_sec = int(val.total_seconds())
        hh = total_sec // 3600
        mm = (total_sec % 3600) // 60
        return f"{hh:02d}:{mm:02d}"
    if isinstance(val, time):
        return val.strftime("%H:%M")
    if isinstance(val, datetime):
        return val.strftime("%H:%M")
    s = str(val).strip()
    if "1900-01-01" in s:
        return "24:00"
    if len(s) >= 5 and ":" in s:
        return s[:5]
    return s


def limpiar_codigo(val):
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def normalizar_tabla_programacion(df_raw: pd.DataFrame) -> pd.DataFrame:
    df = df_raw.copy()
    cols_upper = [str(c).strip().upper() for c in df.columns]
    if not any(k in cols_upper for k in ["CODIGO", "CÓDIGO"]):
        for idx in range(min(10, len(df))):
            row_vals = [str(x).strip().upper() for x in df.iloc[idx].tolist()]
            if "CODIGO" in row_vals or "CÓDIGO" in row_vals:
                df.columns = [str(x).strip() for x in df.iloc[idx].tolist()]
                df = df.iloc[idx + 1 :].reset_index(drop=True)
                break

    col_map = {}
    for c in df.columns:
        cu = str(c).strip().upper()
        if cu in ["FECHA"]:
            col_map[c] = "Fecha"
        elif cu in ["TAREA", "CODIGO SERCON", "CÓDIGO SERCON"]:
            col_map[c] = "Tarea"
        elif cu in ["CODIGO", "CÓDIGO"]:
            col_map[c] = "CODIGO"
        elif cu in ["NOMBRE", "NOMBRE Y APELLIDO", "OPERADOR"]:
            col_map[c] = "NOMBRE"
        elif cu in ["H.INICIAL", "HORA IN
