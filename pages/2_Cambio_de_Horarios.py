import calendar
from datetime import datetime, date, time, timedelta
import io
import re
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import streamlit as st
from fpdf import FPDF

st.set_page_config(
    page_title="Módulo Cambio de Horario — F-ZEN-GO-07",
    page_icon="🚌",
    layout="wide",
)

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
# CATÁLOGOS OFICIALES Y DURACIONES DE JORNADA
# ==========================================
MINUTOS_FULL_TIME = 8 * 60 + 40  # 8:40 (520 min)
MINUTOS_PART_TIME = 6 * 60 + 30  # 6:30 (390 min)

JORNADAS_TIPO_OPCIONES = [
    "Full Time (8:40)",
    "Part Time (6:30)",
    "Catálogo",
    "Libre (-)",
]

RESPONSABLES_DEFAULT = [
    "Kevin Pitti",
    "Roberto Velarde",
    "Keily Agrazal",
    "Kathleen Batista",
    "Alexis Mojica",
    "Gilberto Arauz",
    "Mayrim Fraguela",
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

MESES_TEXTO_A_NUM = {
    "ene": 1, "jan": 1,
    "feb": 2,
    "mar": 3,
    "abr": 4, "apr": 4,
    "may": 5,
    "jun": 6,
    "jul": 7,
    "ago": 8, "aug": 8,
    "sep": 9, "set": 9,
    "oct": 10,
    "nov": 11,
    "dic": 12, "dec": 12,
}

COLUMNAS_HISTORIAL = [
    "REF",
    "COGIDO",
    "FECHA DE SOLICITUD",
    "FECHA DE CAMBIO",
    "HORA DE CAMBIO",
    "RESPONSABLE DEL CAMBIO",
    "OBSERVACIONES",
    "Stamp",
]


# ==========================================
# UTILIDADES DE LIMPIEZA, HORAS Y DETECCIÓN "A ORDEN"
# ==========================================
def limpiar_codigo(val):
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def calcular_horario_desde_inicio(hora_ini: time, tipo_jornada: str, catalogo_val: str = "-") -> str:
    """Calcula HH:MM-HH:MM sumando 8:40 (Full Time) o 6:30 (Part Time), soportando >24:00."""
    if tipo_jornada == "Libre (-)":
        return "-"
    if tipo_jornada == "Catálogo":
        return catalogo_val or "-"
    duracion_min = MINUTOS_PART_TIME if tipo_jornada == "Part Time (6:30)" else MINUTOS_FULL_TIME
    ini_min = hora_ini.hour * 60 + hora_ini.minute
    fin_min = ini_min + duracion_min
    fh, fm = divmod(fin_min, 60)
    return f"{hora_ini.hour:02d}:{hora_ini.minute:02d}-{fh:02d}:{fm:02d}"


def detectar_es_a_orden(*valores) -> bool:
    """Detecta variaciones de 'A ORDEN' / 'A LA ORDEN' en tarea o horario."""
    texto = " ".join(str(v or "") for v in valores).upper()
    texto_norm = re.sub(r"\s+", " ", texto).strip()
    if not texto_norm:
        return False
    return bool(re.search(r"\bA\s*(?:LA\s+)?ORDEN(?:ES)?\b", texto_norm))


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
        if val.year in (1899, 1900):
            base = datetime(1899, 12, 30)
            diff_min = int((val - base).total_seconds() // 60)
            if 0 <= diff_min <= 48 * 60:
                hh, mm = divmod(diff_min, 60)
                return f"{hh:02d}:{mm:02d}"
        return val.strftime("%H:%M")
    s = str(val).strip()
    if "1900-01-01" in s or "1899-12-31" in s:
        try:
            dt = pd.to_datetime(s)
            diff_min = int((dt.to_pydatetime().replace(tzinfo=None) - datetime(1899, 12, 30)).total_seconds() // 60)
            if 0 <= diff_min <= 48 * 60:
                hh, mm = divmod(diff_min, 60)
                return f"{hh:02d}:{mm:02d}"
        except Exception:
            return "24:00"
    if len(s) >= 5 and ":" in s:
        return s[:5]
    return s


def parsear_fecha_segura(val):
    """Convierte fechas YYYY-MM-DD, DD/MM/YYYY o dd-mmm-yy a date sin invertir mes/día."""
    if pd.isna(val) or val is None or str(val).strip() == "":
        return None
    if isinstance(val, pd.Timestamp):
        return val.date()
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    s = str(val).strip()
    m_iso = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})", s)
    if m_iso:
        try:
            return date(int(m_iso.group(1)), int(m_iso.group(2)), int(m_iso.group(3)))
        except ValueError:
            return None
    m_dmy = re.match(r"^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})", s)
    if m_dmy:
        d, m, y = int(m_dmy.group(1)), int(m_dmy.group(2)), int(m_dmy.group(3))
        if y < 100:
            y += 2000
        try:
            return date(y, m, d)
        except ValueError:
            return None
    m_es = re.match(r"^(\d{1,2})[-/\s]([a-zA-ZáéíóúÁÉÍÓÚ]{3,})[-/\s](\d{2,4})$", s)
    if m_es:
        d = int(m_es.group(1))
        mes_txt = m_es.group(2).lower()[:3]
        y = int(m_es.group(3))
        if y < 100:
            y += 2000
        m = MESES_TEXTO_A_NUM.get(mes_txt)
        if m:
            try:
                return date(y, m, d)
            except ValueError:
                return None
    try:
        dt = pd.to_datetime(s, dayfirst=True, errors="coerce")
        if pd.notna(dt):
            return dt.date()
    except Exception:
        pass
    return None


def clasificar_tipo_cambio_row(row) -> str:
    cod = str(row.get("COGIDO", "")).strip()
    obs = str(row.get("OBSERVACIONES", "")).strip().upper()
    hora = str(row.get("HORA DE CAMBIO", "")).strip()

    if "_" in cod or "INTERCAMBIO" in obs or "CON OTRO OPERADOR" in obs:
        return "Cambio con otro operador"

    base_hora = hora.split(" (+")[0].strip()
    partes = base_hora.split("/")
    destino = partes[1].strip().upper() if len(partes) >= 2 else ""

    if (
        destino in ("-", "DESCANSO", "LIBRE")
        or "DÍA LIBRE" in obs
        or "DIA LIBRE" in obs
        or "CAMBIO DE LIBRE" in obs
    ):
        return "Cambio de día libre"

    return "Cambio de horario"


# ==========================================
# CONEXIÓN PERMANENTE A GOOGLE SHEETS
# ==========================================
def obtener_hoja_google_sheets():
    """Conecta con la hoja de cálculo en Google Drive usando gcp_service_account en Secrets."""
    if "gcp_service_account" not in st.secrets or "SHEET_ID" not in st.secrets:
        return None
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]), scopes=scopes
    )
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(st.secrets["SHEET_ID"])
    nombres_hojas = [ws.title for ws in spreadsheet.worksheets()]
    if "HISTORIAL DE CAMBIOS (2)" in nombres_hojas:
        return spreadsheet.worksheet("HISTORIAL DE CAMBIOS (2)")
    if "HISTORIAL" in nombres_hojas:
        return spreadsheet.worksheet("HISTORIAL")
    return spreadsheet.sheet1


def leer_historial_desde_gsheets():
    """Lee todos los registros acumulados desde Google Sheets conservando su número de fila real."""
    try:
        ws = obtener_hoja_google_sheets()
        if ws is None:
            return pd.DataFrame(columns=COLUMNAS_HISTORIAL), False
        valores = ws.get_all_values()
        if not valores or len(valores) <= 1:
            return pd.DataFrame(columns=COLUMNAS_HISTORIAL), True

        encabezados = [str(c).strip() for c in valores[0]]
        filas_limpias = []
        for idx_sheet, fila in enumerate(valores[1:], start=2):
            fila_padded = fila + [""] * (len(encabezados) - len(fila))
            row_dict = dict(zip(encabezados, fila_padded[: len(encabezados)]))
            if "CODIGO" in row_dict and "COGIDO" not in row_dict:
                row_dict["COGIDO"] = row_dict["CODIGO"]
            cod = limpiar_codigo(row_dict.get("COGIDO", ""))
            if not cod:
                continue
            row_dict["COGIDO"] = cod
            row_dict["_GS_ROW"] = idx_sheet
            filas_limpias.append(row_dict)

        if not filas_limpias:
            return pd.DataFrame(columns=COLUMNAS_HISTORIAL), True
        return pd.DataFrame(filas_limpias), True
    except Exception:
        return pd.DataFrame(columns=COLUMNAS_HISTORIAL), False


def agregar_fila_en_gsheets(fila_dict: dict):
    """Inserta el nuevo registro en la fila 2 (arriba del historial) en Google Sheets."""
    try:
        ws = obtener_hoja_google_sheets()
        if ws is None:
            return False, "Faltan credenciales SHEET_ID o gcp_service_account en Secrets."

        encabezados = ws.row_values(1)
        if not encabezados:
            ws.append_row(COLUMNAS_HISTORIAL)
            encabezados = COLUMNAS_HISTORIAL

        nueva_fila_lista = []
        for col in encabezados:
            col_u = str(col).strip().upper()
            if col_u in ("COGIDO", "CODIGO", "CÓDIGO"):
                val = fila_dict.get("COGIDO", "")
            else:
                val = fila_dict.get(col, "")
            if isinstance(val, (pd.Timestamp, datetime)):
                val = val.strftime("%Y-%m-%d %H:%M:%S")
            elif isinstance(val, date):
                val = val.strftime("%Y-%m-%d")
            nueva_fila_lista.append(str(val) if val is not None else "")

        ws.insert_row(nueva_fila_lista, index=2, value_input_option="USER_ENTERED")
        return True, "Guardado en Google Sheets en tiempo real."
    except Exception as e:
        return False, f"Error al escribir en Google Sheets: {e}"


def eliminar_filas_en_gsheets(filas_gs: list[int]):
    """Elimina filas específicas de Google Sheets en orden descendente para no desplazar índices."""
    try:
        ws = obtener_hoja_google_sheets()
        if ws is None:
            return False, "No hay conexión activa con Google Sheets."
        filas_ordenadas = sorted(
            {int(r) for r in filas_gs if pd.notna(r) and int(r) >= 2}, reverse=True
        )
        for row_num in filas_ordenadas:
            ws.delete_rows(row_num)
        return True, f"Se eliminaron {len(filas_ordenadas)} registro(s) de Google Sheets."
    except Exception as e:
        return False, f"Error al eliminar en Google Sheets: {e}"


# ==========================================
# NORMALIZACIÓN DE PROGRAMACIÓN
# ==========================================
def normalizar_tabla_programacion(df_raw: pd.DataFrame, excluir_a_orden: bool = True) -> pd.DataFrame:
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
    destinos_asignados = set()
    for c in df.columns:
        cu = str(c).strip().upper()
        destino = None
        if cu in ["FECHA"]:
            destino = "Fecha"
        elif cu in ["TAREA", "CODIGO SERCON", "CÓDIGO SERCON"]:
            destino = "Tarea"
        elif cu in ["CODIGO", "CÓDIGO"]:
            destino = "CODIGO"
        elif cu in ["NOMBRE", "NOMBRE Y APELLIDO", "OPERADOR"]:
            destino = "NOMBRE"
        elif cu in ["H.INICIAL", "HORA INICIAL", "HORA I"]:
            destino = "H.Inicial"
        elif cu in ["ORIGEN", "PATIO ORIGEN", "P. INICIO"]:
            destino = "origen"
        elif cu in ["H. FINAL", "H.FINAL", "HORA FIN", "HORA FINAL"]:
            destino = "H. Final"
        elif cu in ["PATIO FINAL", "P. FINAL"]:
            destino = "Patio Final"
        elif cu in ["PADMIN", "PATIO"]:
            destino = "PADMIN"
        elif cu == "HORARIO":
            destino = "HORARIO_RAW"

        if destino and destino not in destinos_asignados:
            col_map[c] = destino
            destinos_asignados.add(destino)

    df = df.rename(columns=col_map)
    if "CODIGO" not in df.columns or "Fecha" not in df.columns:
        return pd.DataFrame()

    df["CODIGO_STR"] = df["CODIGO"].apply(limpiar_codigo)
    df = df[df["CODIGO_STR"] != ""].copy()
    df["FECHA_DATE"] = df["Fecha"].apply(parsear_fecha_segura)
    df = df[df["FECHA_DATE"].notna()].copy()

    if "HORARIO_RAW" in df.columns:
        df["HORARIO_STR"] = df["HORARIO_RAW"].fillna("-").astype(str).str.strip()
    else:
        hi = df["H.Inicial"].apply(formatear_hora_celda) if "H.Inicial" in df.columns else [""] * len(df)
        hf = df["H. Final"].apply(formatear_hora_celda) if "H. Final" in df.columns else [""] * len(df)
        df["HORARIO_STR"] = [
            f"{a}-{b}" if (a and b and a != "-" and b != "-") else "-"
            for a, b in zip(hi, hf)
        ]

    for col_req, default_val in [
        ("NOMBRE", "SIN NOMBRE"),
        ("Tarea", "-"),
        ("origen", "-"),
        ("Patio Final", "-"),
        ("PADMIN", "LOS PUEBLOS"),
    ]:
        if col_req not in df.columns:
            df[col_req] = default_val
        else:
            df[col_req] = df[col_req].fillna(default_val).astype(str).str.strip()

    df["ES_A_ORDEN"] = df.apply(
        lambda r: detectar_es_a_orden(r.get("Tarea"), r.get("HORARIO_STR"), r.get("origen"), r.get("Patio Final")),
        axis=1,
    )
    a_orden_detectados = int(df["ES_A_ORDEN"].sum())

    if excluir_a_orden:
        df = df[~df["ES_A_ORDEN"]].copy()

    df.attrs["a_orden_detectados"] = a_orden_detectados
    return df


def indexar_programacion_en_sesion(df_norm: pd.DataFrame, nombre_fuente: str):
    operadores_dict = {}
    turnos_dict = {}
    fechas_unicas = []

    if not df_norm.empty:
        df_ops = df_norm.drop_duplicates(subset=["CODIGO_STR"], keep="last")
        for _, r in df_ops.iterrows():
            operadores_dict[r["CODIGO_STR"]] = {
                "nombre": r["NOMBRE"],
                "patio": r["PADMIN"].upper() if r["PADMIN"] else "LOS PUEBLOS",
            }
        for _, r in df_norm.iterrows():
            turnos_dict[(r["CODIGO_STR"], r["FECHA_DATE"])] = {
                "horario": r["HORARIO_STR"] if r["HORARIO_STR"] else "-",
                "sercon": r["Tarea"] if r["Tarea"] else "-",
                "p_inicio": r["origen"] if r["origen"] else "-",
                "p_fin": r["Patio Final"] if r["Patio Final"] else "-",
            }
        fechas_unicas = sorted(df_norm["FECHA_DATE"].dropna().unique().tolist())

    st.session_state.df_programacion_raw = df_norm.copy()
    st.session_state.a_orden_omitidos = df_norm.attrs.get("a_orden_detectados", 0)
    st.session_state.operadores_dict = operadores_dict
    st.session_state.turnos_dict = turnos_dict
    st.session_state.fechas_disponibles = fechas_unicas
    st.session_state.fecha_min = fechas_unicas[0] if fechas_unicas else None
    st.session_state.fecha_max = fechas_unicas[-1] if fechas_unicas else None
    st.session_state.prog_total_filas = len(df_norm)
    st.session_state.prog_nombre_archivo = nombre_fuente
    st.session_state.prog_cargada = len(df_norm) > 0
    st.session_state.mostrar_cargador = False


def parsear_texto_pegado_excel(texto_tsv: str, excluir_a_orden: bool = True) -> pd.DataFrame:
    lineas = [l for l in texto_tsv.splitlines() if l.strip()]
    if not lineas:
        return pd.DataFrame()

    df_tsv = pd.read_csv(io.StringIO("\n".join(lineas)), sep="\t", header=None)
    primera_fila_upper = [str(x).strip().upper() for x in df_tsv.iloc[0].tolist()]
    if any(k in primera_fila_upper for k in ["CODIGO", "CÓDIGO", "FECHA"]):
        df_tsv.columns = [str(x).strip() for x in df_tsv.iloc[0].tolist()]
        df_tsv = df_tsv.iloc[1:].reset_index(drop=True)
        return normalizar_tabla_programacion(df_tsv, excluir_a_orden=excluir_a_orden)

    ncols = df_tsv.shape[1]
    if ncols >= 11:
        cols_std = [
            "Ref",
            "Fecha",
            "Tarea",
            "CODIGO",
            "NOMBRE",
            "W.T.",
            "H.Inicial",
            "origen",
            "H. Final",
            "Patio Final",
            "PADMIN",
        ] + [f"Extra_{i}" for i in range(ncols - 11)]
        df_tsv.columns = cols_std
    elif ncols >= 9:
        cols_std = [
            "Fecha",
            "Tarea",
            "CODIGO",
            "NOMBRE",
            "W.T.",
            "H.Inicial",
            "origen",
            "H. Final",
            "Patio Final",
        ] + (["PADMIN"] if ncols >= 10 else [])
        df_tsv.columns = cols_std
    return normalizar_tabla_programacion(df_tsv, excluir_a_orden=excluir_a_orden)


@st.cache_data(show_spinner="Indexando programación completa...")
def procesar_archivo_excel(file_bytes: bytes, excluir_a_orden: bool = True):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    sheets = xls.sheet_names

    lista_resp = list(RESPONSABLES_DEFAULT)
    if "Lista" in sheets:
        df_l = pd.read_excel(xls, sheet_name="Lista")
        if not df_l.empty:
            vals = [
                str(x).strip()
                for x in df_l[df_l.columns[0]].dropna().tolist()
                if str(x).strip() != ""
            ]
            if vals:
                for obligatorio in RESPONSABLES_DEFAULT:
                    if obligatorio not in vals:
                        vals.append(obligatorio)
                lista_resp = vals

    df_prog = pd.DataFrame()
    hojas_candidatas = (
        ["ROW-DATA"] + [s for s in sheets if s != "ROW-DATA"]
        if "ROW-DATA" in sheets
        else sheets
    )
    for sh in hojas_candidatas:
        if sh in [
            "F-ZEN-GO-06",
            "F-ZEN-GO-07",
            "BRT",
            "HISTORIAL DE CAMBIOS",
            "HISTORIAL DE CAMBIOS (2)",
            "Lista",
        ]:
            continue
        df_try = pd.read_excel(xls, sheet_name=sh)
        df_norm = normalizar_tabla_programacion(df_try, excluir_a_orden=excluir_a_orden)
        if not df_norm.empty:
            df_prog = df_norm
            break

    df_hist = pd.DataFrame()
    hist_sheet = (
        "HISTORIAL DE CAMBIOS (2)"
        if "HISTORIAL DE CAMBIOS (2)" in sheets
        else ("HISTORIAL DE CAMBIOS" if "HISTORIAL DE CAMBIOS" in sheets else None)
    )
    if hist_sheet:
        df_h = pd.read_excel(xls, sheet_name=hist_sheet)
        df_h.columns = [str(c).strip() for c in df_h.columns]
        df_h = df_h.dropna(how="all")
        if "CODIGO" in df_h.columns and "COGIDO" not in df_h.columns:
            df_h["COGIDO"] = df_h["CODIGO"]
        if "COGIDO" in df_h.columns:
            df_h = df_h[df_h["COGIDO"].notna()].copy()
            df_h["COGIDO"] = df_h["COGIDO"].apply(limpiar_codigo)
        df_hist = df_h

    return df_prog, df_hist, lista_resp


def obtener_resumen_anual_operador(df_hist: pd.DataFrame, codigo: str, anio_objetivo: int):
    cod = limpiar_codigo(codigo)
    if not cod or df_hist.empty or "COGIDO" not in df_hist.columns:
        return 0, 0, []

    def contiene_codigo(val):
        partes = [p.strip() for p in str(val).split("_")]
        return cod in partes

    sub = df_hist[df_hist["COGIDO"].apply(contiene_codigo)].copy()
    if sub.empty:
        return 0, 0, []

    f_cambio = sub["FECHA DE CAMBIO"].apply(parsear_fecha_segura) if "FECHA DE CAMBIO" in sub.columns else pd.Series([None] * len(sub))
    f_sol = sub["FECHA DE SOLICITUD"].apply(parsear_fecha_segura) if "FECHA DE SOLICITUD" in sub.columns else pd.Series([None] * len(sub))
    fecha_ref = f_cambio.fillna(f_sol)

    sub["FECHA_REF"] = pd.to_datetime(fecha_ref, errors="coerce")
    total_historico = len(sub)

    sub_anio = sub[sub["FECHA_REF"].dt.year == anio_objetivo].copy()
    total_anio = len(sub_anio)

    if sub_anio.empty:
        return total_anio, total_historico, []

    sub_anio["MES_NUM"] = sub_anio["FECHA_REF"].dt.month
    conteo = (
        sub_anio.groupby("MES_NUM")
        .size()
        .reset_index(name="cant")
        .sort_values("MES_NUM", ascending=False)
    )
    detalle_meses = [
        f"{int(r['cant'])} en {MESES_ES.get(int(r['MES_NUM']), str(r['MES_NUM']))}"
        for _, r in conteo.iterrows()
    ]
    return total_anio, total_historico, detalle_meses


def renderizar_tarjeta_lateral_anual(codigo: str, df_hist: pd.DataFrame, anio_ref: int):
    if not codigo:
        st.markdown(
            '<div class="panel-lateral-anual-ok">📋 Ingresa un código de operador para ver sus cambios acumulados en el año.</div>',
            unsafe_allow_html=True,
        )
        return

    tot_anio, tot_hist, det_meses = obtener_resumen_anual_operador(
        df_hist, codigo, anio_ref
    )
    if tot_anio > 0:
        desglose_txt = " · ".join(det_meses)
        extra_hist = (
            f"<br><span style='font-size:11.5px; color:#fde68a;'>Acumulado total historial: {tot_hist} cambio(s)</span>"
            if tot_hist > tot_anio
            else ""
        )
        st.markdown(
            f'<div class="panel-lateral-anual">'
            f'⚠️ <b>Resumen Anual {anio_ref} — Op. {codigo}</b><br>'
            f'Lleva <b>{tot_anio} cambio(s)</b> en el año:<br>'
            f'<b>{desglose_txt}</b>'
            f'{extra_hist}'
            f'</div>',
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            f'<div class="panel-lateral-anual-ok">'
            f'✅ <b>Resumen Anual {anio_ref} — Op. {codigo}</b><br>'
            f'Lleva <b>0 cambios</b> registrados en {anio_ref}.'
            f'</div>',
            unsafe_allow_html=True,
        )


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


def fecha_larga_es(fecha_obj: date) -> str:
    return f"{MESES_LARGO_ES[fecha_obj.month]} {fecha_obj.day}, {fecha_obj.year}"


def fecha_corta_es(fecha_obj: date) -> str:
    return f"{fecha_obj.day:02d}-{MESES_ES[fecha_obj.month].lower()}-{str(fecha_obj.year)[2:]}"


# ==========================================
# GENERADOR PDF FORMATO F-ZEN-GO-07
# ==========================================
def generar_pdf_f_zen_go_07(bol: dict) -> bytes:
    pdf = FPDF(orientation="P", unit="mm", format="Letter")
    pdf.set_auto_page_break(auto=False)
    pdf.add_page()

    azul_hdr = (0, 112, 192)
    gris_sub = (235, 238, 242)
    amarillo_cod = (255, 255, 0)

    def dibujar_logo_mibus(x, y):
        pdf.set_fill_color(0, 112, 192)
        pdf.rect(x, y, 12, 2.2, style="F")
        pdf.set_fill_color(0, 176, 80)
        pdf.rect(x + 2, y + 2.6, 12, 2.2, style="F")
        pdf.set_text_color(0, 51, 102)
        pdf.set_font("Arial", "B", 8)
        pdf.set_xy(x, y + 5.2)
        pdf.cell(15, 3.5, "MiBus", align="C")
        pdf.set_text_color(0, 0, 0)

    def dibujar_checkbox(x, y, marcado: bool, etiqueta: str):
        pdf.set_xy(x, y)
        pdf.set_font("Arial", "B", 6.5)
        pdf.cell(3.2, 3.2, "X" if marcado else "", border=1, align="C")
        pdf.set_xy(x + 4, y - 0.3)
        pdf.cell(45, 3.8, limpiar_texto_pdf(etiqueta), align="L")

    def dibujar_bloque_firmas(y_start, resp_nombre):
        pdf.set_xy(10, y_start)
        pdf.set_font("Arial", "B", 6.2)
        pdf.cell(34, 4, "Nombre del Operador 1:", align="L")
        pdf.set_font("Arial", "", 6.2)
        pdf.cell(62, 4, limpiar_texto_pdf(bol["nombre_1"]), border="B", align="C")
        pdf.set_font("Arial", "B", 6.2)
        pdf.cell(34, 4, "Firma del Operador 1:", align="R")
        pdf.cell(66, 4, "", border="B", ln=1)

        pdf.set_x(10)
        pdf.cell(34, 4, "Nombre del Operador 2:", align="L")
        pdf.set_font("Arial", "", 6.2)
        pdf.cell(
            62,
            4,
            limpiar_texto_pdf(bol["nombre_2"] if bol["es_intercambio"] else ""),
            border="B",
            align="C",
        )
        pdf.set_font("Arial", "B", 6.2)
        pdf.cell(34, 4, "Firma del Operador 2:", align="R")
        pdf.cell(66, 4, "", border="B", ln=1)

        pdf.set_x(10)
        pdf.cell(34, 4, "Nombre Analista de Op.:", align="L")
        pdf.set_font("Arial", "", 6.2)
        pdf.cell(62, 4, limpiar_texto_pdf(resp_nombre), border="B", align="C")
        pdf.set_font("Arial", "B", 6.2)
        pdf.cell(34, 4, "Firma Profesional de Op.:", align="R")
        pdf.cell(66, 4, limpiar_texto_pdf(resp_nombre), border="B", ln=1)

    dibujar_logo_mibus(11, 8)
    pdf.set_xy(28, 8)
    pdf.set_font("Arial", "B", 11)
    pdf.cell(115, 5, "CAMBIO DE HORARIO", align="C")
    pdf.set_font("Arial", "", 7)
    pdf.cell(53, 5, "F-ZEN-GO-07 / V1-12-2023", align="R", ln=1)

    pdf.set_x(28)
    pdf.set_font("Arial", "B", 7.5)
    pdf.cell(168, 4, limpiar_texto_pdf(bol["motivo"]), align="L", ln=1)

    pdf.set_xy(10, 18)
    pdf.set_fill_color(*azul_hdr)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Arial", "B", 7.5)
    pdf.cell(196, 4.5, "I. DATOS GENERALES DEL OPERADOR", border=1, ln=1, align="C", fill=True)

    pdf.set_text_color(0, 0, 0)
    pdf.set_fill_color(*gris_sub)
    pdf.set_font("Arial", "B", 6.5)
    pdf.set_x(10)
    pdf.cell(72, 4, "Nombre y Apellido del Operador", border=1, align="C", fill=True)
    pdf.cell(34, 4, "Codigo del Operador", border=1, align="C", fill=True)
    pdf.cell(42, 4, "Fecha de Solicitud", border=1, align="C", fill=True)
    pdf.cell(48, 4, "Patio", border=1, align="C", fill=True, ln=1)

    pdf.set_x(10)
    pdf.set_font("Arial", "B", 7)
    pdf.cell(72, 5.5, limpiar_texto_pdf(bol["nombre_1"])[:38], border=1, align="C")
    pdf.set_fill_color(*amarillo_cod)
    pdf.cell(34, 5.5, limpiar_texto_pdf(bol["cod_1"]), border=1, align="C", fill=True)
    pdf.cell(42, 5.5, limpiar_texto_pdf(bol["fecha_sol_larga"]), border=1, align="C")
    pdf.cell(48, 5.5, limpiar_texto_pdf(bol["patio"]), border=1, align="C", ln=1)

    pdf.set_x(10)
    pdf.set_fill_color(*azul_hdr)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Arial", "B", 7.5)
    pdf.cell(196, 4.5, "II. DETALLE DEL CAMBIO", border=1, ln=1, align="C", fill=True)
    pdf.set_text_color(0, 0, 0)

    pdf.set_x(10)
    pdf.cell(196, 5.5, "", border=1, ln=1)
    dibujar_checkbox(18, 38.2, bol["modalidad"] == "Cambio con otro operador", "Cambio con otro operador")
    dibujar_checkbox(86, 38.2, bol["modalidad"] == "Cambio de día libre", "Cambio de dia libre")
    dibujar_checkbox(148, 38.2, bol["modalidad"] == "Cambio de horario", "Cambio de horario")

    def dibujar_doble_tabla(y_top, titulo_izq, titulo_der, filas_izq, filas_der):
        pdf.set_xy(10, y_top)
        pdf.set_fill_color(*azul_hdr)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Arial", "B", 6.8)
        pdf.cell(98, 4.2, titulo_izq, border=1, align="C", fill=True)
        pdf.cell(98, 4.2, titulo_der, border=1, align="C", fill=True, ln=1)

        pdf.set_text_color(0, 0, 0)
        pdf.set_fill_color(*gris_sub)
        pdf.set_font("Arial", "B", 5.8)
        w_izq = [6, 17, 13, 21, 20, 21]
        h_izq = ["N", "Fecha", "Codigo", "Horario", "Cod. Sercon", "P. Inicio"]
        w_der = [17, 13, 21, 23, 24]
        h_der = ["Fecha", "Codigo", "Horario", "Cod. Sercon", "P. Inicio"]

        pdf.set_x(10)
        for w, h in zip(w_izq, h_izq):
            pdf.cell(w, 3.8, h, border=1, align="C", fill=True)
        for w, h in zip(w_der, h_der):
            pdf.cell(w, 3.8, h, border=1, align="C", fill=True)
        pdf.ln()

        pdf.set_font("Arial", "", 5.7)
        for idx in range(7):
            pdf.set_x(10)
            f1 = filas_izq[idx] if idx < len(filas_izq) else {}
            f2 = filas_der[idx] if idx < len(filas_der) else {}

            vals1 = [
                str(idx + 1),
                f1.get("fecha_corta", ""),
                f1.get("codigo", ""),
                f1.get("horario", ""),
                f1.get("sercon", ""),
                f1.get("p_inicio", ""),
            ]
            vals2 = [
                f2.get("fecha_corta", ""),
                f2.get("codigo", ""),
                f2.get("horario", ""),
                f2.get("sercon", ""),
                f2.get("p_inicio", ""),
            ]
            for w, v in zip(w_izq, vals1):
                pdf.cell(w, 3.6, limpiar_texto_pdf(v)[:16], border=1, align="C")
            for w, v in zip(w_der, vals2):
                pdf.cell(w, 3.6, limpiar_texto_pdf(v)[:17], border=1, align="C")
            pdf.ln()

        pdf.set_line_width(0.6)
        pdf.line(108, y_top, 108, y_top + 4.2 + 3.8 + 7 * 3.6)
        pdf.set_line_width(0.2)

    dibujar_doble_tabla(
        43,
        "HORARIO PROGRAMADO OPERADOR 1",
        "HORARIO PROGRAMADO OPERADOR 2",
        bol["prog_op1"],
        bol["prog_op2"] if bol["es_intercambio"] else [],
    )
    dibujar_doble_tabla(
        77.5,
        "HORARIO CAMBIO OPERADOR 1",
        "HORARIO CAMBIO OPERADOR 2",
        bol["cambio_op1"],
        bol["cambio_op2"] if bol["es_intercambio"] else [],
    )

    dibujar_bloque_firmas(112, bol["responsable"])

    def dibujar_seccion_copia(y_base, titulo_tabla, filas_cambio):
        pdf.set_xy(10, y_base)
        pdf.set_font("Arial", "", 6)
        pdf.cell(196, 2.5, "- " * 88, align="C", ln=1)

        dibujar_logo_mibus(11, y_base + 3)
        pdf.set_xy(28, y_base + 3)
        pdf.set_font("Arial", "B", 10)
        pdf.cell(168, 4.5, "CAMBIO DE HORARIO", align="C", ln=1)

        pdf.set_x(28)
        pdf.set_font("Arial", "B", 6.8)
        pdf.cell(95, 4, limpiar_texto_pdf(bol["motivo"]), align="L")
        pdf.cell(
            73,
            4,
            limpiar_texto_pdf(f"Fecha de Autorizacion:   {bol['fecha_sol']}"),
            align="R",
            ln=1,
        )

        pdf.set_xy(10, y_base + 12)
        pdf.set_fill_color(*azul_hdr)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Arial", "B", 7)
        pdf.cell(196, 4.2, titulo_tabla, border=1, align="C", fill=True, ln=1)

        pdf.set_text_color(0, 0, 0)
        pdf.set_fill_color(*gris_sub)
        pdf.set_font("Arial", "B", 6.2)
        anchos = [8, 25, 22, 33, 38, 38, 32]
        headers = ["N", "Fecha", "Codigo", "Horario", "P. Inicio", "P. Final", "Codigo Sercon"]
        pdf.set_x(10)
        for w, h in zip(anchos, headers):
            pdf.cell(w, 3.8, h, border=1, align="C", fill=True)
        pdf.ln()

        pdf.set_font("Arial", "", 6.2)
        for idx in range(7):
            pdf.set_x(10)
            f = filas_cambio[idx] if idx < len(filas_cambio) else {}
            vals = [
                str(idx + 1),
                f.get("fecha_corta", ""),
                f.get("codigo", ""),
                f.get("horario", ""),
                f.get("p_inicio", ""),
                f.get("p_fin", ""),
                f.get("sercon", ""),
            ]
            for w, v in zip(anchos, vals):
                pdf.cell(w, 3.6, limpiar_texto_pdf(v)[:24], border=1, align="C")
            pdf.ln()

        dibujar_bloque_firmas(y_base + 46, bol["responsable"])
        pdf.set_xy(10, y_base + 58.5)
        pdf.set_font("Arial", "B", 5.8)
        pdf.cell(196, 3, "COPIA NO CONTROLADA", align="C")

    dibujar_seccion_copia(126, "HORARIO DE CAMBIO OPERADOR 1", bol["cambio_op1"])
    dibujar_seccion_copia(
        190,
        "HORARIO DE CAMBIO OPERADOR 2",
        bol["cambio_op2"] if bol["es_intercambio"] else [],
    )

    pdf.set_draw_color(0, 112, 192)
    pdf.set_line_width(0.5)
    pdf.rect(9, 7, 198, 247)

    out = pdf.output(dest="S")
    return out.encode("latin-1") if isinstance(out, str) else bytes(out)


# ==========================================
# INICIALIZAR ESTADO PERSISTENTE EN SESIÓN
# ==========================================
if "prog_cargada" not in st.session_state:
    st.session_state.prog_cargada = False
    st.session_state.prog_nombre_archivo = ""
    st.session_state.df_programacion_raw = pd.DataFrame()
    st.session_state.a_orden_omitidos = 0
    st.session_state.excluir_a_orden = True
    st.session_state.operadores_dict = {}
    st.session_state.turnos_dict = {}
    st.session_state.fechas_disponibles = []
    st.session_state.fecha_min = None
    st.session_state.fecha_max = None
    st.session_state.prog_total_filas = 0
    st.session_state.lista_responsables = list(RESPONSABLES_DEFAULT)
    st.session_state.mostrar_cargador = True
    st.session_state.num_filas_ind = 1
    st.session_state.num_filas_int = 1

for r_def in RESPONSABLES_DEFAULT:
    if r_def not in st.session_state.lista_responsables:
        st.session_state.lista_responsables.append(r_def)

if "historial_sincronizado_gs" not in st.session_state:
    df_gs, ok_gs = leer_historial_desde_gsheets()
    st.session_state.historial_cambios = df_gs
    st.session_state.historial_en_gsheets = ok_gs
    st.session_state.historial_sincronizado_gs = True


# ==========================================
# CALLBACKS DE JORNADA Y MODALIDAD
# ==========================================
def on_cambio_jornada_base():
    nueva_jornada = st.session_state.get("jornada_base_global", "Full Time (8:40)")
    for i in range(7):
        k_tipo = f"brt_ind_tipo_j_{i}"
        k_ini = f"brt_ind_hini_{i}"
        k_hor = f"brt_ind_hor_{i}"
        if st.session_state.get(k_tipo) in ("Full Time (8:40)", "Part Time (6:30)"):
            st.session_state[k_tipo] = nueva_jornada
            h_ini = st.session_state.get(k_ini, time(6, 0))
            st.session_state[k_hor] = calcular_horario_desde_inicio(h_ini, nueva_jornada)


def on_cambio_modalidad():
    mod = st.session_state.get("modalidad_sel", "Cambio de horario")
    jornada_base = st.session_state.get("jornada_base_global", "Full Time (8:40)")
    for i in range(7):
        if mod == "Cambio de día libre":
            st.session_state[f"brt_ind_tipo_j_{i}"] = "Libre (-)"
            st.session_state[f"brt_ind_hor_{i}"] = "-"
            st.session_state[f"brt_ind_tar_{i}"] = "DESCANSO"
        else:
            if st.session_state.get(f"brt_ind_tipo_j_{i}") == "Libre (-)":
                st.session_state[f"brt_ind_tipo_j_{i}"] = jornada_base
            tipo_j = st.session_state.get(f"brt_ind_tipo_j_{i}", jornada_base)
            h_ini = st.session_state.get(f"brt_ind_hini_{i}", time(6, 0))
            st.session_state[f"brt_ind_hor_{i}"] = calcular_horario_desde_inicio(
                h_ini, tipo_j, st.session_state.get(f"brt_ind_cat_{i}", "-")
            )
            if st.session_state.get(f"brt_ind_tar_{i}") == "DESCANSO":
                st.session_state[f"brt_ind_tar_{i}"] = "DISPONIBILIDAD"


# ==========================================
# SIDEBAR: CARGA PERSISTENTE + GOOGLE SHEETS
# ==========================================
st.sidebar.header("📂 1. Programación Activa")

st.session_state.excluir_a_orden = st.sidebar.checkbox(
    "🧹 Eliminar turnos 'A ORDEN' automáticamente",
    value=st.session_state.excluir_a_orden,
)

if st.session_state.prog_cargada and not st.session_state.mostrar_cargador:
    f_ini_txt = st.session_state.fecha_min.strftime("%d/%m/%Y")
    f_fin_txt = st.session_state.fecha_max.strftime("%d/%m/%Y")
    msg_a_orden = (
        f"\n• 🧹 **{st.session_state.a_orden_omitidos:,}** turnos 'A ORDEN' eliminados"
        if st.session_state.a_orden_omitidos > 0
        else ""
    )
    st.sidebar.success(
        f"📌 **Fuente activa:**\n{st.session_state.prog_nombre_archivo}\n\n"
        f"📅 **Fecha Inicial:** {f_ini_txt}\n"
        f"📅 **Fecha Final:** {f_fin_txt}\n\n"
        f"• **{st.session_state.prog_total_filas:,}** turnos válidos indexados"
        f"{msg_a_orden}"
    )
    if st.sidebar.button("🔄 Cambiar / Pegar nueva programación", use_container_width=True):
        st.session_state.mostrar_cargador = True
        st.rerun()
else:
    metodo_carga = st.sidebar.radio(
        "¿Cómo deseas cargar los horarios?",
        ["📁 Subir archivo Excel", "📋 Pegar desde Excel (Ctrl + V)"],
    )

    if metodo_carga == "📁 Subir archivo Excel":
        archivo_subido = st.sidebar.file_uploader(
            "Sube CAMBIO DE HORARIOS 2026 o ASIGNACIÓN (.xlsx / .xlsm)",
            type=["xlsx", "xlsm"],
            key="uploader_prog_persistente",
        )
        if archivo_subido is not None:
            df_prog_l, df_h_loaded, resp_l = procesar_archivo_excel(
                archivo_subido.getvalue(),
                excluir_a_orden=st.session_state.excluir_a_orden,
            )
            st.session_state.lista_responsables = resp_l
            if st.session_state.historial_cambios.empty and not df_h_loaded.empty:
                st.session_state.historial_cambios = df_h_loaded.copy()
            indexar_programacion_en_sesion(df_prog_l, archivo_subido.name)
            st.rerun()
    else:
        st.sidebar.caption(
            "Copia las filas en Excel (`Ctrl + C`) desde **Fecha** hasta **PADMIN** y pégalas aquí:"
        )
        texto_pegado = st.sidebar.text_area(
            "Pega aquí con Ctrl + V",
            height=160,
            placeholder="Pega las celdas copiadas de Excel...",
            key="area_pegar_excel",
        )
        if st.sidebar.button(
            "✅ Cargar programación pegada",
            type="primary",
            use_container_width=True,
        ):
            if texto_pegado.strip():
                df_pegado = parsear_texto_pegado_excel(
                    texto_pegado, excluir_a_orden=st.session_state.excluir_a_orden
                )
                if not df_pegado.empty:
                    indexar_programacion_en_sesion(
                        df_pegado,
                        f"Pegado desde Excel ({datetime.now().strftime('%H:%M')})",
                    )
                    st.rerun()
                else:
                    st.sidebar.error(
                        "⚠️ No se reconocieron las columnas. Asegúrate de copiar desde Fecha hasta PADMIN."
                    )
            else:
                st.sidebar.warning("⚠️ Pega primero los datos copiados de Excel.")

    if st.session_state.prog_cargada:
        if st.sidebar.button("Cancelar cambio", use_container_width=True):
            st.session_state.mostrar_cargador = False
            st.rerun()

st.sidebar.markdown("---")
st.sidebar.header("☁️ 2. Google Drive / Sheets")
if st.session_state.historial_en_gsheets:
    st.sidebar.success(
        f"🟢 Conectado a Google Sheets\n\n📊 **{len(st.session_state.historial_cambios)}** registros sincronizados."
    )
else:
    st.sidebar.warning(
        "⚠️ Verifica `SHEET_ID` y `gcp_service_account` en Secrets de Streamlit."
    )

if st.sidebar.button("🔄 Recargar desde Google Sheets", use_container_width=True):
    df_gs, ok_gs = leer_historial_desde_gsheets()
    st.session_state.historial_cambios = df_gs
    st.session_state.historial_en_gsheets = ok_gs
    if ok_gs:
        st.sidebar.success("✅ Historial actualizado desde Google Drive.")
    st.rerun()


def buscar_operador_rapido(codigo: str):
    cod = limpiar_codigo(codigo)
    info = st.session_state.operadores_dict.get(cod)
    if info:
        return info["nombre"], info["patio"], True
    return "OPERADOR NO ENCONTRADO EN PROGRAMACIÓN", "LOS PUEBLOS", False


def buscar_turno_rapido(codigo: str, fecha_obj: date):
    cod = limpiar_codigo(codigo)
    turno = st.session_state.turnos_dict.get((cod, fecha_obj))
    if turno:
        return turno
    return {"horario": "-", "sercon": "-", "p_inicio": "-", "p_fin": "-"}


# ==========================================
# ENCABEZADO TIPO MÓDULO BRT + RANGO ACTIVO
# ==========================================
st.markdown(
    '<div class="header-brt">🚌 MÓDULO CAMBIO DE HORARIO — EMISIÓN F-ZEN-GO-07</div>',
    unsafe_allow_html=True,
)

if st.session_state.prog_cargada:
    f_ini_b = st.session_state.fecha_min.strftime("%d/%m/%Y")
    f_fin_b = st.session_state.fecha_max.strftime("%d/%m/%Y")
    n_dias_prog = len(st.session_state.fechas_disponibles)
    st.markdown(
        f'<div class="rango-banner">'
        f'📅 <b>Programación Cargada:</b> Desde el <b>{f_ini_b}</b> hasta el <b>{f_fin_b}</b> '
        f'({n_dias_prog} días operativos disponibles · {st.session_state.prog_total_filas:,} turnos válidos) &nbsp;|&nbsp; '
        f'<i>Turnos "A ORDEN" filtrados automáticamente.</i>'
        f'</div>',
        unsafe_allow_html=True,
    )
else:
    st.warning(
        "⚠️ **Aún no has cargado ninguna programación.** Sube el archivo Excel o pega los horarios en la barra lateral izquierda (o en la pestaña 5) para habilitar el calendario."
    )

# ==========================================
# PESTAÑAS PRINCIPALES
# ==========================================
tab_registro, tab_boleta, tab_historial, tab_estadisticas, tab_pegar_tabla = st.tabs(
    [
        "📝 1. Módulo Cambio de Horario (BRT)",
        "🖨️ 2. Boleta Oficial F-ZEN-GO-07 (PDF)",
        "📋 3. Historial Acumulado y Descarga Excel",
        "📈 4. Estadísticas Anuales",
        "🗂️ 5. Tabla Tipo Excel (Pegar / Editar)",
    ]
)

# ------------------------------------------
# TAB 1: INTERFAZ COMPACTA TIPO HOJA BRT
# ------------------------------------------
with tab_registro:
    c_top1, c_top2, c_top3, c_top4 = st.columns([1.3, 1, 1.3, 1.5])
    with c_top1:
        modalidad = st.selectbox(
            "Tipo de Movimiento",
            [
                "Cambio de horario",
                "Cambio con otro operador",
                "Cambio de día libre",
            ],
            key="modalidad_sel",
            on_change=on_cambio_modalidad,
        )
    with c_top2:
        fecha_solicitud = st.date_input("Fecha de Solicitud", value=date.today())
    with c_top3:
        responsable = st.selectbox(
            "Responsable del Cambio", st.session_state.lista_responsables
        )
    with c_top4:
        motivo_sel = st.selectbox("Razón / Autorización", MOTIVOS_ESTANDAR)

    motivo_extra = st.text_input(
        "Observación adicional (Opcional)",
        placeholder="Ej. Cita médica / Autoriza JOP / Cubre servicio...",
    )
    motivo_final = (
        f"{motivo_sel} - {motivo_extra.strip().upper()}"
        if motivo_extra.strip()
        else motivo_sel
    )

    es_intercambio = modalidad == "Cambio con otro operador"
    anio_consulta = (
        st.session_state.fecha_min.year
        if st.session_state.prog_cargada and st.session_state.fecha_min
        else fecha_solicitud.year
    )

    if not st.session_state.prog_cargada:
        st.info("👈 Carga primero la programación en el panel izquierdo para comenzar.")
    else:
        fechas_disp = st.session_state.fechas_disponibles
        mapa_fechas_label = {
            f"{f.strftime('%d/%m/%Y')} ({fecha_corta_es(f)})": f for f in fechas_disp
        }
        opciones_fecha_sel = ["— Sin cambio —"] + list(mapa_fechas_label.keys())

        filas_prog_op1, filas_cambio_op1 = [], []
        filas_prog_op2, filas_cambio_op2 = [], []
        cod_op1, nom_op1, patio_op1 = "", "", "LOS PUEBLOS"
        cod_op2, nom_op2 = "", ""

        if not es_intercambio:
            st.markdown(
                '<div class="sub-header-brt">1. DATOS DEL OPERADOR Y DETALLE DE CAMBIO (FILAS 1 A 7)</div>',
                unsafe_allow_html=True,
            )
            col_op_cod, col_op_nom, col_op_lat = st.columns([1, 1.8, 1.8])
            with col_op_cod:
                cod_op1 = limpiar_codigo(st.text_input("Código Operador", value="82601"))
            with col_op_nom:
                nom_op1, patio_op1, ok1 = buscar_operador_rapido(cod_op1)
                st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
                if ok1:
                    st.success(f"👤 **{nom_op1}** · Patio: **{patio_op1}**")
                else:
                    st.warning(f"👤 **{nom_op1}**")
            with col_op_lat:
                renderizar_tarjeta_lateral_anual(
                    cod_op1,
                    st.session_state.historial_cambios,
                    anio_consulta,
                )

            c_ctrl1, c_ctrl2, c_ctrl3 = st.columns([1.8, 1.1, 1.1])
            with c_ctrl1:
                st.radio(
                    "Jornada base para cálculo desde Hora de Inicio:",
                    ["Full Time (8:40)", "Part Time (6:30)"],
                    horizontal=True,
                    key="jornada_base_global",
                    on_change=on_cambio_jornada_base,
                    disabled=(modalidad == "Cambio de día libre"),
                )
            with c_ctrl2:
                if st.button("➕ Agregar fecha (máx. 7)", use_container_width=True, disabled=st.session_state.num_filas_ind >= 7):
                    st.session_state.num_filas_ind = min(7, st.session_state.num_filas_ind + 1)
                    st.rerun()
            with c_ctrl3:
                st.caption(f"Filas activas: **{st.session_state.num_filas_ind} / 7**")

            h_cols = st.columns([0.3, 1.35, 1.2, 1.05, 0.9, 1.05, 1.05, 1.1, 1.1, 0.45])
            h_cols[0].markdown("**#**")
            h_cols[1].markdown("**Fecha**")
            h_cols[2].markdown("**Turno Actual**")
            h_cols[3].markdown("**Jornada**")
            h_cols[4].markdown("**H. Inicio**")
            h_cols[5].markdown("**Nuevo Horario**")
            h_cols[6].markdown("**Tarea / Sercon**")
            h_cols[7].markdown("**P. Inicio**")
            h_cols[8].markdown("**P. Final**")
            h_cols[9].markdown("**Quitar**")

            for idx in range(st.session_state.num_filas_ind):
                r_cols = st.columns([0.3, 1.35, 1.2, 1.05, 0.9, 1.05, 1.05, 1.1, 1.1, 0.45])
                r_cols[0].markdown(f"**{idx + 1}**")
                default_idx = 1 if (idx == 0 and len(opciones_fecha_sel) > 1) else 0
                sel_f_str = r_cols[1].selectbox(
                    f"F{idx+1}",
                    opciones_fecha_sel,
                    index=default_idx,
                    key=f"brt_ind_fecha_{idx}",
                    label_visibility="collapsed",
                )

                if sel_f_str != "— Sin cambio —":
                    f_obj = mapa_fechas_label[sel_f_str]
                    prog1 = buscar_turno_rapido(cod_op1, f_obj)
                    r_cols[2].info(f"`{prog1['horario']}` · {prog1['sercon']}")

                    tipo_key = f"brt_ind_tipo_j_{idx}"
                    if tipo_key not in st.session_state:
                        st.session_state[tipo_key] = (
                            "Libre (-)"
                            if modalidad == "Cambio de día libre"
                            else st.session_state.get("jornada_base_global", "Full Time (8:40)")
                        )

                    tipo_j = r_cols[3].selectbox(
                        f"TJ{idx+1}",
                        JORNADAS_TIPO_OPCIONES,
                        key=tipo_key,
                        label_visibility="collapsed",
                    )

                    if tipo_j in ("Full Time (8:40)", "Part Time (6:30)"):
                        h_ini_key = f"brt_ind_hini_{idx}"
                        if h_ini_key not in st.session_state:
                            st.session_state[h_ini_key] = time(6, 0)
                        h_ini_val = r_cols[4].time_input(
                            f"HI{idx+1}",
                            key=h_ini_key,
                            step=timedelta(minutes=10),
                            label_visibility="collapsed",
                        )
                        nuevo_hor = calcular_horario_desde_inicio(h_ini_val, tipo_j)
                        r_cols[5].success(f"**{nuevo_hor}**")
                    elif tipo_j == "Catálogo":
                        r_cols[4].caption("Catálogo")
                        nuevo_hor = r_cols[5].selectbox(
                            f"CAT{idx+1}",
                            JORNADAS_CATALOGO,
                            index=6,
                            key=f"brt_ind_cat_{idx}",
                            label_visibility="collapsed",
                        )
                    else:
                        r_cols[4].caption("—")
                        nuevo_hor = "-"
                        r_cols[5].caption("Libre (-)")

                    tar_key = f"brt_ind_tar_{idx}"
                    if tar_key not in st.session_state:
                        st.session_state[tar_key] = (
                            "DESCANSO" if modalidad == "Cambio de día libre" else "DISPONIBILIDAD"
                        )
                    nueva_tar = r_cols[6].selectbox(
                        f"T{idx+1}",
                        TAREAS_CATALOGO,
                        key=tar_key,
                        label_visibility="collapsed",
                    )
                    nuevo_pi = r_cols[7].selectbox(
                        f"PI{idx+1}",
                        PATIOS_CATALOGO,
                        key=f"brt_ind_pi_{idx}",
                        label_visibility="collapsed",
                    )
                    nuevo_pf = r_cols[8].selectbox(
                        f"PF{idx+1}",
                        PATIOS_CATALOGO,
                        key=f"brt_ind_pf_{idx}",
                        label_visibility="collapsed",
                    )

                    filas_prog_op1.append(
                        {
                            "fecha": f_obj.strftime("%d/%m/%Y"),
                            "fecha_corta": fecha_corta_es(f_obj),
                            "codigo": cod_op1,
                            **prog1,
                        }
                    )
                    filas_cambio_op1.append(
                        {
                            "fecha": f_obj.strftime("%d/%m/%Y"),
                            "fecha_corta": fecha_corta_es(f_obj),
                            "codigo": cod_op1,
                            "horario": nuevo_hor,
                            "sercon": nueva_tar,
                            "p_inicio": nuevo_pi,
                            "p_fin": nuevo_pf,
                        }
                    )
                else:
                    for col_empty in r_cols[2:9]:
                        col_empty.caption("—")

                if r_cols[9].button("🗑️", key=f"del_ind_{idx}", help="Eliminar fila"):
                    if st.session_state.num_filas_ind > 1:
                        for j in range(idx, st.session_state.num_filas_ind - 1):
                            for pref in ("brt_ind_fecha_", "brt_ind_tipo_j_", "brt_ind_hini_", "brt_ind_cat_", "brt_ind_tar_", "brt_ind_pi_", "brt_ind_pf_"):
                                if f"{pref}{j+1}" in st.session_state:
                                    st.session_state[f"{pref}{j}"] = st.session_state[f"{pref}{j+1}"]
                        last = st.session_state.num_filas_ind - 1
                        for pref in ("brt_ind_fecha_", "brt_ind_tipo_j_", "brt_ind_hini_", "brt_ind_cat_", "brt_ind_tar_", "brt_ind_pi_", "brt_ind_pf_"):
                            st.session_state.pop(f"{pref}{last}", None)
                        st.session_state.num_filas_ind -= 1
                    else:
                        st.session_state["brt_ind_fecha_0"] = "— Sin cambio —"
                    st.rerun()

        else:
            st.markdown(
                '<div class="sub-header-brt">2. INTERCAMBIO ENTRE OPERADORES (FILAS 1 A 7)</div>',
                unsafe_allow_html=True,
            )
            c1_cod, c1_nom, c1_lat = st.columns([0.9, 1.8, 1.8])
            with c1_cod:
                cod_op1 = limpiar_codigo(st.text_input("Operador 1 (Código)", value="82601"))
            with c1_nom:
                nom_op1, patio_op1, ok1 = buscar_operador_rapido(cod_op1)
                st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
                if ok1:
                    st.success(f"👤 **Op. 1:** {nom_op1} · Patio: **{patio_op1}**")
                else:
                    st.warning(f"👤 **Op. 1:** {nom_op1}")
            with c1_lat:
                renderizar_tarjeta_lateral_anual(
                    cod_op1,
                    st.session_state.historial_cambios,
                    anio_consulta,
                )

            c2_cod, c2_nom, c2_lat = st.columns([0.9, 1.8, 1.8])
            with c2_cod:
                cod_op2 = limpiar_codigo(st.text_input("Operador 2 (Código)", value="82602"))
            with c2_nom:
                nom_op2, _, ok2 = buscar_operador_rapido(cod_op2)
                st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
                if ok2:
                    st.success(f"👤 **Op. 2:** {nom_op2}")
                else:
                    st.warning(f"👤 **Op. 2:** {nom_op2}")
            with c2_lat:
                renderizar_tarjeta_lateral_anual(
                    cod_op2,
                    st.session_state.historial_cambios,
                    anio_consulta,
                )

            ci_btn, ci_cap = st.columns([1.2, 3])
            with ci_btn:
                if st.button("➕ Agregar fecha (máx. 7)", key="add_int_row", use_container_width=True, disabled=st.session_state.num_filas_int >= 7):
                    st.session_state.num_filas_int = min(7, st.session_state.num_filas_int + 1)
                    st.rerun()
            with ci_cap:
                st.caption(f"Filas activas: **{st.session_state.num_filas_int} / 7**")

            hi_cols = st.columns([0.35, 1.45, 2.1, 2.1, 0.45])
            hi_cols[0].markdown("**#**")
            hi_cols[1].markdown("**Fecha del Intercambio**")
            hi_cols[2].markdown(f"**Programado Operador 1 ({cod_op1})**")
            hi_cols[3].markdown(f"**Programado Operador 2 ({cod_op2})**")
            hi_cols[4].markdown("**Quitar**")

            for idx in range(st.session_state.num_filas_int):
                ri_cols = st.columns([0.35, 1.45, 2.1, 2.1, 0.45])
                ri_cols[0].markdown(f"**{idx + 1}**")
                default_idx = 1 if (idx == 0 and len(opciones_fecha_sel) > 1) else 0
                sel_f_str = ri_cols[1].selectbox(
                    f"FI{idx+1}",
                    opciones_fecha_sel,
                    index=default_idx,
                    key=f"brt_int_fecha_{idx}",
                    label_visibility="collapsed",
                )

                if sel_f_str != "— Sin cambio —":
                    f_obj = mapa_fechas_label[sel_f_str]
                    prog1 = buscar_turno_rapido(cod_op1, f_obj)
                    prog2 = buscar_turno_rapido(cod_op2, f_obj)

                    ri_cols[2].info(
                        f"`{prog1['horario']}` | {prog1['sercon']} | {prog1['p_inicio']} → {prog1['p_fin']}"
                    )
                    ri_cols[3].info(
                        f"`{prog2['horario']}` | {prog2['sercon']} | {prog2['p_inicio']} → {prog2['p_fin']}"
                    )

                    filas_prog_op1.append(
                        {
                            "fecha": f_obj.strftime("%d/%m/%Y"),
                            "fecha_corta": fecha_corta_es(f_obj),
                            "codigo": cod_op1,
                            **prog1,
                        }
                    )
                    filas_prog_op2.append(
                        {
                            "fecha": f_obj.strftime("%d/%m/%Y"),
                            "fecha_corta": fecha_corta_es(f_obj),
                            "codigo": cod_op2,
                            **prog2,
                        }
                    )
                    filas_cambio_op1.append(
                        {
                            "fecha": f_obj.strftime("%d/%m/%Y"),
                            "fecha_corta": fecha_corta_es(f_obj),
                            "codigo": cod_op1,
                            "horario": prog2["horario"],
                            "sercon": prog2["sercon"],
                            "p_inicio": prog2["p_inicio"],
                            "p_fin": prog2["p_fin"],
                        }
                    )
                    filas_cambio_op2.append(
                        {
                            "fecha": f_obj.strftime("%d/%m/%Y"),
                            "fecha_corta": fecha_corta_es(f_obj),
                            "codigo": cod_op2,
                            "horario": prog1["horario"],
                            "sercon": prog1["sercon"],
                            "p_inicio": prog1["p_inicio"],
                            "p_fin": prog1["p_fin"],
                        }
                    )
                else:
                    ri_cols[2].caption("—")
                    ri_cols[3].caption("—")

                if ri_cols[4].button("🗑️", key=f"del_int_{idx}", help="Eliminar fila"):
                    if st.session_state.num_filas_int > 1:
                        for j in range(idx, st.session_state.num_filas_int - 1):
                            if f"brt_int_fecha_{j+1}" in st.session_state:
                                st.session_state[f"brt_int_fecha_{j}"] = st.session_state[f"brt_int_fecha_{j+1}"]
                        st.session_state.pop(f"brt_int_fecha_{st.session_state.num_filas_int - 1}", None)
                        st.session_state.num_filas_int -= 1
                    else:
                        st.session_state["brt_int_fecha_0"] = "— Sin cambio —"
                    st.rerun()

        st.markdown("---")
        if st.button(
            "✅ Guardar en Google Sheets y Generar Boleta F-ZEN-GO-07",
            type="primary",
            use_container_width=True,
        ):
            fechas_selec = [f["fecha"] for f in filas_cambio_op1]
            if not cod_op1 or (es_intercambio and not cod_op2):
                st.error("⚠️ Debes ingresar el código del operador antes de guardar.")
            elif es_intercambio and cod_op1 == cod_op2:
                st.error("⚠️ El Operador 1 y el Operador 2 deben ser distintos.")
            elif not filas_cambio_op1:
                st.error("⚠️ Selecciona al menos 1 fecha válida en la tabla para registrar el cambio.")
            elif len(fechas_selec) != len(set(fechas_selec)):
                st.warning("⚠️ Hay fechas repetidas en las filas. Elimina o ajusta la fila duplicada.")
            else:
                codigo_registro = (
                    f"{cod_op1}_{cod_op2}" if es_intercambio else str(cod_op1)
                )
                fecha_cambio_principal = filas_cambio_op1[0]["fecha"]
                hora_cambio_resumen = (
                    f"{filas_prog_op1[0]['horario']}/{filas_cambio_op1[0]['horario']}"
                    if len(filas_cambio_op1) == 1
                    else f"{filas_prog_op1[0]['horario']}/{filas_cambio_op1[0]['horario']} (+{len(filas_cambio_op1) - 1} día{'s' if len(filas_cambio_op1) > 2 else ''})"
                )

                df_h_actual = st.session_state.historial_cambios
                ya_existe = False
                if not df_h_actual.empty and "COGIDO" in df_h_actual.columns:
                    fecha_cmp_date = parsear_fecha_segura(fecha_cambio_principal)
                    fechas_hist = df_h_actual["FECHA DE CAMBIO"].apply(parsear_fecha_segura)
                    dups = df_h_actual[
                        (df_h_actual["COGIDO"].astype(str).str.strip() == codigo_registro)
                        & (fechas_hist == fecha_cmp_date)
                        & (df_h_actual["HORA DE CAMBIO"].astype(str).str.strip() == hora_cambio_resumen)
                    ]
                    ya_existe = not dups.empty

                if ya_existe:
                    st.warning(
                        "⚠️ Este cambio ya está registrado en el historial (duplicado bloqueado). Si necesitas corregirlo, elimínalo desde la pestaña 3."
                    )
                else:
                    ref_num = (fecha_solicitud - date(1899, 12, 30)).days
                    fecha_cambio_obj = parsear_fecha_segura(fecha_cambio_principal)
                    nuevo_registro = {
                        "REF": ref_num,
                        "COGIDO": codigo_registro,
                        "FECHA DE SOLICITUD": fecha_solicitud.strftime("%Y-%m-%d"),
                        "FECHA DE CAMBIO": fecha_cambio_obj.strftime("%Y-%m-%d") if fecha_cambio_obj else fecha_cambio_principal,
                        "HORA DE CAMBIO": hora_cambio_resumen,
                        "RESPONSABLE DEL CAMBIO": responsable,
                        "OBSERVACIONES": motivo_final,
                        "Stamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    }

                    ok_gs, msg_gs = agregar_fila_en_gsheets(nuevo_registro)
                    if ok_gs:
                        df_gs_nuevo, ok_recarga = leer_historial_desde_gsheets()
                        if ok_recarga:
                            st.session_state.historial_cambios = df_gs_nuevo
                        else:
                            st.session_state.historial_cambios = pd.concat(
                                [pd.DataFrame([nuevo_registro]), df_h_actual],
                                ignore_index=True,
                            )
                        st.success(f"☁️ {msg_gs}")
                    else:
                        st.session_state.historial_cambios = pd.concat(
                            [pd.DataFrame([nuevo_registro]), df_h_actual],
                            ignore_index=True,
                        )
                        st.warning(f"💾 Guardado en sesión local. ({msg_gs})")

                    st.session_state.boleta_activa = {
                        "modalidad": modalidad,
                        "es_intercambio": es_intercambio,
                        "fecha_sol": fecha_solicitud.strftime("%d/%m/%Y"),
                        "fecha_sol_larga": fecha_larga_es(fecha_solicitud),
                        "patio": patio_op1,
                        "responsable": responsable,
                        "motivo": motivo_final,
                        "cod_1": cod_op1,
                        "nombre_1": nom_op1,
                        "cod_2": cod_op2,
                        "nombre_2": nom_op2,
                        "prog_op1": filas_prog_op1,
                        "cambio_op1": filas_cambio_op1,
                        "prog_op2": filas_prog_op2,
                        "cambio_op2": filas_cambio_op2,
                    }
                    st.info("📄 ¡Boleta F-ZEN-GO-07 lista! Ve a la **pestaña 2** para descargarla o imprimirla.")

# ------------------------------------------
# TAB 2: VISTA PREVIA Y PDF F-ZEN-GO-07
# ------------------------------------------
with tab_boleta:
    if "boleta_activa" not in st.session_state or st.session_state.boleta_activa is None:
        st.info("Genera primero un cambio en la pestaña **1. Módulo Cambio de Horario** para visualizar la boleta PDF.")
    else:
        bol = st.session_state.boleta_activa
        pdf_bytes = generar_pdf_f_zen_go_07(bol)

        c_btn1, c_btn2 = st.columns([1.6, 2])
        with c_btn1:
            st.download_button(
                label="📄 Descargar Boleta Oficial F-ZEN-GO-07 (PDF)",
                data=pdf_bytes,
                file_name=f"F-ZEN-GO-07_{bol['cod_1']}_{bol['fecha_sol'].replace('/', '')}.pdf",
                mime="application/pdf",
                type="primary",
                use_container_width=True,
            )
        with c_btn2:
            st.info(
                "💡 El PDF incluye las **3 secciones en una sola página** (Original Planilla, Copia Operador 1 y Copia Operador 2)."
            )

        st.markdown("---")
        st.markdown(
            f"""
            <div class="boleta-preview">
                <h4 style="margin:0; text-align:center; color:#0070c0;">CAMBIO DE HORARIO — F-ZEN-GO-07 / V1-12-2023</h4>
                <p style="margin:4px 0; font-size:13px;">
                    <b>Operador 1:</b> {bol['nombre_1']} (<span style="background:#ffff00; padding:1px 4px;"><b>{bol['cod_1']}</b></span>) &nbsp;|&nbsp;
                    <b>Patio:</b> {bol['patio']} &nbsp;|&nbsp;
                    <b>Fecha Solicitud:</b> {bol['fecha_sol_larga']}
                </p>
                <p style="margin:4px 0; font-size:13px;">
                    <b>Modalidad:</b> {bol['modalidad']} &nbsp;|&nbsp;
                    <b>Motivo:</b> {bol['motivo']} &nbsp;|&nbsp;
                    <b>Analista:</b> {bol['responsable']}
                </p>
            </div>
            """,
            unsafe_allow_html=True,
        )

        col_v1, col_v2 = st.columns(2)
        with col_v1:
            st.markdown("##### Horario Programado Operador 1")
            st.dataframe(
                pd.DataFrame(bol["prog_op1"])[["fecha_corta", "codigo", "horario", "sercon", "p_inicio"]],
                use_container_width=True,
                hide_index=True,
            )
            st.markdown("##### Horario de Cambio Operador 1")
            st.dataframe(
                pd.DataFrame(bol["cambio_op1"])[["fecha_corta", "codigo", "horario", "sercon", "p_inicio", "p_fin"]],
                use_container_width=True,
                hide_index=True,
            )
        with col_v2:
            if bol["es_intercambio"]:
                st.markdown(f"##### Horario Programado Operador 2 ({bol['nombre_2']})")
                st.dataframe(
                    pd.DataFrame(bol["prog_op2"])[["fecha_corta", "codigo", "horario", "sercon", "p_inicio"]],
                    use_container_width=True,
                    hide_index=True,
                )
                st.markdown(f"##### Horario de Cambio Operador 2 ({bol['nombre_2']})")
                st.dataframe(
                    pd.DataFrame(bol["cambio_op2"])[["fecha_corta", "codigo", "horario", "sercon", "p_inicio", "p_fin"]],
                    use_container_width=True,
                    hide_index=True,
                )

# ------------------------------------------
# TAB 3: HISTORIAL ACUMULADO, ELIMINACIÓN Y EXCEL
# ------------------------------------------
with tab_historial:
    st.subheader("📋 Historial Acumulado de Cambios")
    df_h_show = st.session_state.historial_cambios.copy()

    col_f1, col_f2 = st.columns(2)
    with col_f1:
        filtro_op = st.text_input("Buscar por Código de Operador:", "")
    with col_f2:
        filtro_resp = st.multiselect(
            "Filtrar por Responsable:",
            options=sorted(
                df_h_show["RESPONSABLE DEL CAMBIO"].dropna().astype(str).unique().tolist()
            )
            if not df_h_show.empty and "RESPONSABLE DEL CAMBIO" in df_h_show.columns
            else [],
        )

    df_filtrado = df_h_show.copy()
    if filtro_op.strip() and "COGIDO" in df_filtrado.columns:
        df_filtrado = df_filtrado[
            df_filtrado["COGIDO"].astype(str).str.contains(filtro_op.strip(), case=False, na=False)
        ]
    if filtro_resp and "RESPONSABLE DEL CAMBIO" in df_filtrado.columns:
        df_filtrado = df_filtrado[
            df_filtrado["RESPONSABLE DEL CAMBIO"].astype(str).isin(filtro_resp)
        ]

    st.caption(f"Mostrando **{len(df_filtrado)}** de **{len(df_h_show)}** registros acumulados.")

    if not df_filtrado.empty:
        df_con_check = df_filtrado.copy()
        df_con_check.insert(0, "Seleccionar", False)
        cols_ocultas = ["_GS_ROW"] if "_GS_ROW" in df_con_check.columns else []
        df_editado_hist = st.data_editor(
            df_con_check.drop(columns=cols_ocultas),
            use_container_width=True,
            hide_index=True,
            disabled=[c for c in df_con_check.columns if c not in ("Seleccionar", "_GS_ROW")],
            key="editor_seleccion_historial",
        )

        indices_borrar = df_filtrado.index[df_editado_hist["Seleccionar"]].tolist()
        if indices_borrar:
            if st.button(
                f"🗑️ Eliminar {len(indices_borrar)} registro(s) seleccionado(s) del Historial",
                type="secondary",
            ):
                filas_gs = (
                    df_h_show.loc[indices_borrar, "_GS_ROW"].dropna().tolist()
                    if "_GS_ROW" in df_h_show.columns
                    else []
                )
                if filas_gs and st.session_state.historial_en_gsheets:
                    ok_del, msg_del = eliminar_filas_en_gsheets(filas_gs)
                    if ok_del:
                        df_gs_new, _ = leer_historial_desde_gsheets()
                        st.session_state.historial_cambios = df_gs_new
                        st.success(f"☁️ {msg_del}")
                    else:
                        st.error(msg_del)
                else:
                    st.session_state.historial_cambios = df_h_show.drop(index=indices_borrar).reset_index(drop=True)
                    st.success(f"🗑️ Se eliminaron {len(indices_borrar)} registro(s) de la sesión.")
                st.rerun()
    else:
        st.info("No hay registros que coincidan con el filtro actual.")

    cols_export = [c for c in df_h_show.columns if c != "_GS_ROW"]
    buffer_xlsx = io.BytesIO()
    with pd.ExcelWriter(buffer_xlsx, engine="openpyxl") as writer:
        df_h_show[cols_export].to_excel(
            writer, sheet_name="HISTORIAL DE CAMBIOS (2)", index=False
        )
        pd.DataFrame({"NOMBRE": st.session_state.lista_responsables}).to_excel(
            writer, sheet_name="Lista", index=False
        )
    buffer_xlsx.seek(0)

    st.download_button(
        label="📥 Descargar Historial Completo Actualizado (Excel .xlsx)",
        data=buffer_xlsx,
        file_name=f"HISTORIAL_CAMBIOS_HORARIO_{date.today().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

# ------------------------------------------
# TAB 4: ESTADÍSTICAS ANUALES (3 GRÁFICAS AJUSTADAS)
# ------------------------------------------
with tab_estadisticas:
    st.subheader("📈 Historial Acumulado y Estadísticas Anuales")
    df_stats_raw = st.session_state.historial_cambios.copy()

    if df_stats_raw.empty:
        st.info("No hay registros en el historial para generar estadísticas.")
    else:
        f_cam_s = (
            df_stats_raw["FECHA DE CAMBIO"].apply(parsear_fecha_segura)
            if "FECHA DE CAMBIO" in df_stats_raw.columns
            else pd.Series([None] * len(df_stats_raw))
        )
        f_sol_s = (
            df_stats_raw["FECHA DE SOLICITUD"].apply(parsear_fecha_segura)
            if "FECHA DE SOLICITUD" in df_stats_raw.columns
            else pd.Series([None] * len(df_stats_raw))
        )
        df_stats_raw["_fecha_dt"] = pd.to_datetime(f_cam_s.fillna(f_sol_s), errors="coerce")
        df_stats_valid = df_stats_raw[df_stats_raw["_fecha_dt"].notna()].copy()
        df_stats_valid["_anio"] = df_stats_valid["_fecha_dt"].dt.year
        df_stats_valid["_mes"] = df_stats_valid["_fecha_dt"].dt.month
        df_stats_valid["_dia"] = df_stats_valid["_fecha_dt"].dt.day

        hoy = date.today()
        anios_disp = sorted(
            list({hoy.year} | set(df_stats_valid["_anio"].dropna().astype(int).tolist())),
            reverse=True,
        )
        anio_sel = st.selectbox("Año de análisis:", options=anios_disp, index=0, key="stats_anio_sel")

        df_anio = df_stats_valid[df_stats_valid["_anio"] == anio_sel].copy()

        ops_unicos = set()
        if not df_anio.empty and "COGIDO" in df_anio.columns:
            for c_val in df_anio["COGIDO"].astype(str):
                for parte in c_val.split("_"):
                    if parte.strip():
                        ops_unicos.add(parte.strip())

        cambios_mes_actual = int((df_anio["_mes"] == hoy.month).sum()) if not df_anio.empty else 0

        k1, k2, k3, k4 = st.columns(4)
        k1.metric(f"Total Cambios ({anio_sel})", f"{len(df_anio):,}")
        k2.metric("Total Acumulado Histórico", f"{len(df_stats_raw):,}")
        k3.metric(f"Operadores con Cambios ({anio_sel})", f"{len(ops_unicos):,}")
        k4.metric(f"Cambios en {MESES_ES[hoy.month]}", f"{cambios_mes_actual:,}")

        st.markdown("---")

        # 1. GRÁFICA PRINCIPAL (MÁS GRANDE, ANCHO COMPLETO): ACUMULADO ANUAL POR MES
        c_hdr1, c_hdr2 = st.columns([2.5, 1.2])
        with c_hdr1:
            st.markdown(f"#### Acumulado Anual por Mes ({anio_sel})")
        with c_hdr2:
            vista_anual = st.radio(
                "Vista anual",
                ["Acumulado Anual", "Por Mes"],
                horizontal=True,
                label_visibility="collapsed",
                key="stats_vista_anual",
            )

        df_meses = pd.DataFrame({"mes_num": range(1, 13)})
        df_meses["Mes"] = df_meses["mes_num"].apply(lambda m: f"{m:02d}. {MESES_ES[m]}")
        conteo_mes = df_anio.groupby("_mes").size().to_dict() if not df_anio.empty else {}
        df_meses["Cambios del Mes"] = df_meses["mes_num"].map(lambda m: conteo_mes.get(m, 0))
        df_meses["Acumulado Anual"] = df_meses["Cambios del Mes"].cumsum()

        if anio_sel == hoy.year:
            ultimo_mes_activo = max([hoy.month] + [m for m, c in conteo_mes.items() if c > 0])
            df_meses.loc[df_meses["mes_num"] > ultimo_mes_activo, "Acumulado Anual"] = 0

        col_serie = "Acumulado Anual" if vista_anual == "Acumulado Anual" else "Cambios del Mes"
        st.bar_chart(
            df_meses.set_index("Mes")[[col_serie]],
            color="#0070c0",
            height=340,
            use_container_width=True,
        )

        # 2 y 3. GRÁFICAS INFERIORES: DÍAS DEL MES ACTUAL Y TIPO DE CAMBIO
        col_g_dias, col_g_tipo = st.columns(2)

        with col_g_dias:
            mes_graf_sel = st.selectbox(
                "Mes a visualizar por día:",
                options=list(range(1, 13)),
                index=hoy.month - 1,
                format_func=lambda m: f"{MESES_LARGO_ES[m].capitalize()} {anio_sel}",
                key="stats_mes_diario",
            )
            dias_en_mes = calendar.monthrange(anio_sel, mes_graf_sel)[1]
            df_dias = pd.DataFrame({"Dia": range(1, dias_en_mes + 1)})

            if not df_anio.empty:
                df_mes_filtrado = df_anio[df_anio["_mes"] == mes_graf_sel]
                conteo_dias = df_mes_filtrado.groupby("_dia").size().to_dict()
            else:
                conteo_dias = {}

            df_dias["Cambios"] = df_dias["Dia"].map(lambda d: conteo_dias.get(d, 0))
            df_dias["Día"] = df_dias["Dia"].apply(lambda d: f"{d:02d}")
            st.markdown(f"#### Cambios por Día ({MESES_ES[mes_graf_sel]} {anio_sel})")
            st.bar_chart(
                df_dias.set_index("Día")[["Cambios"]],
                color="#00b050",
                use_container_width=True,
            )

        with col_g_tipo:
            st.markdown(f"#### Distribución por Tipo de Cambio ({anio_sel})")
            if not df_anio.empty:
                df_anio["_tipo_cambio"] = df_anio.apply(clasificar_tipo_cambio_row, axis=1)
                conteo_tipo = (
                    df_anio["_tipo_cambio"]
                    .value_counts()
                    .reindex(
                        ["Cambio de horario", "Cambio con otro operador", "Cambio de día libre"],
                        fill_value=0,
                    )
                    .reset_index()
                )
                conteo_tipo.columns = ["Tipo de Cambio", "Cambios"]
                st.bar_chart(
                    conteo_tipo.set_index("Tipo de Cambio")[["Cambios"]],
                    color="#0070c0",
                    use_container_width=True,
                )
            else:
                st.info(f"Sin registros en {anio_sel} para clasificar por tipo de cambio.")

        # Ranking de operadores con más cambios en el año
        st.markdown(f"#### Operadores con Más Cambios en {anio_sel}")
        if not df_anio.empty and "COGIDO" in df_anio.columns:
            conteo_ops = {}
            for c_val in df_anio["COGIDO"].astype(str):
                for parte in c_val.split("_"):
                    p = parte.strip()
                    if p:
                        conteo_ops[p] = conteo_ops.get(p, 0) + 1
            df_top = (
                pd.DataFrame(list(conteo_ops.items()), columns=["Operador", "Cambios"])
                .sort_values("Cambios", ascending=False)
                .head(12)
            )
            st.dataframe(df_top, use_container_width=True, hide_index=True)

# ------------------------------------------
# TAB 5: CUADRÍCULA INTERACTIVA TIPO EXCEL (CTRL + V)
# ------------------------------------------
with tab_pegar_tabla:
    st.subheader("🗂️ Cuadrícula Tipo Excel — Pega o Edita Horarios Directamente")
    st.markdown(
        "Haz clic en la primera celda (`Fecha`), presiona **`Ctrl + V`** para pegar las filas copiadas desde Excel y luego pulsa **Fijar programación**."
    )

    if not st.session_state.df_programacion_raw.empty:
        df_base_editor = st.session_state.df_programacion_raw[
            ["Fecha", "Tarea", "CODIGO", "NOMBRE", "H.Inicial", "origen", "H. Final", "Patio Final", "PADMIN"]
        ].copy()
    else:
        df_base_editor = pd.DataFrame(
            [
                {
                    "Fecha": "",
                    "Tarea": "",
                    "CODIGO": "",
                    "NOMBRE": "",
                    "H.Inicial": "",
                    "origen": "",
                    "H. Final": "",
                    "Patio Final": "",
                    "PADMIN": "",
                }
                for _ in range(15)
            ]
        )

    df_editor = st.data_editor(
        df_base_editor,
        num_rows="dynamic",
        use_container_width=True,
        key="editor_excel_grid",
    )

    if st.button(
        "💾 Fijar programación de esta cuadrícula en sesión",
        type="primary",
    ):
        df_norm_grid = normalizar_tabla_programacion(
            df_editor, excluir_a_orden=st.session_state.excluir_a_orden
        )
        if not df_norm_grid.empty:
            indexar_programacion_en_sesion(
                df_norm_grid,
                f"Cuadrícula Excel ({len(df_norm_grid)} turnos - {datetime.now().strftime('%H:%M')})",
            )
            st.success(
                f"✅ Se cargaron **{len(df_norm_grid)}** turnos válidos desde la cuadrícula. Ya puedes ir a la pestaña 1."
            )
            st.rerun()
        else:
            st.warning(
                "⚠️ Asegúrate de que al menos las columnas **Fecha** y **CODIGO** tengan datos válidos antes de guardar."
            )
