from datetime import datetime, date, time, timedelta
import io
import re
import unicodedata
import altair as alt
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
# CATÁLOGOS OFICIALES
# ==========================================
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

OPCIONES_JORNADA_FILA = [
    "Full Time (8:40)",
    "Part Time (6:30)",
    "Catálogo",
    "Libre (-)",
]

MINUTOS_FULL_TIME = 8 * 60 + 40  # 520 min (8h 40m)
MINUTOS_PART_TIME = 6 * 60 + 30  # 390 min (6h 30m)

PLACEHOLDER_RESP = "— Seleccionar responsable —"
PLACEHOLDER_MOTIVO = "— Seleccionar razón / autorización —"

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


def limpiar_codigo(val):
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def calcular_horario_desde_inicio(hora_ini: time, tipo_jornada: str, horario_cat: str = "-") -> str:
    if tipo_jornada == "Libre (-)":
        return "-"
    if tipo_jornada == "Catálogo":
        return horario_cat or "-"
    if hora_ini is None:
        return "-"
    inicio_min = hora_ini.hour * 60 + hora_ini.minute
    duracion = MINUTOS_PART_TIME if tipo_jornada == "Part Time (6:30)" else MINUTOS_FULL_TIME
    fin_min = inicio_min + duracion
    fin_hh = fin_min // 60
    fin_mm = fin_min % 60
    return f"{hora_ini.hour:02d}:{hora_ini.minute:02d}-{fin_hh:02d}:{fin_mm:02d}"


def es_turno_a_orden(*campos) -> bool:
    for c in campos:
        if c is None or (isinstance(c, float) and pd.isna(c)):
            continue
        s = str(c).strip()
        if not s:
            continue
        norm = unicodedata.normalize("NFD", s)
        norm = "".join(ch for ch in norm if unicodedata.category(ch) != "Mn")
        norm = re.sub(r"[_-]+", " ", norm.upper())
        norm = re.sub(r"\s+", " ", norm).strip()
        if (
            norm == "A ORDEN"
            or norm == "A LA ORDEN"
            or "A ORDEN" in norm
            or "A LA ORDEN" in norm
            or norm == "ORDEN"
        ):
            return True
    return False


def detectar_dayfirst(serie: pd.Series) -> bool:
    only_dmy = 0
    only_mdy = 0
    parsed_pairs = []

    for val in serie.dropna():
        if isinstance(val, (datetime, date, pd.Timestamp)):
            continue
        s = str(val).strip()
        m = re.match(r"^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})(?:\s|$)", s)
        if not m:
            continue
        p1, p2, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if y < 100:
            y += 2000
        if p1 > 12 and 1 <= p2 <= 12:
            only_dmy += 1
        elif p2 > 12 and 1 <= p1 <= 12:
            only_mdy += 1
        parsed_pairs.append((p1, p2, y))

    if only_mdy > 0 and only_dmy == 0:
        return False
    if only_dmy > 0 and only_mdy == 0:
        return True
    if not parsed_pairs:
        return True

    dmy_dates, mdy_dates = [], []
    for p1, p2, y in parsed_pairs:
        if 1 <= p2 <= 12 and 1 <= p1 <= 31:
            try:
                dmy_dates.append(date(y, p2, p1))
            except ValueError:
                pass
        if 1 <= p1 <= 12 and 1 <= p2 <= 31:
            try:
                mdy_dates.append(date(y, p1, p2))
            except ValueError:
                pass

    if len(dmy_dates) >= 2 and len(mdy_dates) >= 2:
        span_dmy = (max(dmy_dates) - min(dmy_dates)).days
        span_mdy = (max(mdy_dates) - min(mdy_dates)).days
        if span_mdy < span_dmy:
            return False
        if span_dmy < span_mdy:
            return True

    return True


def parsear_fecha_segura(val, dayfirst: bool = True):
    if pd.isna(val) or val is None or str(val).strip() == "":
        return pd.NaT
    if isinstance(val, pd.Timestamp):
        return val.normalize()
    if isinstance(val, datetime):
        return pd.Timestamp(val.date())
    if isinstance(val, date):
        return pd.Timestamp(val)

    if isinstance(val, (int, float)) and not isinstance(val, bool) and val > 20000:
        try:
            return pd.to_datetime(val, unit="D", origin="1899-12-30").normalize()
        except Exception:
            pass

    s = str(val).strip()
    m_iso = re.match(r"^(\d{4})[-/](\d{1,2})[-/](\d{1,2})", s)
    if m_iso:
        try:
            return pd.Timestamp(date(int(m_iso.group(1)), int(m_iso.group(2)), int(m_iso.group(3))))
        except ValueError:
            return pd.NaT

    m_slash = re.match(r"^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})", s)
    if m_slash:
        p1, p2, y = int(m_slash.group(1)), int(m_slash.group(2)), int(m_slash.group(3))
        if y < 100:
            y += 2000
        if p1 > 12 and 1 <= p2 <= 12:
            d, m = p1, p2
        elif p2 > 12 and 1 <= p1 <= 12:
            m, d = p1, p2
        elif dayfirst:
            d, m = p1, p2
        else:
            m, d = p1, p2
        try:
            return pd.Timestamp(date(y, m, d))
        except ValueError:
            return pd.NaT

    return pd.to_datetime(s, dayfirst=dayfirst, errors="coerce")


def reparar_fechas_invertidas_excel(serie_fechas: pd.Series) -> pd.Series:
    """Detecta fechas nativas de Excel leídas con mes/día invertidos (ej. 12/01..12/06 leídas como 2026-01-12..2026-06-12)."""
    fechas_validas = serie_fechas.dropna()
    if len(fechas_validas) < 2:
        return serie_fechas

    unicas = sorted(fechas_validas.unique())
    if len(unicas) < 2:
        return serie_fechas

    span_orig = (unicas[-1] - unicas[0]).days
    if span_orig <= 35:
        return serie_fechas

    swapped = {}
    for d in unicas:
        if 1 <= d.day <= 12:
            try:
                swapped[d] = date(d.year, d.day, d.month)
            except ValueError:
                return serie_fechas
        else:
            swapped[d] = d

    vals_swap = sorted(set(swapped.values()))
    span_swap = (vals_swap[-1] - vals_swap[0]).days
    if span_swap <= 35 and span_swap < span_orig * 0.6:
        return serie_fechas.map(lambda x: swapped.get(x, x) if pd.notna(x) else x)
    return serie_fechas


def clasificar_tipo_cambio(row: pd.Series) -> str:
    cod = str(row.get("COGIDO", "")).strip()
    obs = str(row.get("OBSERVACIONES", "")).strip().upper()
    hora = str(row.get("HORA DE CAMBIO", "")).strip()

    if "_" in cod or "INTERCAMBIO" in obs or "CON OTRO OPERADOR" in obs:
        return "Cambio con otro operador"

    base_hora = hora.split(" (+")[0].strip()
    partes = base_hora.split("/")
    destino = partes[1].strip().upper() if len(partes) >= 2 else ""

    if (
        destino in ["-", "DESCANSO", "LIBRE"]
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
    try:
        ws = obtener_hoja_google_sheets()
        if ws is None:
            return pd.DataFrame(columns=COLUMNAS_HISTORIAL), False
        records = ws.get_all_records()
        if not records:
            return pd.DataFrame(columns=COLUMNAS_HISTORIAL), True
        df_h = pd.DataFrame(records)
        df_h.columns = [str(c).strip() for c in df_h.columns]
        if "COGIDO" in df_h.columns:
            df_h = df_h[df_h["COGIDO"].astype(str).str.strip() != ""].copy()
            df_h["COGIDO"] = df_h["COGIDO"].apply(limpiar_codigo)
        for col_f in ["FECHA DE SOLICITUD", "FECHA DE CAMBIO"]:
            if col_f in df_h.columns:
                df_flag = detectar_dayfirst(df_h[col_f])
                df_h[col_f] = df_h[col_f].apply(lambda x: parsear_fecha_segura(x, dayfirst=df_flag))
        return df_h, True
    except Exception:
        return pd.DataFrame(columns=COLUMNAS_HISTORIAL), False


def agregar_fila_en_gsheets(fila_dict: dict):
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


def eliminar_fila_en_gsheets(codigo: str, fecha_cambio_iso: str, hora_cambio: str):
    try:
        ws = obtener_hoja_google_sheets()
        if ws is None:
            return False, "No hay conexión activa a Google Sheets."
        all_vals = ws.get_all_values()
        if len(all_vals) <= 1:
            return False, "La hoja está vacía."
        headers = [h.strip() for h in all_vals[0]]
        idx_cod = headers.index("COGIDO") if "COGIDO" in headers else 1
        idx_fc = headers.index("FECHA DE CAMBIO") if "FECHA DE CAMBIO" in headers else 3
        idx_hc = headers.index("HORA DE CAMBIO") if "HORA DE CAMBIO" in headers else 4

        for row_num in range(2, len(all_vals) + 1):
            row = all_vals[row_num - 1]
            c_val = limpiar_codigo(row[idx_cod]) if idx_cod < len(row) else ""
            f_val = row[idx_fc].strip() if idx_fc < len(row) else ""
            h_val = row[idx_hc].strip() if idx_hc < len(row) else ""
            f_parsed = parsear_fecha_segura(f_val, dayfirst=True)
            f_norm = f_parsed.strftime("%Y-%m-%d") if pd.notna(f_parsed) else f_val[:10]
            if c_val == limpiar_codigo(codigo) and f_norm == fecha_cambio_iso and h_val == str(hora_cambio).strip():
                ws.delete_rows(row_num)
                return True, "Registro eliminado de Google Sheets."
        return False, "No se encontró coincidencia exacta en Google Sheets."
    except Exception as e:
        return False, f"Error al eliminar en Google Sheets: {e}"


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
        if val.year in (1899, 1900):
            diff_days = (val.date() - date(1899, 12, 30)).days
            if 0 <= diff_days < 5:
                total_min = diff_days * 24 * 60 + val.hour * 60 + val.minute
                return f"{total_min // 60:02d}:{total_min % 60:02d}"
        return val.strftime("%H:%M")
    if isinstance(val, (int, float)) and not isinstance(val, bool) and 0 <= val < 5:
        total_min = int(round(val * 24 * 60))
        return f"{total_min // 60:02d}:{total_min % 60:02d}"
    s = str(val).strip()
    m1900 = re.search(r"1900-01-0(\d)[T\s](\d{1,2}):(\d{2})", s)
    if m1900:
        hh = int(m1900.group(1)) * 24 + int(m1900.group(2))
        mm = int(m1900.group(3))
        return f"{hh:02d}:{mm:02d}"
    if len(s) >= 5 and ":" in s:
        return s[:5]
    return s


def normalizar_tabla_programacion(df_raw: pd.DataFrame) -> tuple:
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
        elif cu in ["H.INICIAL", "HORA INICIAL", "HORA I"]:
            col_map[c] = "H.Inicial"
        elif cu in ["ORIGEN", "PATIO ORIGEN", "P. INICIO"]:
            col_map[c] = "origen"
        elif cu in ["H. FINAL", "H.FINAL", "HORA FIN", "HORA FINAL"]:
            col_map[c] = "H. Final"
        elif cu in ["PATIO FINAL", "P. FINAL"]:
            col_map[c] = "Patio Final"
        elif cu in ["PADMIN", "PATIO"]:
            col_map[c] = "PADMIN"
        elif "UNNAMED" in cu or "HORARIO" in cu:
            col_map[c] = "HORARIO_RAW"

    df = df.rename(columns=col_map)
    if "CODIGO" not in df.columns or "Fecha" not in df.columns:
        return pd.DataFrame(), 0

    df["CODIGO_STR"] = df["CODIGO"].apply(limpiar_codigo)
    df = df[df["CODIGO_STR"] != ""].copy()

    es_dayfirst = detectar_dayfirst(df["Fecha"])
    fechas_parsed = df["Fecha"].apply(lambda x: parsear_fecha_segura(x, dayfirst=es_dayfirst))
    df["FECHA_DATE"] = pd.to_datetime(fechas_parsed, errors="coerce").dt.date
    df = df[df["FECHA_DATE"].notna()].copy()
    df["FECHA_DATE"] = reparar_fechas_invertidas_excel(df["FECHA_DATE"])

    if "HORARIO_RAW" in df.columns:
        df["HORARIO_STR"] = df["HORARIO_RAW"].fillna("-").astype(str).str.strip()
    else:
        hi = df["H.Inicial"].apply(formatear_hora_celda) if "H.Inicial" in df.columns else ""
        hf = df["H. Final"].apply(formatear_hora_celda) if "H. Final" in df.columns else ""
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

    mask_a_orden = df.apply(
        lambda r: es_turno_a_orden(
            r.get("Tarea"),
            r.get("HORARIO_STR"),
            r.get("origen"),
            r.get("Patio Final"),
        ),
        axis=1,
    )
    a_orden_eliminados = int(mask_a_orden.sum())
    df = df[~mask_a_orden].copy()

    return df, a_orden_eliminados


def indexar_programacion_en_sesion(df_norm: pd.DataFrame, nombre_fuente: str, a_orden_elim: int = 0):
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

    st.session_state.operadores_dict = operadores_dict
    st.session_state.turnos_dict = turnos_dict
    st.session_state.fechas_disponibles = fechas_unicas
    st.session_state.fecha_min = fechas_unicas[0] if fechas_unicas else None
    st.session_state.fecha_max = fechas_unicas[-1] if fechas_unicas else None
    st.session_state.prog_total_filas = len(df_norm)
    st.session_state.prog_a_orden_eliminados = a_orden_elim
    st.session_state.prog_nombre_archivo = nombre_fuente
    st.session_state.prog_cargada = len(df_norm) > 0
    st.session_state.mostrar_cargador = False


def parsear_texto_pegado_excel(texto_tsv: str):
    lineas = [l for l in texto_tsv.splitlines() if l.strip()]
    if not lineas:
        return pd.DataFrame(), 0

    df_tsv = pd.read_csv(io.StringIO("\n".join(lineas)), sep="\t", header=None)
    primera_fila_upper = [str(x).strip().upper() for x in df_tsv.iloc[0].tolist()]
    if any(k in primera_fila_upper for k in ["CODIGO", "CÓDIGO", "FECHA"]):
        df_tsv.columns = [str(x).strip() for x in df_tsv.iloc[0].tolist()]
        df_tsv = df_tsv.iloc[1:].reset_index(drop=True)
        return normalizar_tabla_programacion(df_tsv)

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
    return normalizar_tabla_programacion(df_tsv)


@st.cache_data(show_spinner="Indexando programación completa...")
def procesar_archivo_excel(file_bytes: bytes):
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
                for d_resp in RESPONSABLES_DEFAULT:
                    if not any(v.lower() == d_resp.lower() for v in vals):
                        vals.append(d_resp)
                lista_resp = vals

    df_prog = pd.DataFrame()
    a_orden_total = 0
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
        df_norm, a_ord = normalizar_tabla_programacion(df_try)
        if not df_norm.empty:
            df_prog = df_norm
            a_orden_total = a_ord
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
        if "COGIDO" in df_h.columns:
            df_h = df_h[df_h["COGIDO"].notna()].copy()
            df_h["COGIDO"] = df_h["COGIDO"].apply(limpiar_codigo)
        for col_f in ["FECHA DE SOLICITUD", "FECHA DE CAMBIO"]:
            if col_f in df_h.columns:
                df_flag = detectar_dayfirst(df_h[col_f])
                df_h[col_f] = df_h[col_f].apply(lambda x: parsear_fecha_segura(x, dayfirst=df_flag))
        df_hist = df_h

    return df_prog, df_hist, lista_resp, a_orden_total


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

    fc_flag = detectar_dayfirst(sub["FECHA DE CAMBIO"]) if "FECHA DE CAMBIO" in sub.columns else True
    fs_flag = detectar_dayfirst(sub["FECHA DE SOLICITUD"]) if "FECHA DE SOLICITUD" in sub.columns else True
    f_cambio = (
        sub["FECHA DE CAMBIO"].apply(lambda x: parsear_fecha_segura(x, dayfirst=fc_flag))
        if "FECHA DE CAMBIO" in sub.columns
        else pd.Series(pd.NaT, index=sub.index)
    )
    f_sol = (
        sub["FECHA DE SOLICITUD"].apply(lambda x: parsear_fecha_segura(x, dayfirst=fs_flag))
        if "FECHA DE SOLICITUD" in sub.columns
        else pd.Series(pd.NaT, index=sub.index)
    )
    fecha_ref = pd.to_datetime(f_cambio.fillna(f_sol), errors="coerce")

    sub["FECHA_REF"] = fecha_ref
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
# CALLBACKS PARA FILAS Y MODALIDAD
# ==========================================
def inicializar_estados_filas():
    if "brt_num_filas" not in st.session_state:
        st.session_state.brt_num_filas = 1
    if "brt_jornada_base" not in st.session_state:
        st.session_state.brt_jornada_base = "Full Time (8:40)"
    for i in range(7):
        if f"brt_ind_jornada_{i}" not in st.session_state:
            st.session_state[f"brt_ind_jornada_{i}"] = st.session_state.brt_jornada_base
        if f"brt_ind_hini_{i}" not in st.session_state:
            st.session_state[f"brt_ind_hini_{i}"] = time(5, 0)
        if f"brt_ind_tar_{i}" not in st.session_state:
            st.session_state[f"brt_ind_tar_{i}"] = "DISPONIBILIDAD"
        if f"brt_ind_pi_{i}" not in st.session_state:
            st.session_state[f"brt_ind_pi_{i}"] = PATIOS_CATALOGO[0]
        if f"brt_ind_pf_{i}" not in st.session_state:
            st.session_state[f"brt_ind_pf_{i}"] = PATIOS_CATALOGO[0]


def on_cambio_modalidad():
    mod = st.session_state.get("brt_modalidad", "Cambio de horario")
    j_base = st.session_state.get("brt_jornada_base", "Full Time (8:40)")
    for i in range(7):
        if mod == "Cambio de día libre":
            st.session_state[f"brt_ind_jornada_{i}"] = "Libre (-)"
            st.session_state[f"brt_ind_tar_{i}"] = "DESCANSO"
            st.session_state[f"brt_ind_pi_{i}"] = "-"
            st.session_state[f"brt_ind_pf_{i}"] = "-"
        else:
            if st.session_state.get(f"brt_ind_jornada_{i}") == "Libre (-)":
                st.session_state[f"brt_ind_jornada_{i}"] = j_base
            if st.session_state.get(f"brt_ind_tar_{i}") == "DESCANSO":
                st.session_state[f"brt_ind_tar_{i}"] = "DISPONIBILIDAD"
            if st.session_state.get(f"brt_ind_pi_{i}") == "-":
                st.session_state[f"brt_ind_pi_{i}"] = PATIOS_CATALOGO[0]
            if st.session_state.get(f"brt_ind_pf_{i}") == "-":
                st.session_state[f"brt_ind_pf_{i}"] = PATIOS_CATALOGO[0]


def on_cambio_jornada_base():
    j_base = st.session_state.get("brt_jornada_base", "Full Time (8:40)")
    for i in range(7):
        actual = st.session_state.get(f"brt_ind_jornada_{i}")
        if actual in ("Full Time (8:40)", "Part Time (6:30)"):
            st.session_state[f"brt_ind_jornada_{i}"] = j_base


def on_cambio_jornada_fila(idx: int):
    j_fila = st.session_state.get(f"brt_ind_jornada_{idx}", "Full Time (8:40)")
    if j_fila == "Libre (-)":
        st.session_state[f"brt_ind_tar_{idx}"] = "DESCANSO"
        st.session_state[f"brt_ind_pi_{idx}"] = "-"
        st.session_state[f"brt_ind_pf_{idx}"] = "-"
    else:
        if st.session_state.get(f"brt_ind_tar_{idx}") == "DESCANSO":
            st.session_state[f"brt_ind_tar_{idx}"] = "DISPONIBILIDAD"
        if st.session_state.get(f"brt_ind_pi_{idx}") == "-":
            st.session_state[f"brt_ind_pi_{idx}"] = PATIOS_CATALOGO[0]
        if st.session_state.get(f"brt_ind_pf_{idx}") == "-":
            st.session_state[f"brt_ind_pf_{idx}"] = PATIOS_CATALOGO[0]


def eliminar_fila_borrador(idx_borrar: int, prefix: str = "ind"):
    n_act = st.session_state.brt_num_filas
    if n_act <= 1:
        st.session_state[f"brt_{prefix}_fecha_0"] = "— Sin cambio —"
        if prefix == "ind":
            mod = st.session_state.get("brt_modalidad", "Cambio de horario")
            es_lib = mod == "Cambio de día libre"
            st.session_state["brt_ind_jornada_0"] = "Libre (-)" if es_lib else st.session_state.brt_jornada_base
            st.session_state["brt_ind_hini_0"] = time(5, 0)
            st.session_state["brt_ind_tar_0"] = "DESCANSO" if es_lib else "DISPONIBILIDAD"
            st.session_state["brt_ind_pi_0"] = "-" if es_lib else PATIOS_CATALOGO[0]
            st.session_state["brt_ind_pf_0"] = "-" if es_lib else PATIOS_CATALOGO[0]
        return

    for j in range(idx_borrar, n_act - 1):
        st.session_state[f"brt_{prefix}_fecha_{j}"] = st.session_state.get(
            f"brt_{prefix}_fecha_{j+1}", "— Sin cambio —"
        )
        if prefix == "ind":
            for campo in ["jornada", "hini", "hor", "tar", "pi", "pf"]:
                k_src = f"brt_ind_{campo}_{j+1}"
                k_dst = f"brt_ind_{campo}_{j}"
                if k_src in st.session_state:
                    st.session_state[k_dst] = st.session_state[k_src]

    ult = n_act - 1
    st.session_state[f"brt_{prefix}_fecha_{ult}"] = "— Sin cambio —"
    if prefix == "ind":
        mod = st.session_state.get("brt_modalidad", "Cambio de horario")
        es_lib = mod == "Cambio de día libre"
        st.session_state[f"brt_ind_jornada_{ult}"] = "Libre (-)" if es_lib else st.session_state.brt_jornada_base
        st.session_state[f"brt_ind_hini_{ult}"] = time(5, 0)
        st.session_state[f"brt_ind_tar_{ult}"] = "DESCANSO" if es_lib else "DISPONIBILIDAD"
        st.session_state[f"brt_ind_pi_{ult}"] = "-" if es_lib else PATIOS_CATALOGO[0]
        st.session_state[f"brt_ind_pf_{ult}"] = "-" if es_lib else PATIOS_CATALOGO[0]
    st.session_state.brt_num_filas = n_act - 1


# ==========================================
# INICIALIZAR ESTADO PERSISTENTE EN SESIÓN
# ==========================================
if "prog_cargada" not in st.session_state:
    st.session_state.prog_cargada = False
    st.session_state.prog_nombre_archivo = ""
    st.session_state.operadores_dict = {}
    st.session_state.turnos_dict = {}
    st.session_state.fechas_disponibles = []
    st.session_state.fecha_min = None
    st.session_state.fecha_max = None
    st.session_state.prog_total_filas = 0
    st.session_state.prog_a_orden_eliminados = 0
    st.session_state.lista_responsables = list(RESPONSABLES_DEFAULT)
    st.session_state.mostrar_cargador = True

inicializar_estados_filas()

if "historial_sincronizado_gs" not in st.session_state:
    df_gs, ok_gs = leer_historial_desde_gsheets()
    st.session_state.historial_cambios = df_gs
    st.session_state.historial_en_gsheets = ok_gs
    st.session_state.historial_sincronizado_gs = True


# ==========================================
# SIDEBAR: CARGA PERSISTENTE + GOOGLE SHEETS
# ==========================================
st.sidebar.header("📂 1. Programación Activa")

if st.session_state.prog_cargada and not st.session_state.mostrar_cargador:
    f_ini_txt = st.session_state.fecha_min.strftime("%d/%m/%Y")
    f_fin_txt = st.session_state.fecha_max.strftime("%d/%m/%Y")
    st.sidebar.success(
        f"📌 **Fuente activa:**\n{st.session_state.prog_nombre_archivo}\n\n"
        f"📅 **Fecha Inicial:** {f_ini_txt}\n"
        f"📅 **Fecha Final:** {f_fin_txt}\n\n"
        f"• **{st.session_state.prog_total_filas:,}** turnos válidos\n"
        f"• **{st.session_state.prog_a_orden_eliminados:,}** turnos 'A ORDEN' excluidos"
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
            df_prog_l, df_h_loaded, resp_l, a_ord_l = procesar_archivo_excel(
                archivo_subido.getvalue()
            )
            st.session_state.lista_responsables = resp_l
            if st.session_state.historial_cambios.empty and not df_h_loaded.empty:
                st.session_state.historial_cambios = df_h_loaded.copy()
            indexar_programacion_en_sesion(df_prog_l, archivo_subido.name, a_ord_l)
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
                df_pegado, a_ord_p = parsear_texto_pegado_excel(texto_pegado)
                if not df_pegado.empty:
                    indexar_programacion_en_sesion(
                        df_pegado,
                        f"Pegado desde Excel ({datetime.now().strftime('%H:%M')})",
                        a_ord_p,
                    )
                    st.rerun()
                else:
                    st.sidebar.error(
                        "⚠️ No se reconocieron las columnas. Asegúrate de copiar desde Fecha hasta Patio Final/PADMIN."
                    )
            else:
                st.sidebar.warning("⚠️ Pega primero los datos copiados de Excel.")

    if st.session_state.prog_cargada:
        if st.session_state.sidebar.button("Cancelar cambio", use_container_width=True) if False else st.sidebar.button("Cancelar cambio", use_container_width=True):
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
        "⚠️ Verifica `SHEET_ID` y `gcp_service_account` en Secrets de Streamlit Cloud."
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
    if not cod:
        return "Ingresa un código de operador", "—", False, True
    info = st.session_state.operadores_dict.get(cod)
    if info:
        return info["nombre"], info["patio"], True, False
    return "OPERADOR NO ENCONTRADO EN PROGRAMACIÓN", "LOS PUEBLOS", False, False


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
        "⚠️ **Aún no has cargado ninguna programación.** Sube el archivo Excel o pega los horarios en la barra lateral izquierda (o en la pestaña 5)."
    )

# ==========================================
# PESTAÑAS PRINCIPALES
# ==========================================
tab_registro, tab_boleta, tab_historial, tab_estadisticas, tab_pegar_tabla = st.tabs(
    [
        "📝 1. Módulo Cambio de Horario (BRT)",
        "🖨️ 2. Boleta Oficial F-ZEN-GO-07 (PDF)",
        "📊 3. Historial Acumulado y Descarga Excel",
        "📈 4. Estadísticas Anuales",
        "📋 5. Tabla Tipo Excel (Pegar / Editar)",
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
            key="brt_modalidad",
            on_change=on_cambio_modalidad,
        )
    with c_top2:
        fecha_solicitud = st.date_input("Fecha de Solicitud", value=date.today())
    with c_top3:
        opciones_resp = [PLACEHOLDER_RESP] + st.session_state.lista_responsables
        responsable = st.selectbox(
            "Responsable del Cambio",
            opciones_resp,
            index=0,
            key="brt_responsable",
        )
    with c_top4:
        opciones_motivo = [PLACEHOLDER_MOTIVO] + MOTIVOS_ESTANDAR
        motivo_sel = st.selectbox(
            "Razón / Autorización",
            opciones_motivo,
            index=0,
            key="brt_motivo",
        )

    motivo_extra = st.text_input(
        "Observación adicional (Opcional)",
        placeholder="Ej. Cita médica / Autoriza JOP / Cubre servicio...",
    )
    motivo_final = (
        f"{motivo_sel} - {motivo_extra.strip().upper()}"
        if motivo_sel != PLACEHOLDER_MOTIVO and motivo_extra.strip()
        else (motivo_sel if motivo_sel != PLACEHOLDER_MOTIVO else "")
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
                cod_op1 = limpiar_codigo(
                    st.text_input(
                        "Código Operador",
                        value="",
                        placeholder="Ej. 82601",
                        key="brt_cod_op1",
                    )
                )
            with col_op_nom:
                nom_op1, patio_op1, ok1, vacio1 = buscar_operador_rapido(cod_op1)
                st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
                if ok1:
                    st.success(f"👤 **{nom_op1}** · Patio: **{patio_op1}**")
                elif vacio1:
                    st.info(f"👤 {nom_op1}")
                else:
                    st.warning(f"👤 **{nom_op1}**")
            with col_op_lat:
                renderizar_tarjeta_lateral_anual(
                    cod_op1,
                    st.session_state.historial_cambios,
                    anio_consulta,
                )

            c_bar1, c_bar2, c_bar3 = st.columns([2.4, 1.2, 1.2])
            with c_bar1:
                st.radio(
                    "Jornada base para cálculo desde Hora de Inicio:",
                    ["Full Time (8:40)", "Part Time (6:30)"],
                    horizontal=True,
                    key="brt_jornada_base",
                    on_change=on_cambio_jornada_base,
                    disabled=(modalidad == "Cambio de día libre"),
                )
            with c_bar2:
                st.markdown("<div style='height: 24px;'></div>", unsafe_allow_html=True)
                if st.button(
                    "➕ Agregar fecha (máx. 7)",
                    use_container_width=True,
                    disabled=(st.session_state.brt_num_filas >= 7),
                ):
                    st.session_state.brt_num_filas += 1
                    st.rerun()
            with c_bar3:
                st.markdown("<div style='height: 30px;'></div>", unsafe_allow_html=True)
                st.caption(f"Filas activas: **{st.session_state.brt_num_filas} / 7**")

            h_cols = st.columns([0.32, 1.35, 1.25, 1.05, 0.85, 0.95, 1.1, 1.15, 1.15, 0.45])
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

            fechas_seleccionadas_vista = []
            for idx in range(st.session_state.brt_num_filas):
                r_cols = st.columns([0.32, 1.35, 1.25, 1.05, 0.85, 0.95, 1.1, 1.15, 1.15, 0.45])
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
                    fechas_seleccionadas_vista.append(f_obj)
                    prog1 = buscar_turno_rapido(cod_op1, f_obj)
                    r_cols[2].info(f"`{prog1['horario']}` · {prog1['sercon']}")

                    tipo_j = r_cols[3].selectbox(
                        f"J{idx+1}",
                        OPCIONES_JORNADA_FILA,
                        key=f"brt_ind_jornada_{idx}",
                        on_change=on_cambio_jornada_fila,
                        args=(idx,),
                        label_visibility="collapsed",
                    )

                    es_libre_fila = tipo_j == "Libre (-)"

                    # Sincronizar valores bloqueados ANTES de renderizar los selectbox de Tarea / P. Inicio / P. Final
                    if es_libre_fila:
                        st.session_state[f"brt_ind_tar_{idx}"] = "DESCANSO"
                        st.session_state[f"brt_ind_pi_{idx}"] = "-"
                        st.session_state[f"brt_ind_pf_{idx}"] = "-"
                    else:
                        if st.session_state.get(f"brt_ind_pi_{idx}") == "-":
                            st.session_state[f"brt_ind_pi_{idx}"] = PATIOS_CATALOGO[0]
                        if st.session_state.get(f"brt_ind_pf_{idx}") == "-":
                            st.session_state[f"brt_ind_pf_{idx}"] = PATIOS_CATALOGO[0]

                    if tipo_j in ("Full Time (8:40)", "Part Time (6:30)"):
                        h_ini_val = r_cols[4].time_input(
                            f"HI{idx+1}",
                            key=f"brt_ind_hini_{idx}",
                            step=300,
                            label_visibility="collapsed",
                        )
                        nuevo_hor = calcular_horario_desde_inicio(h_ini_val, tipo_j)
                        r_cols[5].success(f"**{nuevo_hor}**")
                    elif tipo_j == "Catálogo":
                        r_cols[4].caption("Catálogo")
                        nuevo_hor = r_cols[5].selectbox(
                            f"H{idx+1}",
                            JORNADAS_CATALOGO,
                            index=2,
                            key=f"brt_ind_hor_{idx}",
                            label_visibility="collapsed",
                        )
                    else:
                        r_cols[4].caption("—")
                        nuevo_hor = "-"
                        r_cols[5].caption("Libre (-)")

                    nueva_tar = r_cols[6].selectbox(
                        f"T{idx+1}",
                        TAREAS_CATALOGO,
                        key=f"brt_ind_tar_{idx}",
                        disabled=es_libre_fila,
                        label_visibility="collapsed",
                    )
                    nuevo_pi = r_cols[7].selectbox(
                        f"PI{idx+1}",
                        ["-"] + PATIOS_CATALOGO,
                        key=f"brt_ind_pi_{idx}",
                        disabled=es_libre_fila,
                        label_visibility="collapsed",
                    )
                    nuevo_pf = r_cols[8].selectbox(
                        f"PF{idx+1}",
                        ["-"] + PATIOS_CATALOGO,
                        key=f"brt_ind_pf_{idx}",
                        disabled=es_libre_fila,
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
                            "sercon": "DESCANSO" if es_libre_fila else nueva_tar,
                            "p_inicio": "-" if es_libre_fila else nuevo_pi,
                            "p_fin": "-" if es_libre_fila else nuevo_pf,
                        }
                    )
                else:
                    r_cols[2].caption("—")
                    r_cols[3].caption("—")
                    r_cols[4].caption("—")
                    r_cols[5].caption("—")
                    r_cols[6].caption("—")
                    r_cols[7].caption("—")
                    r_cols[8].caption("—")

                r_cols[9].button(
                    "🗑️",
                    key=f"del_ind_{idx}",
                    on_click=eliminar_fila_borrador,
                    args=(idx, "ind"),
                    help="Quitar esta fila",
                )

        else:
            st.markdown(
                '<div class="sub-header-brt">2. INTERCAMBIO ENTRE OPERADORES (FILAS 1 A 7)</div>',
                unsafe_allow_html=True,
            )
            c1_cod, c1_nom, c1_lat = st.columns([0.9, 1.8, 1.8])
            with c1_cod:
                cod_op1 = limpiar_codigo(
                    st.text_input(
                        "Operador 1 (Código)",
                        value="",
                        placeholder="Ej. 82601",
                        key="brt_int_cod_op1",
                    )
                )
            with c1_nom:
                nom_op1, patio_op1, ok1, vacio1 = buscar_operador_rapido(cod_op1)
                st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
                if ok1:
                    st.success(f"👤 **Op. 1:** {nom_op1} · Patio: **{patio_op1}**")
                elif vacio1:
                    st.info(f"👤 **Op. 1:** {nom_op1}")
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
                cod_op2 = limpiar_codigo(
                    st.text_input(
                        "Operador 2 (Código)",
                        value="",
                        placeholder="Ej. 82602",
                        key="brt_int_cod_op2",
                    )
                )
            with c2_nom:
                nom_op2, _, ok2, vacio2 = buscar_operador_rapido(cod_op2)
                st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
                if ok2:
                    st.success(f"👤 **Op. 2:** {nom_op2}")
                elif vacio2:
                    st.info(f"👤 **Op. 2:** {nom_op2}")
                else:
                    st.warning(f"👤 **Op. 2:** {nom_op2}")
            with c2_lat:
                renderizar_tarjeta_lateral_anual(
                    cod_op2,
                    st.session_state.historial_cambios,
                    anio_consulta,
                )

            ci_b1, ci_b2, _ = st.columns([1.4, 1.2, 3.0])
            with ci_b1:
                if st.button(
                    "➕ Agregar fecha intercambio (máx. 7)",
                    use_container_width=True,
                    disabled=(st.session_state.brt_num_filas >= 7),
                ):
                    st.session_state.brt_num_filas += 1
                    st.rerun()
            with ci_b2:
                st.caption(f"Filas activas: **{st.session_state.brt_num_filas} / 7**")

            hi_cols = st.columns([0.35, 1.5, 2.2, 2.2, 0.45])
            hi_cols[0].markdown("**#**")
            hi_cols[1].markdown("**Fecha del Intercambio**")
            hi_cols[2].markdown(f"**Programado Operador 1 ({cod_op1 or '—'})**")
            hi_cols[3].markdown(f"**Programado Operador 2 ({cod_op2 or '—'})**")
            hi_cols[4].markdown("**Quitar**")

            for idx in range(st.session_state.brt_num_filas):
                ri_cols = st.columns([0.35, 1.5, 2.2, 2.2, 0.45])
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

                ri_cols[4].button(
                    "🗑️",
                    key=f"del_int_{idx}",
                    on_click=eliminar_fila_borrador,
                    args=(idx, "int"),
                    help="Quitar esta fila",
                )

        st.markdown("---")
        if st.button(
            "✅ Guardar en Google Sheets y Generar Boleta F-ZEN-GO-07",
            type="primary",
            use_container_width=True,
        ):
            fechas_unicas_filas = {r["fecha"] for r in filas_cambio_op1}
            if responsable == PLACEHOLDER_RESP:
                st.error("⚠️ Debes seleccionar el **Responsable del Cambio** antes de guardar.")
            elif motivo_sel == PLACEHOLDER_MOTIVO:
                st.error("⚠️ Debes seleccionar la **Razón / Autorización** antes de guardar.")
            elif not cod_op1 or (es_intercambio and not cod_op2):
                st.error("⚠️ Debes ingresar el código del operador antes de guardar.")
            elif es_intercambio and cod_op1 == cod_op2:
                st.error("⚠️ En un intercambio, el Operador 1 y el Operador 2 deben ser distintos.")
            elif not filas_cambio_op1:
                st.error("⚠️ Selecciona al menos 1 fecha válida en la tabla para registrar el cambio.")
            elif len(fechas_unicas_filas) != len(filas_cambio_op1):
                st.warning("⚠️ Hay fechas repetidas en las filas del cambio. Quita o ajusta la fila duplicada.")
            else:
                codigo_registro = (
                    f"{cod_op1}_{cod_op2}" if es_intercambio else str(cod_op1)
                )
                fecha_cambio_principal = filas_cambio_op1[0]["fecha"]
                hora_cambio_resumen = (
                    f"{filas_prog_op1[0]['horario']}/{filas_cambio_op1[0]['horario']}"
                    if len(filas_cambio_op1) == 1
                    else f"{filas_prog_op1[0]['horario']}/{filas_cambio_op1[0]['horario']} (+{len(filas_cambio_op1)-1} días)"
                )

                df_h_actual = st.session_state.historial_cambios
                ya_existe = False
                if not df_h_actual.empty:
                    fecha_cmp_str = str(
                        datetime.strptime(fecha_cambio_principal, "%d/%m/%Y").date()
                    )
                    fc_norm = pd.to_datetime(
                        df_h_actual["FECHA DE CAMBIO"], errors="coerce"
                    ).dt.strftime("%Y-%m-%d")
                    dups = df_h_actual[
                        (df_h_actual["COGIDO"].astype(str).str.strip() == codigo_registro)
                        & (fc_norm == fecha_cmp_str)
                        & (
                            df_h_actual["HORA DE CAMBIO"].astype(str).str.strip()
                            == hora_cambio_resumen
                        )
                    ]
                    ya_existe = not dups.empty

                if ya_existe:
                    st.warning(
                        "⚠️ Este cambio ya se encuentra registrado en el historial (duplicado bloqueado)."
                    )
                else:
                    ref_num = (fecha_solicitud - date(1899, 12, 30)).days
                    nuevo_registro = {
                        "REF": ref_num,
                        "COGIDO": codigo_registro,
                        "FECHA DE SOLICITUD": pd.Timestamp(fecha_solicitud),
                        "FECHA DE CAMBIO": pd.Timestamp(
                            datetime.strptime(
                                fecha_cambio_principal, "%d/%m/%Y"
                            ).date()
                        ),
                        "HORA DE CAMBIO": hora_cambio_resumen,
                        "RESPONSABLE DEL CAMBIO": responsable,
                        "OBSERVACIONES": motivo_final,
                        "Stamp": datetime.now().replace(microsecond=0),
                    }
                    df_actualizado = pd.concat(
                        [pd.DataFrame([nuevo_registro]), df_h_actual], ignore_index=True
                    )
                    st.session_state.historial_cambios = df_actualizado

                    ok_gs, msg_gs = agregar_fila_en_gsheets(nuevo_registro)
                    if ok_gs:
                        st.success(
                            f"✅ Registro guardado en Google Sheets ({len(df_actualizado)} registros totales). Ve a la pestaña **2. Boleta Oficial F-ZEN-GO-07 (PDF)** para descargar."
                        )
                    else:
                        st.warning(
                            f"✅ Registro añadido en pantalla ({len(df_actualizado)} registros), pero no se sincronizó con Google Sheets: {msg_gs}"
                        )

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

# ------------------------------------------
# TAB 2: BOLETA OFICIAL F-ZEN-GO-07 Y PDF
# ------------------------------------------
with tab_boleta:
    st.subheader("🖨️ Emisión de Formulario Oficial F-ZEN-GO-07")

    if "boleta_activa" not in st.session_state:
        st.info(
            "ℹ️ Registra primero un cambio en la pestaña **1. Módulo Cambio de Horario (BRT)** para generar aquí su boleta oficial."
        )
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
                pd.DataFrame(bol["prog_op1"])[
                    ["fecha_corta", "codigo", "horario", "sercon", "p_inicio"]
                ],
                use_container_width=True,
                hide_index=True,
            )
            st.markdown("##### Horario de Cambio Operador 1")
            st.dataframe(
                pd.DataFrame(bol["cambio_op1"])[
                    ["fecha_corta", "codigo", "horario", "sercon", "p_inicio", "p_fin"]
                ],
                use_container_width=True,
                hide_index=True,
            )
        with col_v2:
            if bol["es_intercambio"]:
                st.markdown(f"##### Horario Programado Operador 2 ({bol['nombre_2']})")
                st.dataframe(
                    pd.DataFrame(bol["prog_op2"])[
                        ["fecha_corta", "codigo", "horario", "sercon", "p_inicio"]
                    ],
                    use_container_width=True,
                    hide_index=True,
                )
                st.markdown(f"##### Horario de Cambio Operador 2 ({bol['nombre_2']})")
                st.dataframe(
                    pd.DataFrame(bol["cambio_op2"])[
                        ["fecha_corta", "codigo", "horario", "sercon", "p_inicio", "p_fin"]
                    ],
                    use_container_width=True,
                    hide_index=True,
                )

# ------------------------------------------
# TAB 3: HISTORIAL ACUMULADO, ELIMINACIÓN Y EXCEL
# ------------------------------------------
with tab_historial:
    st.subheader("📊 Historial Acumulado de Cambios")
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
    if filtro_op.strip() and not df_filtrado.empty:
        df_filtrado = df_filtrado[
            df_filtrado["COGIDO"].astype(str).str.contains(filtro_op.strip(), case=False)
        ]
    if filtro_resp and not df_filtrado.empty:
        df_filtrado = df_filtrado[
            df_filtrado["RESPONSABLE DEL CAMBIO"].astype(str).isin(filtro_resp)
        ]

    st.caption(f"Mostrando **{len(df_filtrado)}** de **{len(df_h_show)}** registros acumulados.")
    st.dataframe(df_filtrado, use_container_width=True, hide_index=True)

    if not df_filtrado.empty:
        with st.expander("🗑️ Eliminar un registro del historial (y Google Sheets)"):
            opciones_del = {}
            for i_row, r in df_filtrado.iterrows():
                fc_val = pd.to_datetime(r.get("FECHA DE CAMBIO"), errors="coerce")
                fc_str = fc_val.strftime("%Y-%m-%d") if pd.notna(fc_val) else str(r.get("FECHA DE CAMBIO", ""))[:10]
                lbl = f"Fila {i_row} | Op. {r.get('COGIDO', '')} | {fc_str} | {r.get('HORA DE CAMBIO', '')} | {r.get('RESPONSABLE DEL CAMBIO', '')}"
                opciones_del[lbl] = (i_row, str(r.get("COGIDO", "")), fc_str, str(r.get("HORA DE CAMBIO", "")))
            sel_del = st.selectbox("Selecciona el registro a eliminar:", list(opciones_del.keys()))
            if st.button("🗑️ Confirmar eliminación de registro seleccionado", type="secondary"):
                idx_drop, cod_d, fc_d, hc_d = opciones_del[sel_del]
                st.session_state.historial_cambios = st.session_state.historial_cambios.drop(index=idx_drop).reset_index(drop=True)
                ok_d, msg_d = eliminar_fila_en_gsheets(cod_d, fc_d, hc_d)
                if ok_d:
                    st.success(f"✅ {msg_d}")
                else:
                    st.warning(f"Eliminado de la vista local. Nota Google Sheets: {msg_d}")
                st.rerun()

    buffer_xlsx = io.BytesIO()
    with pd.ExcelWriter(buffer_xlsx, engine="openpyxl") as writer:
        df_h_show.to_excel(
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
# TAB 4: ESTADÍSTICAS ANUALES (3 GRÁFICAS)
# ------------------------------------------
with tab_estadisticas:
    st.subheader("📈 Historial Acumulado y Estadísticas Anuales")
    df_stats_raw = st.session_state.historial_cambios.copy()

    if df_stats_raw.empty:
        st.info("No hay registros en el historial para generar estadísticas.")
    else:
        fc_flag = detectar_dayfirst(df_stats_raw["FECHA DE CAMBIO"]) if "FECHA DE CAMBIO" in df_stats_raw.columns else True
        fs_flag = detectar_dayfirst(df_stats_raw["FECHA DE SOLICITUD"]) if "FECHA DE SOLICITUD" in df_stats_raw.columns else True
        s_fc = df_stats_raw["FECHA DE CAMBIO"].apply(lambda x: parsear_fecha_segura(x, dayfirst=fc_flag))
        s_fs = df_stats_raw["FECHA DE SOLICITUD"].apply(lambda x: parsear_fecha_segura(x, dayfirst=fs_flag))
        df_stats_raw["FECHA_ANALISIS"] = pd.to_datetime(s_fc.fillna(s_fs), errors="coerce")
        df_stats_val = df_stats_raw[df_stats_raw["FECHA_ANALISIS"].notna()].copy()

        anios_disp = sorted(df_stats_val["FECHA_ANALISIS"].dt.year.unique().tolist(), reverse=True)
        if not anios_disp:
            anios_disp = [date.today().year]

        c_sel_anio, c_sel_modo, _ = st.columns([1, 1.5, 2.5])
        with c_sel_anio:
            anio_sel = st.selectbox("Año de análisis:", anios_disp, index=0)
        with c_sel_modo:
            modo_anual = st.radio(
                "Vista gráfica anual:",
                ["Acumulado", "Por Mes"],
                horizontal=True,
            )

        df_anio = df_stats_val[df_stats_val["FECHA_ANALISIS"].dt.year == anio_sel].copy()
        df_anio["MES"] = df_anio["FECHA_ANALISIS"].dt.month
        df_anio["DIA"] = df_anio["FECHA_ANALISIS"].dt.day
        df_anio["TIPO_CAMBIO"] = df_anio.apply(clasificar_tipo_cambio, axis=1)

        hoy = date.today()
        cambios_mes_actual = int(
            ((df_stats_val["FECHA_ANALISIS"].dt.year == hoy.year) & (df_stats_val["FECHA_ANALISIS"].dt.month == hoy.month)).sum()
        )
        ops_unicos = set()
        for c_str in df_anio["COGIDO"].astype(str):
            for parte in c_str.split("_"):
                if parte.strip():
                    ops_unicos.add(parte.strip())

        k1, k2, k3, k4 = st.columns(4)
        k1.metric(f"Total Cambios ({anio_sel})", f"{len(df_anio):,}")
        k2.metric("Total Acumulado Histórico", f"{len(df_stats_raw):,}")
        k3.metric(f"Operadores con Cambios ({anio_sel})", f"{len(ops_unicos):,}")
        k4.metric(f"Cambios en {MESES_ES[hoy.month]}", f"{cambios_mes_actual:,}")

        st.markdown("---")

        # GRÁFICA 1 (GRANDE): ACUMULADO ANUAL POR MES
        st.markdown(f"### 📅 Acumulado Anual por Mes ({anio_sel})")
        conteo_mes = df_anio.groupby("MES").size().to_dict()
        ultimo_mes_datos = max(conteo_mes.keys()) if conteo_mes else 0
        limite_mes = max(hoy.month, ultimo_mes_datos) if anio_sel == hoy.year else 12

        rows_mes = []
        acum = 0
        for m in range(1, 13):
            cant_m = conteo_mes.get(m, 0)
            acum += cant_m
            es_futuro = (anio_sel == hoy.year and m > limite_mes and cant_m == 0)
            valor_mostrar = 0 if (modo_anual == "Acumulado" and es_futuro) else (acum if modo_anual == "Acumulado" else cant_m)
            rows_mes.append({
                "Orden": m,
                "Mes": f"{m:02d}. {MESES_ES[m]}",
                "Cambios del Mes": cant_m,
                "Acumulado": 0 if es_futuro else acum,
                "Valor": valor_mostrar,
            })
        df_meses_chart = pd.DataFrame(rows_mes)

        chart_anual = (
            alt.Chart(df_meses_chart)
            .mark_bar(color="#0070c0", cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
            .encode(
                x=alt.X("Mes:N", sort=list(df_meses_chart["Mes"]), title=None),
                y=alt.Y("Valor:Q", title="Acumulado de Cambios" if modo_anual == "Acumulado" else "Cambios del Mes"),
                tooltip=["Mes", "Cambios del Mes", "Acumulado"],
            )
            .properties(height=360)
        )
        st.altair_chart(chart_anual, use_container_width=True)

        # GRÁFICA 2 Y GRÁFICA 3
        col_g2, col_g3 = st.columns([1.3, 1])

        with col_g2:
            mes_sel = st.selectbox(
                "📆 Cambios por Día del Mes Actual / Seleccionado:",
                list(range(1, 13)),
                index=hoy.month - 1,
                format_func=lambda m: f"{MESES_LARGO_ES[m].capitalize()} {anio_sel}",
            )
            dias_en_mes = pd.Timestamp(year=anio_sel, month=mes_sel, day=1).days_in_month
            df_mes_dia = df_anio[df_anio["MES"] == mes_sel]
            conteo_dia = df_mes_dia.groupby("DIA").size().to_dict()
            df_dias_chart = pd.DataFrame([
                {"Día": f"{d:02d}", "Orden": d, "Cambios": conteo_dia.get(d, 0)}
                for d in range(1, dias_en_mes + 1)
            ])
            chart_dias = (
                alt.Chart(df_dias_chart)
                .mark_bar(color="#0ea5e9", cornerRadiusTopLeft=3, cornerRadiusTopRight=3)
                .encode(
                    x=alt.X("Día:N", sort=list(df_dias_chart["Día"]), title="Día del mes"),
                    y=alt.Y("Cambios:Q", title="Cambios"),
                    tooltip=["Día", "Cambios"],
                )
                .properties(height=300)
            )
            st.altair_chart(chart_dias, use_container_width=True)

        with col_g3:
            st.markdown(f"#### 🔀 Distribución por Tipo de Cambio ({anio_sel})")
            tipos_orden = ["Cambio de horario", "Cambio con otro operador", "Cambio de día libre"]
            conteo_tipo = df_anio.groupby("TIPO_CAMBIO").size().to_dict()
            total_anio_n = max(1, len(df_anio))
            df_tipo_chart = pd.DataFrame([
                {
                    "Tipo de Cambio": t,
                    "Cantidad": conteo_tipo.get(t, 0),
                    "Porcentaje": round(conteo_tipo.get(t, 0) * 100 / total_anio_n, 1),
                }
                for t in tipos_orden
            ])
            chart_tipo = (
                alt.Chart(df_tipo_chart)
                .mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4)
                .encode(
                    y=alt.Y("Tipo de Cambio:N", sort="-x", title=None),
                    x=alt.X("Cantidad:Q", title="Total"),
                    color=alt.Color(
                        "Tipo de Cambio:N",
                        scale=alt.Scale(
                            domain=tipos_orden,
                            range=["#0070c0", "#0f766e", "#d97706"],
                        ),
                        legend=None,
                    ),
                    tooltip=["Tipo de Cambio", "Cantidad", "Porcentaje"],
                )
                .properties(height=280)
            )
            st.altair_chart(chart_tipo, use_container_width=True)

        st.markdown(f"#### 🏅 Operadores con Más Cambios en {anio_sel}")
        top_ops_counts = {}
        for c_str in df_anio["COGIDO"].astype(str):
            for parte in c_str.split("_"):
                p = parte.strip()
                if p:
                    top_ops_counts[p] = top_ops_counts.get(p, 0) + 1
        if top_ops_counts:
            df_top_ops = (
                pd.DataFrame(
                    [{"Código Operador": k, "Cambios": v} for k, v in top_ops_counts.items()]
                )
                .sort_values("Cambios", ascending=False)
                .head(15)
            )
            st.dataframe(df_top_ops, use_container_width=True, hide_index=True)

# ------------------------------------------
# TAB 5: CUADRÍCULA INTERACTIVA TIPO EXCEL (CTRL + V)
# ------------------------------------------
with tab_pegar_tabla:
    st.subheader("📋 Cuadrícula Tipo Excel — Pega o Edita Horarios Directamente")
    st.markdown(
        "Haz clic en la primera celda (`Fecha`), presiona **`Ctrl + V`** para pegar las filas copiadas desde Excel y luego pulsa **Fijar programación en sesión**."
    )

    plantilla_excel = pd.DataFrame(
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
        plantilla_excel,
        num_rows="dynamic",
        use_container_width=True,
        key="editor_excel_grid",
    )

    if st.button(
        "💾 Fijar programación de esta cuadrícula en sesión",
        type="primary",
    ):
        df_norm_grid, a_ord_grid = normalizar_tabla_programacion(df_editor)
        if not df_norm_grid.empty:
            indexar_programacion_en_sesion(
                df_norm_grid,
                f"Cuadrícula Excel ({len(df_norm_grid)} turnos - {datetime.now().strftime('%H:%M')})",
                a_ord_grid,
            )
            st.success(
                f"✅ Se cargaron **{len(df_norm_grid)}** turnos válidos desde la cuadrícula. Ya puedes ir a la pestaña 1."
            )
            st.rerun()
        else:
            st.warning(
                "⚠️ Asegúrate de que al menos las columnas **Fecha** y **CODIGO** tengan datos válidos antes de guardar."
            )
