from datetime import datetime, date, time, timedelta
import io
import pandas as pd
import streamlit as st
from fpdf import FPDF

# Estilo CSS para impresión directa limpia desde navegador (Ctrl + P)
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
    .boleta-preview {
        border: 2px solid #0f2952;
        padding: 14px;
        border-radius: 6px;
        background-color: #ffffff;
        color: #0f172a;
        font-family: Arial, sans-serif;
    }
    .alerta-historial {
        background-color: #fef3c7;
        border-left: 5px solid #d97706;
        color: #78350f;
        padding: 8px 12px;
        border-radius: 4px;
        font-size: 13.5px;
        margin-top: 4px;
        margin-bottom: 8px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("🚌 Módulo de Cambio de Horarios & Emisión F-ZEN-GO-07")

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
    """Normaliza tanto ROW-DATA como el reporte ASIGNACIÓN o datos pegados desde Excel."""
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
        return pd.DataFrame()

    df["CODIGO_STR"] = df["CODIGO"].apply(limpiar_codigo)
    df = df[df["CODIGO_STR"] != ""].copy()
    df["FECHA_DATE"] = pd.to_datetime(df["Fecha"], errors="coerce").dt.date
    df = df[df["FECHA_DATE"].notna()].copy()

    if "HORARIO_RAW" in df.columns:
        df["HORARIO_STR"] = df["HORARIO_RAW"].fillna("-").astype(str).str.strip()
    else:
        hi = df["H.Inicial"].apply(formatear_hora_celda) if "H.Inicial" in df.columns else ""
        hf = df["H. Final"].apply(formatear_hora_celda) if "H. Final" in df.columns else ""
        df["HORARIO_STR"] = [
            f"{a} - {b}" if (a and b and a != "-" and b != "-") else "-"
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

    return df


def indexar_programacion_en_sesion(df_norm: pd.DataFrame, nombre_fuente: str):
    """Guarda en st.session_state los diccionarios indexados para consulta instantánea."""
    operadores_dict = {}
    turnos_dict = {}
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

    st.session_state.operadores_dict = operadores_dict
    st.session_state.turnos_dict = turnos_dict
    st.session_state.prog_total_filas = len(df_norm)
    st.session_state.prog_nombre_archivo = nombre_fuente
    st.session_state.prog_cargada = len(df_norm) > 0
    st.session_state.mostrar_cargador = False


def parsear_texto_pegado_excel(texto_tsv: str) -> pd.DataFrame:
    """Convierte texto copiado directo desde Excel (separado por tabulaciones) en DataFrame normalizado."""
    lineas = [l for l in texto_tsv.splitlines() if l.strip()]
    if not lineas:
        return pd.DataFrame()

    # Leer como TSV
    df_tsv = pd.read_csv(io.StringIO("\n".join(lineas)), sep="\t", header=None)
    # Revisar si la primera fila contiene encabezados reales
    primera_fila_upper = [str(x).strip().upper() for x in df_tsv.iloc[0].tolist()]
    if any(k in primera_fila_upper for k in ["CODIGO", "CÓDIGO", "FECHA"]):
        df_tsv.columns = [str(x).strip() for x in df_tsv.iloc[0].tolist()]
        df_tsv = df_tsv.iloc[1:].reset_index(drop=True)
        return normalizar_tabla_programacion(df_tsv)

    # Si pegaron sin encabezados, inferir según cantidad de columnas (formato ASIGNACIÓN o ROW-DATA)
    ncols = df_tsv.shape[1]
    if ncols >= 11:
        # Formato estándar A..K: Ref, Fecha, Tarea, CODIGO, NOMBRE, W.T., H.Inicial, origen, H. Final, Patio Final, PADMIN
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
        # Sin columna Ref (B..J/K): Fecha, Tarea, CODIGO, NOMBRE, W.T., H.Inicial, origen, H. Final, Patio Final, [PADMIN]
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
        df_norm = normalizar_tabla_programacion(df_try)
        if not df_norm.empty:
            df_prog = df_norm
            break

    df_hist = pd.DataFrame(
        columns=[
            "REF",
            "COGIDO",
            "FECHA DE SOLICITUD",
            "FECHA DE CAMBIO",
            "HORA DE CAMBIO",
            "RESPONSABLE DEL CAMBIO",
            "OBSERVACIONES",
            "Stamp",
        ]
    )
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
        df_hist = df_h

    return df_prog, df_hist, lista_resp


def obtener_resumen_mensual_operador(df_hist: pd.DataFrame, codigo: str):
    cod = limpiar_codigo(codigo)
    if not cod or df_hist.empty or "COGIDO" not in df_hist.columns:
        return 0, []

    def contiene_codigo(val):
        partes = [p.strip() for p in str(val).split("_")]
        return cod in partes

    sub = df_hist[df_hist["COGIDO"].apply(contiene_codigo)].copy()
    if sub.empty:
        return 0, []

    f_cambio = pd.to_datetime(sub.get("FECHA DE CAMBIO"), errors="coerce")
    f_sol = pd.to_datetime(sub.get("FECHA DE SOLICITUD"), errors="coerce")
    fecha_ref = f_cambio.fillna(f_sol)

    sub["ANIO_MES"] = fecha_ref.dt.to_period("M")
    sub_valid = sub[sub["ANIO_MES"].notna()]
    total = len(sub)

    if sub_valid.empty:
        return total, []

    conteo = (
        sub_valid.groupby("ANIO_MES")
        .size()
        .reset_index(name="cant")
        .sort_values("ANIO_MES", ascending=False)
    )
    detalle = []
    for _, r in conteo.iterrows():
        p = r["ANIO_MES"]
        mes_nom = MESES_ES.get(p.month, str(p.month))
        detalle.append(f"{r['cant']} en {mes_nom} {p.year}")

    return total, detalle


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
# GENERADOR PDF IDÉNTICO AL FORMATO F-ZEN-GO-07
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
    st.session_state.operadores_dict = {}
    st.session_state.turnos_dict = {}
    st.session_state.prog_total_filas = 0
    st.session_state.lista_responsables = list(RESPONSABLES_DEFAULT)
    st.session_state.historial_cambios = pd.DataFrame(
        columns=[
            "REF",
            "COGIDO",
            "FECHA DE SOLICITUD",
            "FECHA DE CAMBIO",
            "HORA DE CAMBIO",
            "RESPONSABLE DEL CAMBIO",
            "OBSERVACIONES",
            "Stamp",
        ]
    )
    st.session_state.mostrar_cargador = True


# ==========================================
# SIDEBAR: CARGA PERSISTENTE (ARCHIVO O PEGAR EXCEL)
# ==========================================
st.sidebar.header("📂 1. Programación Activa")

if st.session_state.prog_cargada and not st.session_state.mostrar_cargador:
    st.sidebar.success(
        f"📌 **Fuente activa:**\n{st.session_state.prog_nombre_archivo}\n\n"
        f"• **{st.session_state.prog_total_filas:,}** turnos indexados\n"
        f"• **{len(st.session_state.historial_cambios)}** cambios en historial"
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
                archivo_subido.getvalue()
            )
            st.session_state.lista_responsables = resp_l
            if not df_h_loaded.empty:
                st.session_state.historial_cambios = df_h_loaded.copy()
            indexar_programacion_en_sesion(df_prog_l, archivo_subido.name)
            st.rerun()
    else:
        st.sidebar.caption(
            "Copia las filas en Excel (`Ctrl + C`) desde **Fecha** hasta **PADMIN** (o toda la tabla de ASIGNACIÓN) y pégalas aquí:"
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
                df_pegado = parsear_texto_pegado_excel(texto_pegado)
                if not df_pegado.empty:
                    indexar_programacion_en_sesion(
                        df_pegado,
                        f"Pegado desde Excel ({datetime.now().strftime('%H:%M')})",
                    )
                    st.rerun()
                else:
                    st.sidebar.error(
                        "⚠️ No se reconocieron las columnas. Asegúrate de copiar desde la columna Fecha hasta Patio Final/PADMIN."
                    )
            else:
                st.sidebar.warning("⚠️ Pega primero los datos copiados de Excel.")

    if st.session_state.prog_cargada:
        if st.sidebar.button("Cancelar cambio", use_container_width=True):
            st.session_state.mostrar_cargador = False
            st.rerun()
    else:
        st.sidebar.info(
            "ℹ️ La programación quedará fija en pantalla mientras trabajas hasta que decidas cambiarla."
        )


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
# PESTAÑAS PRINCIPALES
# ==========================================
tab_registro, tab_boleta, tab_historial, tab_pegar_tabla = st.tabs(
    [
        "📝 1. Registrar Cambio de Horario",
        "🖨️ 2. Boleta F-ZEN-GO-07 (PDF / Imprimir)",
        "📊 3. Historial Acumulado y Descarga Excel",
        "📋 4. Tabla Tipo Excel (Pegar / Editar)",
    ]
)

# ------------------------------------------
# TAB 1: FORMULARIO + CONTADOR HISTÓRICO MENSUAL
# ------------------------------------------
with tab_registro:
    st.subheader("Captura de Solicitud de Cambio")

    col_m1, col_m2, col_m3, col_m4 = st.columns([1.3, 1, 1.2, 1.5])
    with col_m1:
        modalidad = st.selectbox(
            "Tipo de Movimiento",
            [
                "Cambio de horario",
                "Cambio con otro operador",
                "Cambio de día libre",
            ],
        )
    with col_m2:
        fecha_solicitud = st.date_input("Fecha de Solicitud", value=date.today())
    with col_m3:
        responsable = st.selectbox(
            "Responsable del Cambio", st.session_state.lista_responsables
        )
    with col_m4:
        motivo_sel = st.selectbox("Motivo / Autorización", MOTIVOS_ESTANDAR)

    motivo_extra = st.text_input(
        "Detalle adicional en Observaciones (Opcional)",
        placeholder="Ej. Cita cardiología / Cubre servicio específico...",
    )
    motivo_final = (
        f"{motivo_sel} - {motivo_extra.strip().upper()}"
        if motivo_extra.strip()
        else motivo_sel
    )

    es_intercambio = modalidad == "Cambio con otro operador"

    st.markdown("---")
    c_op1, c_op2 = st.columns(2)
    with c_op1:
        cod_op1 = limpiar_codigo(st.text_input("Código Operador 1", value="82601"))
        nom_op1, patio_op1, ok1 = buscar_operador_rapido(cod_op1)
        if ok1:
            st.success(f"👤 **Operador 1:** {nom_op1} | **Patio:** {patio_op1}")
        else:
            st.warning(f"👤 **Operador 1:** {nom_op1}")

        tot1, det1 = obtener_resumen_mensual_operador(
            st.session_state.historial_cambios, cod_op1
        )
        if tot1 > 0:
            resumen_txt1 = ", ".join(det1) if det1 else "sin fecha detallada"
            st.markdown(
                f'<div class="alerta-historial">📊 <b>Historial Operador {cod_op1}:</b> '
                f'<b>{tot1}</b> cambio(s) registrado(s) → <b>{resumen_txt1}</b></div>',
                unsafe_allow_html=True,
            )
        elif cod_op1:
            st.caption(f"📊 Historial Operador {cod_op1}: 0 cambios previos registrados.")

    cod_op2, nom_op2 = "", ""
    if es_intercambio:
        with c_op2:
            cod_op2 = limpiar_codigo(st.text_input("Código Operador 2", value="82602"))
            nom_op2, _, ok2 = buscar_operador_rapido(cod_op2)
            if ok2:
                st.success(f"👤 **Operador 2:** {nom_op2}")
            else:
                st.warning(f"👤 **Operador 2:** {nom_op2}")

            tot2, det2 = obtener_resumen_mensual_operador(
                st.session_state.historial_cambios, cod_op2
            )
            if tot2 > 0:
                resumen_txt2 = ", ".join(det2) if det2 else "sin fecha detallada"
                st.markdown(
                    f'<div class="alerta-historial">📊 <b>Historial Operador {cod_op2}:</b> '
                    f'<b>{tot2}</b> cambio(s) registrado(s) → <b>{resumen_txt2}</b></div>',
                    unsafe_allow_html=True,
                )
            elif cod_op2:
                st.caption(f"📊 Historial Operador {cod_op2}: 0 cambios previos registrados.")

    st.markdown("---")
    cant_dias = st.number_input(
        "Cantidad de días a modificar (1 a 7)", min_value=1, max_value=7, value=1
    )

    filas_prog_op1, filas_cambio_op1 = [], []
    filas_prog_op2, filas_cambio_op2 = [], []

    for i in range(int(cant_dias)):
        st.markdown(f"**Día #{i + 1}**")
        fecha_d = st.date_input(
            f"Fecha del cambio #{i + 1}",
            value=fecha_solicitud + timedelta(days=i + 1),
            key=f"f_cambio_{i}",
        )

        prog1 = buscar_turno_rapido(cod_op1, fecha_d)
        filas_prog_op1.append(
            {
                "fecha": fecha_d.strftime("%d/%m/%Y"),
                "fecha_corta": fecha_corta_es(fecha_d),
                "codigo": cod_op1,
                **prog1,
            }
        )

        if es_intercambio:
            prog2 = buscar_turno_rapido(cod_op2, fecha_d)
            filas_prog_op2.append(
                {
                    "fecha": fecha_d.strftime("%d/%m/%Y"),
                    "fecha_corta": fecha_corta_es(fecha_d),
                    "codigo": cod_op2,
                    **prog2,
                }
            )

            col_i1, col_i2 = st.columns(2)
            with col_i1:
                st.caption(
                    f"Programado Op 1 ({cod_op1}): `{prog1['horario']}` | Sercon: `{prog1['sercon']}` | {prog1['p_inicio']} → {prog1['p_fin']}"
                )
            with col_i2:
                st.caption(
                    f"Programado Op 2 ({cod_op2}): `{prog2['horario']}` | Sercon: `{prog2['sercon']}` | {prog2['p_inicio']} → {prog2['p_fin']}"
                )

            filas_cambio_op1.append(
                {
                    "fecha": fecha_d.strftime("%d/%m/%Y"),
                    "fecha_corta": fecha_corta_es(fecha_d),
                    "codigo": cod_op1,
                    "horario": prog2["horario"],
                    "sercon": prog2["sercon"],
                    "p_inicio": prog2["p_inicio"],
                    "p_fin": prog2["p_fin"],
                }
            )
            filas_cambio_op2.append(
                {
                    "fecha": fecha_d.strftime("%d/%m/%Y"),
                    "fecha_corta": fecha_corta_es(fecha_d),
                    "codigo": cod_op2,
                    "horario": prog1["horario"],
                    "sercon": prog1["sercon"],
                    "p_inicio": prog1["p_inicio"],
                    "p_fin": prog1["p_fin"],
                }
            )
        else:
            st.caption(
                f"Programado actual ({cod_op1}): `{prog1['horario']}` | Sercon: `{prog1['sercon']}` | {prog1['p_inicio']} → {prog1['p_fin']}"
            )
            c1, c2, c3, c4 = st.columns([1.2, 1.2, 1.3, 1.3])
            with c1:
                nuevo_hor = st.selectbox(
                    f"Nuevo Horario #{i + 1}",
                    JORNADAS_CATALOGO,
                    index=2,
                    key=f"hor_{i}",
                )
            with c2:
                nueva_tarea = st.selectbox(
                    f"Nueva Tarea / Sercon #{i + 1}",
                    TAREAS_CATALOGO,
                    key=f"tar_{i}",
                )
            with c3:
                nuevo_pi = st.selectbox(
                    f"Patio Inicio #{i + 1}",
                    PATIOS_CATALOGO,
                    key=f"pi_{i}",
                )
            with c4:
                nuevo_pf = st.selectbox(
                    f"Patio Final #{i + 1}",
                    PATIOS_CATALOGO,
                    key=f"pf_{i}",
                )

            filas_cambio_op1.append(
                {
                    "fecha": fecha_d.strftime("%d/%m/%Y"),
                    "fecha_corta": fecha_corta_es(fecha_d),
                    "codigo": cod_op1,
                    "horario": nuevo_hor,
                    "sercon": nueva_tarea,
                    "p_inicio": nuevo_pi,
                    "p_fin": nuevo_pf,
                }
            )

    st.markdown("---")
    if st.button(
        "✅ Guardar Cambio en Historial y Generar Boleta F-ZEN-GO-07",
        type="primary",
        use_container_width=True,
    ):
        if not cod_op1 or (es_intercambio and not cod_op2):
            st.error("⚠️ Debes ingresar el código del operador antes de guardar.")
        else:
            codigo_registro = (
                f"{cod_op1}_{cod_op2}" if es_intercambio else str(cod_op1)
            )
            fecha_cambio_principal = filas_cambio_op1[0]["fecha"]
            hora_cambio_resumen = (
                f"{filas_prog_op1[0]['horario']}/{filas_cambio_op1[0]['horario']}"
            )

            df_h_actual = st.session_state.historial_cambios
            ya_existe = False
            if not df_h_actual.empty:
                fecha_cmp_str = str(
                    datetime.strptime(fecha_cambio_principal, "%d/%m/%Y").date()
                )
                dups = df_h_actual[
                    (df_h_actual["COGIDO"].astype(str).str.strip() == codigo_registro)
                    & (
                        df_h_actual["FECHA DE CAMBIO"]
                        .astype(str)
                        .str.slice(0, 10)
                        == fecha_cmp_str
                    )
                    & (
                        df_h_actual["HORA DE CAMBIO"].astype(str).str.strip()
                        == hora_cambio_resumen
                    )
                ]
                ya_existe = not dups.empty

            if ya_existe:
                st.warning(
                    "⚠️ Este cambio ya se encuentra registrado en el historial (duplicado bloqueado). Puedes descargar su boleta en la pestaña 2."
                )
            else:
                ref_num = (fecha_solicitud - date(1899, 12, 30)).days
                nueva_fila = pd.DataFrame(
                    [
                        {
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
                    ]
                )
                st.session_state.historial_cambios = pd.concat(
                    [nueva_fila, df_h_actual], ignore_index=True
                )
                st.success(
                    f"✅ Registro guardado y sumado al historial ({len(st.session_state.historial_cambios)} registros totales). Ve a la pestaña **2. Boleta F-ZEN-GO-07** para descargar el PDF."
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
            "ℹ️ Registra primero un cambio en la pestaña **1. Registrar Cambio de Horario** para generar aquí su boleta oficial."
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
# TAB 3: HISTORIAL ACUMULADO Y EXPORTACIÓN EXCEL
# ------------------------------------------
with tab_historial:
    st.subheader("📊 Historial Acumulado de Cambios")
    df_h_show = st.session_state.historial_cambios

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
    if filtro_op.strip():
        df_filtrado = df_filtrado[
            df_filtrado["COGIDO"].astype(str).str.contains(filtro_op.strip(), case=False)
        ]
    if filtro_resp:
        df_filtrado = df_filtrado[
            df_filtrado["RESPONSABLE DEL CAMBIO"].astype(str).isin(filtro_resp)
        ]

    st.caption(f"Mostrando **{len(df_filtrado)}** de **{len(df_h_show)}** registros acumulados.")
    st.dataframe(df_filtrado, use_container_width=True, hide_index=True)

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
# TAB 4: CUADRÍCULA INTERACTIVA TIPO EXCEL (CTRL + V)
# ------------------------------------------
with tab_pegar_tabla:
    st.subheader("📋 Cuadrícula Tipo Excel — Pega o Edita Horarios Directamente")
    st.markdown(
        "Haz clic en la primera celda (`Fecha`), presiona **`Ctrl + V`** para pegar las filas copiadas desde Excel y luego pulsa **Guardar programación en sesión**."
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
        df_norm_grid = normalizar_tabla_programacion(df_editor)
        if not df_norm_grid.empty:
            indexar_programacion_en_sesion(
                df_norm_grid,
                f"Cuadrícula Excel ({len(df_norm_grid)} turnos - {datetime.now().strftime('%H:%M')})",
            )
            st.success(
                f"✅ Se cargaron **{len(df_norm_grid)}** turnos desde la cuadrícula. Ya puedes ir a la pestaña 1."
            )
            st.rerun()
        else:
            st.warning(
                "⚠️ Asegúrate de que al menos las columnas **Fecha** y **CODIGO** tengan datos válidos antes de guardar."
            )
