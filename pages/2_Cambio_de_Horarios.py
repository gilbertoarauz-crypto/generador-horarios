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
    .boleta-box {
        border: 2px solid #1e293b;
        padding: 16px;
        border-radius: 6px;
        background-color: #ffffff;
        color: #0f172a;
        font-family: Arial, sans-serif;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("🚌 Módulo de Cambio de Horarios & Emisión F-ZEN-GO-07")

# ==========================================
# CATÁLOGOS OFICIALES EXTRAÍDOS DEL LIBRO
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


# ==========================================
# FUNCIONES DE LECTURA Y BÚSQUEDA EN ROW-DATA
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
    return str(val).strip()


@st.cache_data(show_spinner="Cargando base ROW-DATA e Historial...")
def cargar_libro_excel(file_bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    sheets = xls.sheet_names

    lista_resp = list(RESPONSABLES_DEFAULT)
    if "Lista" in sheets:
        df_l = pd.read_excel(xls, sheet_name="Lista")
        if not df_l.empty:
            col0 = df_l.columns[0]
            vals = [
                str(x).strip()
                for x in df_l[col0].dropna().tolist()
                if str(x).strip() != ""
            ]
            if vals:
                lista_resp = vals

    df_row = pd.DataFrame()
    if "ROW-DATA" in sheets:
        df_row = pd.read_excel(xls, sheet_name="ROW-DATA")
        df_row.columns = [str(c).strip() for c in df_row.columns]
        if "Fecha" in df_row.columns:
            df_row["FECHA_DATE"] = pd.to_datetime(
                df_row["Fecha"], errors="coerce"
            ).dt.date
        if "CODIGO" in df_row.columns:
            df_row["CODIGO_STR"] = (
                df_row["CODIGO"]
                .astype(str)
                .str.replace(r"\.0$", "", regex=True)
                .str.strip()
            )
        col_horario = next(
            (c for c in df_row.columns if "Unnamed" in c or "HORARIO" in c.upper()),
            None,
        )
        if col_horario:
            df_row["HORARIO_STR"] = df_row[col_horario].fillna("-").astype(str).str.strip()
        else:
            df_row["HORARIO_STR"] = df_row.apply(
                lambda r: (
                    f"{formatear_hora_celda(r.get('H.Inicial'))} - {formatear_hora_celda(r.get('H. Final'))}"
                    if formatear_hora_celda(r.get("H.Inicial"))
                    else "-"
                ),
                axis=1,
            )

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
            df_h = df_h[df_h["COGIDO"].notna()]
        df_hist = df_h

    return df_row, df_hist, lista_resp


def buscar_operador_info(df_row: pd.DataFrame, codigo: str):
    if df_row.empty or not codigo:
        return "OPERADOR NO ENCONTRADO", "LOS PUEBLOS"
    cod_clean = str(codigo).strip()
    sub = df_row[df_row["CODIGO_STR"] == cod_clean]
    if sub.empty:
        return "OPERADOR NO REGISTRADO EN ROW-DATA", "LOS PUEBLOS"
    nombre = str(sub["NOMBRE"].iloc[0]).strip()
    patio = (
        str(sub["PADMIN"].iloc[0]).strip().upper()
        if "PADMIN" in sub.columns and pd.notna(sub["PADMIN"].iloc[0])
        else "LOS PUEBLOS"
    )
    return nombre, patio


def buscar_programacion_dia(df_row: pd.DataFrame, codigo: str, fecha_obj: date):
    if df_row.empty or not codigo or not fecha_obj:
        return {"horario": "-", "sercon": "-", "p_inicio": "-", "p_fin": "-"}
    cod_clean = str(codigo).strip()
    sub = df_row[
        (df_row["CODIGO_STR"] == cod_clean) & (df_row["FECHA_DATE"] == fecha_obj)
    ]
    if sub.empty:
        return {"horario": "-", "sercon": "-", "p_inicio": "-", "p_fin": "-"}
    fila = sub.iloc[0]
    return {
        "horario": str(fila.get("HORARIO_STR", "-")).strip(),
        "sercon": str(fila.get("Tarea", "-")).strip()
        if pd.notna(fila.get("Tarea"))
        else "-",
        "p_inicio": str(fila.get("origen", "-")).strip()
        if pd.notna(fila.get("origen"))
        else "-",
        "p_fin": str(fila.get("Patio Final", "-")).strip()
        if pd.notna(fila.get("Patio Final"))
        else "-",
    }


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
# GENERADOR DE BOLETA OFICIAL F-ZEN-GO-07 EN PDF
# ==========================================
def generar_pdf_boleta(datos_boleta: dict) -> bytes:
    pdf = FPDF(orientation="P", unit="mm", format="Letter")
    pdf.set_auto_page_break(auto=True, margin=10)
    pdf.add_page()

    def dibujar_encabezado(titulo_extra=""):
        pdf.set_fill_color(25, 45, 85)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Arial", "B", 11)
        pdf.cell(
            0,
            8,
            limpiar_texto_pdf(
                f"TRANSPORTE MASIVO DE PANAMA, S.A. (MiBus) - CAMBIO DE HORARIO {titulo_extra}"
            ),
            border=1,
            ln=1,
            align="C",
            fill=True,
        )
        pdf.set_text_color(0, 0, 0)
        pdf.set_font("Arial", "B", 8)
        pdf.cell(98, 5, "Formulario: F-ZEN-GO-07 / F-ZEN-GO-06", border=1, align="L")
        pdf.cell(
            98,
            5,
            limpiar_texto_pdf(f"MOTIVO / AUTORIZACION: {datos_boleta['motivo']}"),
            border=1,
            ln=1,
            align="L",
        )

    def dibujar_tabla_detalle(titulo_bloque, filas):
        pdf.set_fill_color(230, 236, 245)
        pdf.set_font("Arial", "B", 8)
        pdf.cell(0, 5.5, limpiar_texto_pdf(titulo_bloque), border=1, ln=1, fill=True)

        anchos = [10, 26, 20, 32, 36, 36, 36]
        headers = [
            "N",
            "Fecha",
            "Codigo",
            "Horario",
            "Codigo Sercon",
            "P. Inicio",
            "P. Final",
        ]
        pdf.set_font("Arial", "B", 7.5)
        for w, h in zip(anchos, headers):
            pdf.cell(w, 5, limpiar_texto_pdf(h), border=1, align="C")
        pdf.ln()

        pdf.set_font("Arial", "", 7.5)
        for idx, f in enumerate(filas, start=1):
            vals = [
                str(idx),
                f.get("fecha", ""),
                f.get("codigo", ""),
                f.get("horario", "-"),
                f.get("sercon", "-"),
                f.get("p_inicio", "-"),
                f.get("p_fin", "-"),
            ]
            for w, v in zip(anchos, vals):
                pdf.cell(w, 5, limpiar_texto_pdf(v)[:24], border=1, align="C")
            pdf.ln()

    dibujar_encabezado("(ORIGINAL PLANILLA)")
    pdf.set_font("Arial", "B", 8)
    pdf.cell(
        105,
        5.5,
        limpiar_texto_pdf(f"Operador 1: {datos_boleta['nombre_1']} ({datos_boleta['cod_1']})"),
        border=1,
    )
    pdf.cell(
        45,
        5.5,
        limpiar_texto_pdf(f"Fecha Sol.: {datos_boleta['fecha_sol']}"),
        border=1,
    )
    pdf.cell(
        46,
        5.5,
        limpiar_texto_pdf(f"Patio: {datos_boleta['patio']}"),
        border=1,
        ln=1,
    )

    modalidad_txt = (
        f"Modalidad: {datos_boleta['modalidad']}"
        + (
            f" | Operador 2: {datos_boleta['nombre_2']} ({datos_boleta['cod_2']})"
            if datos_boleta["es_intercambio"]
            else ""
        )
    )
    pdf.set_font("Arial", "", 8)
    pdf.cell(0, 5.5, limpiar_texto_pdf(modalidad_txt), border=1, ln=1)

    dibujar_tabla_detalle(
        "HORARIO PROGRAMADO OPERADOR 1", datos_boleta["prog_op1"]
    )
    dibujar_tabla_detalle(
        "HORARIO DE CAMBIO OPERADOR 1", datos_boleta["cambio_op1"]
    )

    if datos_boleta["es_intercambio"]:
        dibujar_tabla_detalle(
            "HORARIO PROGRAMADO OPERADOR 2", datos_boleta["prog_op2"]
        )
        dibujar_tabla_detalle(
            "HORARIO DE CAMBIO OPERADOR 2", datos_boleta["cambio_op2"]
        )

    pdf.ln(2)
    pdf.set_font("Arial", "", 7.5)
    pdf.cell(65, 10, "Firma Operador 1: _______________________", border=1, align="C")
    pdf.cell(
        65,
        10,
        "Firma Operador 2: _______________________"
        if datos_boleta["es_intercambio"]
        else "Firma Operador 2: N/A",
        border=1,
        align="C",
    )
    pdf.cell(
        66,
        10,
        limpiar_texto_pdf(f"Analista: {datos_boleta['responsable']}"),
        border=1,
        ln=1,
        align="C",
    )

    pdf.ln(4)
    pdf.set_font("Arial", "I", 7)
    pdf.cell(0, 4, "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -", ln=1, align="C")
    pdf.ln(2)

    dibujar_encabezado("(COPIA OPERADOR 1)")
    pdf.set_font("Arial", "B", 8)
    pdf.cell(
        130,
        5.5,
        limpiar_texto_pdf(f"Operador 1: {datos_boleta['nombre_1']} ({datos_boleta['cod_1']})"),
        border=1,
    )
    pdf.cell(
        66,
        5.5,
        limpiar_texto_pdf(f"Fecha Aut.: {datos_boleta['fecha_sol']}"),
        border=1,
        ln=1,
    )
    dibujar_tabla_detalle(
        "NUEVO HORARIO ASIGNADO - OPERADOR 1", datos_boleta["cambio_op1"]
    )
    pdf.set_font("Arial", "", 7.5)
    pdf.cell(98, 9, "Firma Operador 1: ___________________________", border=1, align="C")
    pdf.cell(
        98,
        9,
        limpiar_texto_pdf(f"Autorizado por: {datos_boleta['responsable']}"),
        border=1,
        ln=1,
        align="C",
    )

    if datos_boleta["es_intercambio"]:
        pdf.ln(4)
        dibujar_encabezado("(COPIA OPERADOR 2)")
        pdf.set_font("Arial", "B", 8)
        pdf.cell(
            130,
            5.5,
            limpiar_texto_pdf(f"Operador 2: {datos_boleta['nombre_2']} ({datos_boleta['cod_2']})"),
            border=1,
        )
        pdf.cell(
            66,
            5.5,
            limpiar_texto_pdf(f"Fecha Aut.: {datos_boleta['fecha_sol']}"),
            border=1,
            ln=1,
        )
        dibujar_tabla_detalle(
            "NUEVO HORARIO ASIGNADO - OPERADOR 2", datos_boleta["cambio_op2"]
        )
        pdf.set_font("Arial", "", 7.5)
        pdf.cell(98, 9, "Firma Operador 2: ___________________________", border=1, align="C")
        pdf.cell(
            98,
            9,
            limpiar_texto_pdf(f"Autorizado por: {datos_boleta['responsable']}"),
            border=1,
            ln=1,
            align="C",
        )

    out = pdf.output(dest="S")
    return out.encode("latin-1") if isinstance(out, str) else bytes(out)


# ==========================================
# SIDEBAR: CARGA DEL ARCHIVO MAESTRO
# ==========================================
st.sidebar.header("📂 1. Archivo Maestro Excel")
archivo_maestro = st.sidebar.file_uploader(
    "Sube CAMBIO DE HORARIOS 2026 (.xlsx / .xlsm)",
    type=["xlsx", "xlsm"],
    key="uploader_cambios_maestro",
)

if archivo_maestro is not None:
    df_row_data, df_hist_inicial, lista_responsables = cargar_libro_excel(
        archivo_maestro.getvalue()
    )
    if "historial_cambios" not in st.session_state:
        st.session_state.historial_cambios = df_hist_inicial.copy()
    st.sidebar.success(
        f"✅ ROW-DATA: {len(df_row_data):,} filas | Historial: {len(st.session_state.historial_cambios)} registros"
    )
else:
    df_row_data = pd.DataFrame()
    lista_responsables = list(RESPONSABLES_DEFAULT)
    if "historial_cambios" not in st.session_state:
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
    st.sidebar.info(
        "ℹ️ Puedes trabajar en modo manual o subir tu archivo Excel para autocompletar nombres y turnos desde ROW-DATA."
    )

# ==========================================
# PESTAÑAS PRINCIPALES
# ==========================================
tab_registro, tab_boleta, tab_historial = st.tabs(
    [
        "📝 1. Registrar Cambio de Horario",
        "🖨️ 2. Boleta F-ZEN-GO-07 (PDF / Imprimir)",
        "📊 3. Historial y Auditoría",
    ]
)

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
        responsable = st.selectbox("Responsable del Cambio", lista_responsables)
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
        cod_op1 = st.text_input("Código Operador 1", value="82601").strip()
        nom_op1, patio_op1 = buscar_operador_info(df_row_data, cod_op1)
        st.info(f"👤 **Operador 1:** {nom_op1} | **Patio:** {patio_op1}")

    cod_op2, nom_op2 = "", ""
    if es_intercambio:
        with c_op2:
            cod_op2 = st.text_input("Código Operador 2", value="82602").strip()
            nom_op2, _ = buscar_operador_info(df_row_data, cod_op2)
            st.info(f"👤 **Operador 2:** {nom_op2}")

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

        prog1 = buscar_programacion_dia(df_row_data, cod_op1, fecha_d)
        filas_prog_op1.append(
            {
                "fecha": fecha_d.strftime("%d/%m/%Y"),
                "codigo": cod_op1,
                **prog1,
            }
        )

        if es_intercambio:
            prog2 = buscar_programacion_dia(df_row_data, cod_op2, fecha_d)
            filas_prog_op2.append(
                {
                    "fecha": fecha_d.strftime("%d/%m/%Y"),
                    "codigo": cod_op2,
                    **prog2,
                }
            )

            col_i1, col_i2 = st.columns(2)
            with col_i1:
                st.caption(
                    f"Programado Op 1 ({cod_op1}): `{prog1['horario']}` | Sercon: `{prog1['sercon']}` | {prog1['p_inicio']} -> {prog1['p_fin']}"
                )
            with col_i2:
                st.caption(
                    f"Programado Op 2 ({cod_op2}): `{prog2['horario']}` | Sercon: `{prog2['sercon']}` | {prog2['p_inicio']} -> {prog2['p_fin']}"
                )

            filas_cambio_op1.append(
                {
                    "fecha": fecha_d.strftime("%d/%m/%Y"),
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
                    "codigo": cod_op2,
                    "horario": prog1["horario"],
                    "sercon": prog1["sercon"],
                    "p_inicio": prog1["p_inicio"],
                    "p_fin": prog1["p_fin"],
                }
            )
        else:
            st.caption(
                f"Programado actual ({cod_op1}): `{prog1['horario']}` | Sercon: `{prog1['sercon']}` | {prog1['p_inicio']} -> {prog1['p_fin']}"
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
                    "codigo": cod_op1,
                    "horario": nuevo_hor,
                    "sercon": nueva_tarea,
                    "p_inicio": nuevo_pi,
                    "p_fin": nuevo_pf,
                }
            )

    st.markdown("---")
    if st.button(
        "✅ Guardar Cambio y Preparar Boleta F-ZEN-GO-07",
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
            hora_cambio_resumen = f"{filas_prog_op1[0]['horario']}/{filas_cambio_op1[0]['horario']}"

            df_h_actual = st.session_state.historial_cambios
            ya_existe = False
            if not df_h_actual.empty:
                dups = df_h_actual[
                    (df_h_actual["COGIDO"].astype(str).str.strip() == codigo_registro)
                    & (
                        df_h_actual["FECHA DE CAMBIO"].astype(str).str.strip()
                        == str(
                            datetime.strptime(
                                fecha_cambio_principal, "%d/%m/%Y"
                            ).date()
                        )
                    )
                    & (
                        df_h_actual["HORA DE CAMBIO"].astype(str).str.strip()
                        == hora_cambio_resumen
                    )
                ]
                ya_existe = not dups.empty

            if ya_existe:
                st.warning(
                    "⚠️ Este cambio ya se encuentra registrado en el historial (duplicado bloqueado). Puedes imprimir su boleta en la pestaña 2."
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
                    f"✅ Cambio guardado exitosamente para **{codigo_registro}**. Ve a la pestaña **2. Boleta F-ZEN-GO-07** para descargar el PDF o imprimir."
                )

            st.session_state.boleta_activa = {
                "modalidad": modalidad,
                "es_intercambio": es_intercambio,
                "fecha_sol": fecha_solicitud.strftime("%d/%m/%Y"),
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

with tab_boleta:
    st.subheader("🖨️ Emisión de Formulario F-ZEN-GO-07")

    if "boleta_activa" not in st.session_state:
        st.info(
            "ℹ️ Registra primero un cambio en la pestaña **1. Registrar Cambio de Horario** para generar aquí su boleta oficial."
        )
    else:
        bol = st.session_state.boleta_activa
        pdf_bytes = generar_pdf_boleta(bol)

        c_btn1, c_btn2 = st.columns([1.5, 2])
        with c_btn1:
            st.download_button(
                label="📄 Descargar Boleta Oficial en PDF (Lista para Imprimir)",
                data=pdf_bytes,
                file_name=f"Cambio_Horario_{bol['cod_1']}_{bol['fecha_sol'].replace('/', '')}.pdf",
                mime="application/pdf",
                type="primary",
                use_container_width=True,
            )
        with c_btn2:
            st.info(
                "💡 También puedes presionar **Ctrl + P** en esta vista para imprimir directamente desde el navegador sin menús laterales."
            )

        st.markdown("---")
        st.markdown(
            f"""
            <div class="boleta-box">
                <h4 style="margin:0; text-align:center;">TRANSPORTE MASIVO DE PANAMÁ, S.A. — CAMBIO DE HORARIO (F-ZEN-GO-07)</h4>
                <p style="margin:4px 0; font-size:13px;">
                    <b>Operador 1:</b> {bol['nombre_1']} ({bol['cod_1']}) &nbsp;|&nbsp;
                    <b>Patio:</b> {bol['patio']} &nbsp;|&nbsp;
                    <b>Fecha Solicitud:</b> {bol['fecha_sol']}
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

        st.markdown("##### Horario Programado Operador 1")
        st.dataframe(pd.DataFrame(bol["prog_op1"]), use_container_width=True, hide_index=True)

        st.markdown("##### Horario de Cambio Operador 1")
        st.dataframe(pd.DataFrame(bol["cambio_op1"]), use_container_width=True, hide_index=True)

        if bol["es_intercambio"]:
            st.markdown(f"##### Horario Programado Operador 2 ({bol['nombre_2']})")
            st.dataframe(pd.DataFrame(bol["prog_op2"]), use_container_width=True, hide_index=True)

            st.markdown(f"##### Horario de Cambio Operador 2 ({bol['nombre_2']})")
            st.dataframe(pd.DataFrame(bol["cambio_op2"]), use_container_width=True, hide_index=True)

with tab_historial:
    st.subheader("📊 Historial de Cambios Registrados")
    df_h_show = st.session_state.historial_cambios

    col_f1, col_f2 = st.columns(2)
    with col_f1:
        filtro_op = st.text_input("Filtrar por Código de Operador:", "")
    with col_f2:
        filtro_resp = st.multiselect(
            "Filtrar por Responsable:",
            options=sorted(df_h_show["RESPONSABLE DEL CAMBIO"].dropna().unique().tolist())
            if not df_h_show.empty
            else [],
        )

    df_filtrado = df_h_show.copy()
    if filtro_op.strip():
        df_filtrado = df_filtrado[
            df_filtrado["COGIDO"].astype(str).str.contains(filtro_op.strip(), case=False)
        ]
    if filtro_resp:
        df_filtrado = df_filtrado[
            df_filtrado["RESPONSABLE DEL CAMBIO"].isin(filtro_resp)
        ]

    st.dataframe(df_filtrado, use_container_width=True, hide_index=True)

    buffer_xlsx = io.BytesIO()
    with pd.ExcelWriter(buffer_xlsx, engine="openpyxl") as writer:
        df_h_show.to_excel(
            writer, sheet_name="HISTORIAL DE CAMBIOS", index=False
        )
    buffer_xlsx.seek(0)

    st.download_button(
        label="📥 Descargar Historial Actualizado (Excel .xlsx)",
        data=buffer_xlsx,
        file_name=f"Historial_Cambios_Horario_{date.today().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
