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
    if client is None:
        return None, "Falta [gcp_service_account] en Secrets."
    if "SHEET_ID" not in st.secrets or not str(st.secrets["SHEET_ID"]).strip():
        return None, "Falta SHEET_ID en Secrets."

    spreadsheet = client.open_by_key(str(st.secrets["SHEET_ID"]).strip())
    nombres_hojas = [ws.title for ws in spreadsheet.worksheets()]

    if NOMBRE_HOJA_PERMISOS in nombres_hojas:
        return spreadsheet.worksheet(NOMBRE_HOJA_PERMISOS), "OK"
    if "Registro" in nombres_hojas:
        return spreadsheet.worksheet("Registro"), "OK"

    ws_nueva = spreadsheet.add_worksheet(
        title=NOMBRE_HOJA_PERMISOS, rows=1000, cols=len(COLUMNAS_PERMISOS)
    )
    ws_nueva.append_row(COLUMNAS_PERMISOS)
    return ws_nueva, "OK"


@st.cache_data(ttl=600, show_spinner="Cargando directorio desde Google Drive...")
def cargar_directorio_desde_drive():
    """
    Lee automáticamente el directorio de colaboradores desde Google Drive:
    1. Si existe SHEET_DIRECTORIO_ID en Secrets, abre ese Google Sheet.
    2. Si no, busca la pestaña 'DIRECTORIO' o 'Directorio Consolidado' en el mismo SHEET_ID.
    """
    try:
        client = crear_cliente_gspread()
        if client is None:
            return {}, False, "Sin credenciales gcp_service_account"

        sheet_key = str(
            st.secrets.get("SHEET_DIRECTORIO_ID") or st.secrets.get("SHEET_ID") or ""
        ).strip()
        if not sheet_key:
            return {}, False, "Sin SHEET_DIRECTORIO_ID o SHEET_ID configurado"

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
    """Lee el historial acumulado de permisos sin fallar por columnas vacías o duplicadas."""
    try:
        ws, msg = obtener_hoja_permisos_gsheets()
        if ws is None:
            return pd.DataFrame(columns=COLUMNAS_PERMISOS), False, msg

        valores = ws.get_all_values()
        if not valores or len(valores) <= 1:
            return (
                pd.DataFrame(columns=COLUMNAS_PERMISOS),
                True,
                f"Pestaña '{ws.title}' lista (0 registros)",
            )

        encabezados = [str(c).strip() for c in valores[0]]
        filas = valores[1:]
        df_p = pd.DataFrame(filas, columns=encabezados)
        # Eliminar columnas vacías en encabezado si existen
        df_p = df_p.loc[:, df_p.columns != ""].copy()

        if "Codigo" in df_p.columns:
            df_p = df_p[df_p["Codigo"].astype(str).str.strip() != ""].copy()
            df_p["Codigo"] = df_p["Codigo"].apply(limpiar_codigo)

        return df_p, True, f"Pestaña '{ws.title}' conectada"
    except Exception as e:
        return pd.DataFrame(columns=COLUMNAS_PERMISOS), False, f"Detalle técnico: {e}"


def agregar_permiso_en_gsheets(fila_dict: dict):
    """Inserta el nuevo permiso en la fila 2 de la pestaña HISTORIAL DE PERMISOS."""
    try:
        ws, msg = obtener_hoja_permisos_gsheets()
        if ws is None:
            return False, msg

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
    if sub.empty:
        st.markdown(
            f'<div class="panel-lateral-anual-ok">✅ <b>Historial {anio_ref} — Cód. {codigo}</b><br>Lleva <b>0 permisos/licencias</b> registrados.</div>',
            unsafe_allow_html=True,
        )
        return

    sub["F_SAL_DT"] = pd.to_datetime(sub["Fecha Salida"], errors="coerce")
    sub_anio = sub[sub["F_SAL_DT"].dt.year == anio_ref]
    tot_anio = len(sub_anio)
    dias_anio = int(pd.to_numeric(sub_anio.get("TOTAL DIAS", 1), errors="coerce").fillna(1).sum())
    tot_hist = len(sub)

    if tot_anio > 0:
        tipos_res = sub_anio["Tipo de Solicitud"].value_counts().to_dict()
        detalle_tipos = " · ".join([f"{v} {k}" for k, v in tipos_res.items()])
        st.markdown(
            f'<div class="panel-lateral-anual">'
            f'⚠️ <b>Resumen {anio_ref} — Cód. {codigo}</b><br>'
            f'Lleva <b>{tot_anio} solicitud(es)</b> ({dias_anio} día(s) total):<br>'
            f'<b>{detalle_tipos}</b>'
            f'</div>',
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            f'<div class="panel-lateral-anual-ok">'
            f'✅ <b>Resumen {anio_ref} — Cód. {codigo}</b><br>'
            f'0 registros en {anio_ref} (Acumulado histórico: {tot_hist}).'
            f'</div>',
            unsafe_allow_html=True,
        )


# ==========================================
# GENERADOR PDF IDÉNTICO A F-CORP-GH-PLA-25
# ==========================================
def generar_pdf_f_corp_gh_pla_25(bol: dict) -> bytes:
    """
    Genera el PDF Letter con el formato oficial F-CORP-GH-PLA-25/V1-07-2013
    (Original arriba + Copia idéntica abajo).
    """
    pdf = FPDF(orientation="P", unit="mm", format="Letter")
    pdf.set_auto_page_break(auto=False)
    pdf.add_page()

    azul_hdr = (0, 112, 192)
    gris_lbl = (235, 238, 242)

    def dibujar_logo_mibus(x, y):
        pdf.set_fill_color(0, 112, 192)
        pdf.rect(x, y, 13, 2.3, style="F")
        pdf.set_fill_color(0, 176, 80)
        pdf.rect(x + 2, y + 2.8, 13, 2.3, style="F")
        pdf.set_text_color(0, 51, 102)
        pdf.set_font("Arial", "B", 8.5)
        pdf.set_xy(x, y + 5.5)
        pdf.cell(16, 3.5, "MiBus", align="C")
        pdf.set_text_color(0, 0, 0)

    def dibujar_check(x, y, activo: bool, texto: str, ancho_txt: float = 34):
        pdf.set_xy(x, y)
        pdf.set_font("Arial", "B", 6.5)
        pdf.cell(3.2, 3.2, "X" if activo else "", border=1, align="C")
        pdf.set_xy(x + 3.8, y - 0.2)
        pdf.set_font("Arial", "", 6.2)
        pdf.cell(ancho_txt, 3.6, limpiar_texto_pdf(texto), align="L")

    def dibujar_formulario(y_top: float):
        pdf.set_draw_color(0, 0, 0)
        pdf.set_line_width(0.35)
        pdf.rect(10, y_top, 196, 121)

        dibujar_logo_mibus(13, y_top + 2.2)
        pdf.set_xy(32, y_top + 3)
        pdf.set_font("Arial", "B", 11)
        pdf.cell(125, 6, "REPORTE NOVEDAD AUSENTISMO", align="C")
        pdf.set_xy(158, y_top + 2)
        pdf.set_font("Arial", "", 6.5)
        pdf.cell(46, 4, "F-CORP-GH-PLA-25/V1-07-2013", align="R")

        y1 = y_top + 12
        pdf.set_fill_color(*gris_lbl)
        pdf.set_font("Arial", "B", 6.5)
        pdf.set_xy(10, y1)
        pdf.cell(28, 4.2, "Codigo Colaborador", border=1, align="C", fill=True)
        pdf.cell(88, 4.2, "Nombre y Apellido del Colaborador", border=1, align="C", fill=True)
        pdf.cell(44, 4.2, "Cargo", border=1, align="C", fill=True)
        pdf.cell(36, 4.2, "Fecha Solicitud", border=1, align="C", fill=True)

        y1_val = y1 + 4.2
        pdf.set_xy(10, y1_val)
        pdf.set_font("Arial", "B", 8)
        pdf.cell(28, 7.5, limpiar_texto_pdf(bol["codigo"]), border=1, align="C")
        pdf.cell(88, 7.5, limpiar_texto_pdf(bol["nombre"])[:44], border=1, align="C")
        pdf.set_font("Arial", "", 7.5)
        pdf.cell(44, 7.5, limpiar_texto_pdf(bol["cargo"]), border=1, align="C")

        pdf.set_font("Arial", "B", 5.8)
        pdf.set_xy(170, y1_val)
        pdf.cell(12, 3.2, "DD", border=1, align="C", fill=True)
        pdf.cell(12, 3.2, "MM", border=1, align="C", fill=True)
        pdf.cell(12, 3.2, "AAAA", border=1, align="C", fill=True)

        f_sol = bol["fecha_sol_obj"]
        pdf.set_font("Arial", "", 7)
        pdf.set_xy(170, y1_val + 3.2)
        pdf.cell(12, 4.3, f"{f_sol.day:02d}", border=1, align="C")
        pdf.cell(12, 4.3, f"{f_sol.month:02d}", border=1, align="C")
        pdf.cell(12, 4.3, f"{f_sol.year}", border=1, align="C")

        y2 = y1_val + 7.5
        pdf.set_font("Arial", "B", 6.5)
        pdf.set_xy(10, y2)
        pdf.cell(68, 4.2, "Sucursal", border=1, align="C", fill=True)
        pdf.cell(64, 4.2, "Centro de Costo", border=1, align="C", fill=True)
        pdf.cell(64, 4.2, "Departamento", border=1, align="C", fill=True)

        y2_val = y2 + 4.2
        pdf.set_font("Arial", "", 7)
        pdf.set_xy(10, y2_val)
        pdf.cell(68, 5, limpiar_texto_pdf(bol["sucursal"]), border=1, align="C")
        pdf.cell(64, 5, limpiar_texto_pdf(bol["centro_costo"]), border=1, align="C")
        pdf.cell(64, 5, limpiar_texto_pdf(bol["departamento"]), border=1, align="C")

        y3 = y2_val + 5
        pdf.set_fill_color(*azul_hdr)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Arial", "B", 7)
        pdf.set_xy(10, y3)
        pdf.cell(132, 4.5, "Tipo de Solicitud", border=1, align="C", fill=True)
        pdf.cell(64, 4.5, "Periodo del Ausentismo", border=1, align="C", fill=True)
        pdf.set_text_color(0, 0, 0)

        y3_body = y3 + 4.5
        pdf.set_xy(10, y3_body)
        pdf.cell(132, 13.5, "", border=1)

        tipo_sel = bol["tipo_solicitud"]
        dibujar_check(13, y3_body + 2.0, tipo_sel == "Permiso Remunerado", "Permiso Remunerado", 28)
        dibujar_check(46, y3_body + 2.0, tipo_sel == "Permiso No Remunerado", "Permiso No Remunerado", 31)
        dibujar_check(82, y3_body + 2.0, tipo_sel == "Licencia Remunerada", "Licencia Remunerada", 27)
        dibujar_check(112, y3_body + 2.0, tipo_sel == "Licencia No Remunerada", "Licencia No Remunerada", 28)
        dibujar_check(13, y3_body + 7.8, tipo_sel == "Compensatorio", "Compensatorio", 32)

        pdf.set_fill_color(*gris_lbl)
        pdf.set_font("Arial", "B", 6.2)
        pdf.set_xy(142, y3_body)
        pdf.cell(32, 4.2, "Fecha de Salida", border=1, align="C", fill=True)
        pdf.cell(32, 4.2, "Fecha de Ingreso", border=1, align="C", fill=True)

        pdf.set_font("Arial", "B", 5.5)
        pdf.set_xy(142, y3_body + 4.2)
        for _ in range(2):
            pdf.cell(10, 3.8, "DD", border=1, align="C", fill=True)
            pdf.cell(10, 3.8, "MM", border=1, align="C", fill=True)
            pdf.cell(12, 3.8, "AAAA", border=1, align="C", fill=True)

        f_sal = bol["fecha_salida_obj"]
        f_ing = bol["fecha_ingreso_obj"]
        pdf.set_font("Arial", "B", 7)
        pdf.set_xy(142, y3_body + 8.0)
        pdf.cell(10, 5.5, f"{f_sal.day:02d}", border=1, align="C")
        pdf.cell(10, 5.5, f"{f_sal.month:02d}", border=1, align="C")
        pdf.cell(12, 5.5, f"{f_sal.year}", border=1, align="C")
        pdf.cell(10, 5.5, f"{f_ing.day:02d}", border=1, align="C")
        pdf.cell(10, 5.5, f"{f_ing.month:02d}", border=1, align="C")
        pdf.cell(12, 5.5, f"{f_ing.year}", border=1, align="C")

        y4 = y3_body + 13.5
        pdf.set_fill_color(*azul_hdr)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Arial", "B", 7)
        pdf.set_xy(10, y4)
        pdf.cell(196, 4.5, "Motivo Solicitud", border=1, align="C", fill=True)
        pdf.set_text_color(0, 0, 0)

        y4_body = y4 + 4.5
        pdf.set_xy(10, y4_body)
        pdf.cell(196, 12.5, "", border=1)

        mot_sel = bol["motivo_solicitud"]
        dibujar_check(13, y4_body + 1.8, mot_sel == "Enfermedad", "Enfermedad", 22)
        dibujar_check(42, y4_body + 1.8, mot_sel == "Cita Médica", "Cita Medica", 22)
        dibujar_check(70, y4_body + 1.8, mot_sel == "Fallecimiento", "Fallecimiento", 22)
        dibujar_check(98, y4_body + 1.8, mot_sel == "Nacimiento hijo (a)", "Nacimiento hijo (a)", 28)
        dibujar_check(133, y4_body + 1.8, mot_sel == "Matrimonio", "Matrimonio", 20)
        dibujar_check(160, y4_body + 1.8, mot_sel == "Asuntos Personales", "Asuntos Personales", 35)
        dibujar_check(
            13,
            y4_body + 7.2,
            mot_sel == "Representar a la empresa, el Estado o el país",
            "Representar a la empresa, el Estado o el pais",
            70,
        )
        dibujar_check(98, y4_body + 7.2, mot_sel == "Permiso Sindical", "Permiso Sindical", 30)
        dibujar_check(133, y4_body + 7.2, mot_sel == "Compensatorio", "Compensatorio", 30)

        y5 = y4_body + 12.5
        pdf.set_xy(10, y5)
        pdf.cell(142, 22, "", border=1)
        pdf.cell(54, 22, "", border=1)

        pdf.set_xy(12, y5 + 1.2)
        pdf.set_font("Arial", "B", 6.5)
        pdf.cell(138, 3.8, "Explique de forma clara y breve el motivo de esta solicitud:", ln=1)
        pdf.set_x(12)
        pdf.set_font("Arial", "", 6.8)
        pdf.multi_cell(138, 3.6, limpiar_texto_pdf(bol["observaciones"])[:290])

        pdf.set_xy(152, y5 + 17.5)
        pdf.set_font("Arial", "B", 6.5)
        pdf.cell(54, 3.5, "Firma del Solicitante", border="T", align="C")

        y6 = y5 + 22
        pdf.set_fill_color(*gris_lbl)
        pdf.set_font("Arial", "B", 7)
        pdf.set_xy(10, y6)
        pdf.cell(196, 4.5, "Aprobo", border=1, align="C", fill=True)

        y6_hdr = y6 + 4.5
        anchos_ap = [49, 49, 49, 49]
        titulos_ap = [
            "JEFE INMEDIATO",
            "PROFESIONAL DE GESTION HUMANA",
            "DIRECTOR DE AREA/ZEN",
            "GERENTE DE RELACIONES LABORALES",
        ]
        pdf.set_font("Arial", "B", 5.6)
        pdf.set_xy(10, y6_hdr)
        for w, t in zip(anchos_ap, titulos_ap):
            pdf.cell(w, 4.2, t, border=1, align="C", fill=True)

        y6_body = y6_hdr + 4.2
        nombres_ap = [
            bol["jefe_inmediato"],
            bol["prof_gh"],
            bol["director_area"],
            bol["gerente_rrll"],
        ]
        cargos_ap = [
            "Cargo: Jefe de Operaciones",
            "",
            "",
            "",
        ]
        for i, w in enumerate(anchos_ap):
            x_col = 10 + i * w
            pdf.set_xy(x_col, y6_body)
            pdf.cell(w, 16.5, "", border=1)

            pdf.set_xy(x_col + 1.5, y6_body + 1)
            pdf.set_font("Arial", "B", 6)
            pdf.cell(w - 3, 3.5, limpiar_texto_pdf(f"Nombre: {nombres_ap[i]}"), ln=1)

            if cargos_ap[i]:
                pdf.set_x(x_col + 1.5)
                pdf.set_font("Arial", "", 5.8)
                pdf.cell(w - 3, 3.5, limpiar_texto_pdf(cargos_ap[i]))

            pdf.set_xy(x_col + 1.5, y6_body + 12.2)
            pdf.set_font("Arial", "B", 6)
            pdf.cell(w - 3, 3.5, "Firma: _______________________")

    dibujar_formulario(8)

    pdf.set_xy(10, 132.5)
    pdf.set_font("Arial", "", 6.5)
    pdf.cell(196, 3, "- " * 90, align="C")

    dibujar_formulario(138)

    out = pdf.output(dest="S")
    return out.encode("latin-1") if isinstance(out, str) else bytes(out)


# ==========================================
# INICIALIZAR ESTADO Y SINCRONIZAR DRIVE AUTOMÁTICAMENTE
# ==========================================
if "directorio_drive" not in st.session_state:
    dir_map_gs, ok_dir_gs, msg_dir_gs = cargar_directorio_desde_drive()
    st.session_state.directorio_drive = dir_map_gs
    st.session_state.directorio_drive_ok = ok_dir_gs
    st.session_state.directorio_drive_msg = msg_dir_gs

if "permisos_sincronizados_gs" not in st.session_state:
    df_p_gs, ok_p_gs, msg_p_gs = leer_permisos_desde_gsheets()
    st.session_state.historial_permisos = df_p_gs
    st.session_state.permisos_en_gsheets = ok_p_gs
    st.session_state.permisos_msg_gs = msg_p_gs
    st.session_state.permisos_sincronizados_gs = True


# ==========================================
# SIDEBAR: ESTADO DE GOOGLE SHEETS Y DIRECTORIO EN DRIVE
# ==========================================
st.sidebar.header("☁️ 1. Historial de Permisos")
if st.session_state.permisos_en_gsheets:
    st.sidebar.success(
        f"🟢 Conectado a Google Sheets\n\n"
        f"📑 Pestaña: **{NOMBRE_HOJA_PERMISOS}**\n"
        f"📊 **{len(st.session_state.historial_permisos)}** registros sincronizados."
    )
else:
    st.sidebar.warning(
        f"⚠️ No se pudo leer `SHEET_ID`:\n\n{st.session_state.get('permisos_msg_gs', '')}"
    )

if st.sidebar.button("🔄 Recargar Permisos", use_container_width=True):
    df_p_gs, ok_p_gs, msg_p_gs = leer_permisos_desde_gsheets()
    st.session_state.historial_permisos = df_p_gs
    st.session_state.permisos_en_gsheets = ok_p_gs
    st.session_state.permisos_msg_gs = msg_p_gs
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.header("🗂️ 2. Directorio desde Drive")
if st.session_state.directorio_drive_ok:
    st.sidebar.success(
        f"🟢 Directorio activo desde Drive\n\n"
        f"👥 **{st.session_state.directorio_drive_msg}**"
    )
else:
    st.sidebar.info(
        f"ℹ️ {st.session_state.directorio_drive_msg}.\n\n"
        "Puedes agregar una pestaña llamada **DIRECTORIO** en tu Google Sheet o definir `SHEET_DIRECTORIO_ID` en Secrets."
    )

if st.sidebar.button("🔄 Actualizar Directorio Drive", use_container_width=True):
    cargar_directorio_desde_drive.clear()
    dir_map_gs, ok_dir_gs, msg_dir_gs = cargar_directorio_desde_drive()
    st.session_state.directorio_drive = dir_map_gs
    st.session_state.directorio_drive_ok = ok_dir_gs
    st.session_state.directorio_drive_msg = msg_dir_gs
    st.rerun()


# ==========================================
# ENCABEZADO PRINCIPAL Y PESTAÑAS
# ==========================================
st.markdown(
    '<div class="header-permisos">📋 MÓDULO PERMISOS Y LICENCIAS — REPORTE NOVEDAD AUSENTISMO (F-CORP-GH-PLA-25)</div>',
    unsafe_allow_html=True,
)

tab_reg_p, tab_bol_p, tab_hist_p = st.tabs(
    [
        "📝 1. Registro de Permiso / Licencia",
        "🖨️ 2. Boleta Oficial F-CORP-GH-PLA-25 (PDF)",
        "📊 3. Historial Acumulado y Estadísticas",
    ]
)

# ------------------------------------------
# TAB 1: FORMULARIO DE REGISTRO
# ------------------------------------------
with tab_reg_p:
    st.markdown(
        '<div class="sub-header-permisos">1. DATOS DEL COLABORADOR</div>',
        unsafe_allow_html=True,
    )
    c1, c2, c3 = st.columns([1, 1.8, 1.8])
    with c1:
        cod_colab = limpiar_codigo(st.text_input("Código Colaborador", value="81085"))
    with c2:
        nom_colab, suc_default, ok_colab = buscar_colaborador_permiso(cod_colab)
        st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
        if ok_colab:
            st.success(f"👤 **{nom_colab}**")
        else:
            st.warning(f"👤 **{nom_colab}**")
    with c3:
        renderizar_resumen_anual_permisos(
            cod_colab,
            st.session_state.historial_permisos,
            date.today().year,
        )

    if not ok_colab:
        nom_colab = st.text_input(
            "Ingresa manualmente Nombre y Apellido del Colaborador:",
            placeholder="APELLIDO, NOMBRE",
        ).upper()

    c_org1, c_org2, c_org3, c_org4 = st.columns([1.2, 1.1, 1.1, 1])
    with c_org1:
        idx_suc = (
            SUCURSALES_CATALOGO.index(suc_default)
            if suc_default in SUCURSALES_CATALOGO
            else 0
        )
        sucursal = st.selectbox("Sucursal", SUCURSALES_CATALOGO, index=idx_suc)
    with c_org2:
        cargo = st.text_input("Cargo", value="Operador")
    with c_org3:
        centro_costo = st.text_input("Centro de Costo", value="2101 - Operadores")
    with c_org4:
        departamento = st.text_input("Departamento", value="OPERACIONES")

    st.markdown(
        '<div class="sub-header-permisos">2. DETALLE DE LA SOLICITUD Y FECHAS</div>',
        unsafe_allow_html=True,
    )
    c_sol1, c_sol2, c_sol3 = st.columns([1.3, 1.4, 1.2])
    with c_sol1:
        tipo_solicitud = st.selectbox("Tipo de Solicitud", TIPOS_SOLICITUD)
    with c_sol2:
        motivo_solicitud = st.selectbox("Motivo Solicitud", MOTIVOS_SOLICITUD_BOLETA)
    with c_sol3:
        lista_resp = st.session_state.get("lista_responsables", RESPONSABLES_DEFAULT)
        hecho_por = st.selectbox("Hecho Por (Analista)", lista_resp)

    c_f1, c_f2, c_f3, c_f4 = st.columns([1, 1, 1, 0.9])
    with c_f1:
        fecha_solicitud = st.date_input("Fecha de Solicitud", value=date.today())
    with c_f2:
        fecha_salida = st.date_input("Fecha de Salida", value=date.today())
    with c_f3:
        fecha_ingreso = st.date_input("Fecha de Ingreso (Retorno)", value=date.today())
    with c_f4:
        total_dias = max(1, (fecha_ingreso - fecha_salida).days + 1)
        st.markdown("<div style='height: 26px;'></div>", unsafe_allow_html=True)
        st.markdown(
            f'<div class="dias-badge">TOTAL DÍAS: {total_dias}</div>',
            unsafe_allow_html=True,
        )

    if fecha_ingreso < fecha_salida:
        st.error("⚠️ La **Fecha de Ingreso** no puede ser anterior a la **Fecha de Salida**.")

    observaciones = st.text_area(
        "Explique de forma clara y breve el motivo de esta solicitud (Observaciones):",
        placeholder="Ej. PERMISO REMUNERADO POR CITA MÉDICA ANUNCIADA DE 10:00 A 13:00...",
        height=85,
    )

    with st.expander("✍️ Firmantes de Aprobación en Boleta (Prellenado Oficial)", expanded=False):
        ca1, ca2, ca3, ca4 = st.columns(4)
        jefe_inmediato = ca1.text_input("Jefe Inmediato", value="Gilberto Arauz")
        prof_gh = ca2.text_input("Profesional Gestión Humana", value="Laurys Abrego")
        director_area = ca3.text_input("Director de Área/ZEN", value="Mayrim Fraguela")
        gerente_rrll = ca4.text_input("Gerente Relaciones Laborales", value="")

    st.markdown("---")
    if st.button(
        "✅ Guardar Permiso en Google Sheets y Generar Boleta F-CORP-GH-PLA-25",
        type="primary",
        use_container_width=True,
    ):
        if not cod_colab:
            st.error("⚠️ Debes ingresar el código del colaborador.")
        elif fecha_ingreso < fecha_salida:
            st.error("⚠️ Corrige las fechas antes de guardar.")
        elif not observaciones.strip():
            st.error("⚠️ Escribe una breve explicación del motivo de la solicitud.")
        else:
            serial_excel = (fecha_salida - date(1899, 12, 30)).days
            key_registro = f"{serial_excel}{cod_colab}"
            obs_final = observaciones.strip().upper()

            df_p_actual = st.session_state.historial_permisos
            duplicado = False
            if not df_p_actual.empty and "Key" in df_p_actual.columns:
                duplicado = (
                    df_p_actual["Key"].astype(str).str.strip() == key_registro
                ).any() and (
                    df_p_actual["Observaciones"].astype(str).str.strip().str.upper()
                    == obs_final
                ).any()

            if duplicado:
                st.warning(
                    "⚠️ Este permiso ya se encuentra registrado en el historial. Puedes descargar su boleta en la pestaña 2."
                )
            else:
                nuevo_permiso = {
                    "Key": key_registro,
                    "Codigo": cod_colab,
                    "Tipo de Solicitud": tipo_solicitud,
                    "Fecha Salida": pd.Timestamp(fecha_salida),
                    "Fecha de Ingreso": pd.Timestamp(fecha_ingreso),
                    "Motivo Solicitud": motivo_solicitud,
                    "Observaciones": obs_final,
                    "TOTAL DIAS": total_dias,
                    "Hecho Por": hecho_por,
                    "Astamp": datetime.now().replace(microsecond=0),
                }

                df_p_act = pd.concat(
                    [pd.DataFrame([nuevo_permiso]), df_p_actual], ignore_index=True
                )
                st.session_state.historial_permisos = df_p_act

                ok_gs, msg_gs = agregar_permiso_en_gsheets(nuevo_permiso)
                if ok_gs:
                    st.success(
                        f"✅ {msg_gs} ({len(df_p_act)} registros totales). Ve a la pestaña **2. Boleta Oficial F-CORP-GH-PLA-25 (PDF)** para descargar."
                    )
                else:
                    st.warning(
                        f"✅ Registro guardado en pantalla, pero no se sincronizó con Google Sheets: {msg_gs}"
                    )

            st.session_state.boleta_permiso_activa = {
                "codigo": cod_colab,
                "nombre": nom_colab,
                "cargo": cargo,
                "sucursal": sucursal,
                "centro_costo": centro_costo,
                "departamento": departamento,
                "tipo_solicitud": tipo_solicitud,
                "motivo_solicitud": motivo_solicitud,
                "fecha_sol_obj": fecha_solicitud,
                "fecha_salida_obj": fecha_salida,
                "fecha_ingreso_obj": fecha_ingreso,
                "total_dias": total_dias,
                "observaciones": obs_final,
                "hecho_por": hecho_por,
                "jefe_inmediato": jefe_inmediato,
                "prof_gh": prof_gh,
                "director_area": director_area,
                "gerente_rrll": gerente_rrll,
            }

# ------------------------------------------
# TAB 2: BOLETA OFICIAL F-CORP-GH-PLA-25 (PDF)
# ------------------------------------------
with tab_bol_p:
    st.subheader("🖨️ Emisión de Reporte Novedad Ausentismo (F-CORP-GH-PLA-25)")

    if "boleta_permiso_activa" not in st.session_state:
        st.info(
            "ℹ️ Registra primero una solicitud en la pestaña **1. Registro de Permiso / Licencia** para generar su boleta PDF."
        )
    else:
        bol = st.session_state.boleta_permiso_activa
        pdf_bytes = generar_pdf_f_corp_gh_pla_25(bol)

        c_b1, c_b2 = st.columns([1.6, 2])
        with c_b1:
            st.download_button(
                label="📄 Descargar Boleta Oficial F-CORP-GH-PLA-25 (PDF)",
                data=pdf_bytes,
                file_name=f"F-CORP-GH-PLA-25_{bol['codigo']}_{bol['fecha_salida_obj'].strftime('%Y%m%d')}.pdf",
                mime="application/pdf",
                type="primary",
                use_container_width=True,
            )
        with c_b2:
            st.info(
                "💡 El PDF contiene **Original y Copia** en una sola hoja tamaño Carta lista para firmar."
            )

        st.markdown("---")
        st.markdown(
            f"""
            <div class="boleta-preview">
                <h4 style="margin:0; text-align:center; color:#0070c0;">REPORTE NOVEDAD AUSENTISMO — F-CORP-GH-PLA-25/V1-07-2013</h4>
                <p style="margin:6px 0; font-size:13px;">
                    <b>Colaborador:</b> {bol['nombre']} (<b>Cód. {bol['codigo']}</b>) &nbsp;|&nbsp;
                    <b>Sucursal:</b> {bol['sucursal']} &nbsp;|&nbsp;
                    <b>Fecha Solicitud:</b> {bol['fecha_sol_obj'].strftime('%d/%m/%Y')}
                </p>
                <p style="margin:6px 0; font-size:13px;">
                    <b>Tipo:</b> {bol['tipo_solicitud']} &nbsp;|&nbsp;
                    <b>Motivo:</b> {bol['motivo_solicitud']} &nbsp;|&nbsp;
                    <b>Salida:</b> {bol['fecha_salida_obj'].strftime('%d/%m/%Y')} &nbsp;→&nbsp;
                    <b>Ingreso:</b> {bol['fecha_ingreso_obj'].strftime('%d/%m/%Y')} (<b>{bol['total_dias']} día(s)</b>)
                </p>
                <p style="margin:6px 0; font-size:13px;">
                    <b>Explicación:</b> {bol['observaciones']}
                </p>
            </div>
            """,
            unsafe_allow_html=True,
        )

# ------------------------------------------
# TAB 3: HISTORIAL ACUMULADO Y ESTADÍSTICAS
# ------------------------------------------
with tab_hist_p:
    st.subheader("📊 Historial Acumulado de Permisos y Licencias")
    df_hp = st.session_state.historial_permisos.copy()

    if df_hp.empty:
        st.info("ℹ️ Aún no hay permisos registrados en Google Sheets.")
    else:
        df_hp["F_SAL_DT"] = pd.to_datetime(df_hp["Fecha Salida"], errors="coerce")
        df_hp["DIAS_NUM"] = pd.to_numeric(df_hp.get("TOTAL DIAS", 1), errors="coerce").fillna(1)

        k1, k2, k3 = st.columns(3)
        k1.metric("Total Solicitudes", f"{len(df_hp):,}")
        k2.metric("Total Días Ausentismo", f"{int(df_hp['DIAS_NUM'].sum()):,}")
        k3.metric(
            "Permisos No Remunerados",
            f"{(df_hp['Tipo de Solicitud'] == 'Permiso No Remunerado').sum():,}",
        )

        st.markdown("---")
        cg1, cg2 = st.columns(2)
        with cg1:
            st.markdown("##### 📑 Distribución por Tipo de Solicitud")
            res_t = df_hp["Tipo de Solicitud"].value_counts().reset_index()
            res_t.columns = ["Tipo de Solicitud", "Cantidad"]
            chart_t = (
                alt.Chart(res_t)
                .mark_bar(color="#0070c0", cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
                .encode(
                    x=alt.X("Tipo de Solicitud:N", sort="-y", axis=alt.Axis(labelAngle=-15)),
                    y=alt.Y("Cantidad:Q"),
                    tooltip=["Tipo de Solicitud", "Cantidad"],
                )
            )
            txt_t = chart_t.mark_text(dy=-5, color="#f8fafc", fontWeight="bold").encode(
                text="Cantidad:Q"
            )
            st.altair_chart((chart_t + txt_t).properties(height=280), use_container_width=True)

        with cg2:
            st.markdown("##### 📌 Distribución por Motivo")
            res_m = df_hp["Motivo Solicitud"].value_counts().reset_index()
            res_m.columns = ["Motivo Solicitud", "Cantidad"]
            chart_m = (
                alt.Chart(res_m)
                .mark_bar(color="#00b050", cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
                .encode(
                    x=alt.X("Motivo Solicitud:N", sort="-y", axis=alt.Axis(labelAngle=-15)),
                    y=alt.Y("Cantidad:Q"),
                    tooltip=["Motivo Solicitud", "Cantidad"],
                )
            )
            txt_m = chart_m.mark_text(dy=-5, color="#f8fafc", fontWeight="bold").encode(
                text="Cantidad:Q"
            )
            st.altair_chart((chart_m + txt_m).properties(height=280), use_container_width=True)

        st.markdown("---")
        cf1, cf2 = st.columns(2)
        f_cod = cf1.text_input("Filtrar por Código de Colaborador:", "")
        f_tipo = cf2.multiselect(
            "Filtrar por Tipo de Solicitud:",
            options=sorted(df_hp["Tipo de Solicitud"].dropna().astype(str).unique().tolist()),
        )

        df_show = df_hp.drop(columns=["F_SAL_DT", "DIAS_NUM"], errors="ignore").copy()
        if f_cod.strip():
            df_show = df_show[df_show["Codigo"].astype(str).str.contains(f_cod.strip(), case=False)]
        if f_tipo:
            df_show = df_show[df_show["Tipo de Solicitud"].astype(str).isin(f_tipo)]

        st.dataframe(df_show, use_container_width=True, hide_index=True)

        buf_p = io.BytesIO()
        with pd.ExcelWriter(buf_p, engine="openpyxl") as writer:
            df_show.to_excel(writer, sheet_name="Registro", index=False)
        buf_p.seek(0)

        st.download_button(
            label="📥 Descargar Historial de Permisos (Excel .xlsx)",
            data=buf_p,
            file_name=f"HISTORIAL_PERMISOS_{date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
