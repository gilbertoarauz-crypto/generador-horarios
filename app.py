from datetime import datetime, timedelta, time
import io
import re
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Generador de Horarios Pro", layout="wide")
st.title("📅 Generador de Horarios & Control de Tareas Operativas")

# ==========================================
# CATÁLOGO Y DEFAULTS DE TURNOS
# ==========================================
CATALOGO_TURNOS = [
    "03:00-11:00", "06:00-15:00", "07:00-16:00", "08:00-15:00 CAP",
    "08:00-17:00", "11:00-19:00", "11:00-19:00 AT", "11:00-27:00 AT",
    "19:00-27:00", "19:00-35:00 AT", "20:00-28:00", "22:00-30:00"
]

TURNOS_DEFAULT_POR_CARGO = {
    "ANALISTA DE OPERACIONES": {
        "habil": ["03:00-11:00", "11:00-19:00", "06:00-15:00", "19:00-27:00"],
        "sabado": ["03:00-11:00", "11:00-19:00", "19:00-27:00", "06:00-15:00"],
        "domingo": ["03:00-11:00", "11:00-19:00", "19:00-27:00"]
    },
    "AUXILIAR DE OPERACIONES": {
        "habil": ["03:00-11:00", "03:00-11:00", "11:00-19:00", "11:00-19:00", "19:00-27:00", "20:00-28:00"],
        "sabado": ["03:00-11:00", "03:00-11:00", "11:00-19:00", "11:00-19:00", "19:00-27:00", "22:00-30:00"],
        "domingo": ["03:00-11:00", "03:00-11:00", "11:00-19:00", "11:00-19:00", "19:00-27:00", "20:00-28:00"]
    },
    "OPERADOR LÍDER": {
        "habil": ["20:00-28:00", "20:00-28:00"],
        "sabado": ["20:00-28:00"],
        "domingo": ["20:00-28:00", "20:00-28:00"]
    },
    "TÉCNICO DE OPERACIONES": {
        "habil": ["03:00-11:00", "11:00-19:00", "19:00-27:00", "06:00-15:00"],
        "sabado": ["03:00-11:00", "11:00-19:00", "19:00-27:00", "06:00-15:00"],
        "domingo": ["03:00-11:00", "11:00-19:00", "19:00-27:00", "06:00-15:00"]
    },
    "AUXILIAR DE ALISTAMIENTO": {
        "habil": ["08:00-17:00", "22:00-30:00", "20:00-28:00", "20:00-28:00"],
        "sabado": ["08:00-17:00", "20:00-28:00", "22:00-30:00"],
        "domingo": ["08:00-17:00", "20:00-28:00", "22:00-30:00", "20:00-28:00"]
    }
}

PATRONES_ANALISTAS_5 = [
    ["06:00-15:00", "06:00-15:00", "06:00-15:00", "L", "03:00-11:00", "03:00-11:00", "08:00-17:00", "06:00-15:00", "06:00-15:00", "L", "03:00-11:00", "03:00-11:00", "03:00-11:00", "06:00-15:00"],
    ["08:00-17:00", "08:00-17:00", "08:00-17:00", "08:00-17:00", "08:00-17:00", "11:00-19:00", "L", "08:00-17:00", "08:00-17:00", "06:00-15:00", "06:00-15:00", "06:00-15:00", "L", "03:00-11:00"],
    ["11:00-19:00", "11:00-19:00", "11:00-19:00", "11:00-19:00", "11:00-19:00", "L", "03:00-11:00", "03:00-11:00", "03:00-11:00", "03:00-11:00", "03:00-11:00", "L", "19:00-27:00", "19:00-27:00"],
    ["03:00-11:00", "03:00-11:00", "03:00-11:00", "03:00-11:00", "L", "19:00-27:00", "19:00-27:00", "19:00-27:00", "19:00-27:00", "19:00-27:00", "19:00-27:00", "19:00-27:00", "L", "11:00-19:00"],
    ["19:00-27:00", "19:00-27:00", "19:00-27:00", "19:00-27:00", "19:00-27:00", "L", "11:00-19:00", "11:00-19:00", "11:00-19:00", "11:00-19:00", "11:00-19:00", "11:00-19:00", "11:00-19:00", "L"]
]

PATRONES_ANALISTAS_4 = [
    ["06:00-15:00", "06:00-15:00", "06:00-15:00", "L", "03:00-11:00", "03:00-11:00", "06:00-15:00", "06:00-15:00", "06:00-15:00", "06:00-15:00", "L", "03:00-11:00", "03:00-11:00", "03:00-11:00"],
    ["vacaciones", "vacaciones", "vacaciones", "vacaciones", "vacaciones", "vacaciones", "vacaciones", "vacaciones", "vacaciones", "vacaciones", "vacaciones", "vacaciones", "vacaciones", "vacaciones"],
    ["11:00-19:00", "11:00-19:00", "11:00-19:00", "11:00-19:00", "11:00-19:00", "L", "03:00-11:00", "03:00-11:00", "03:00-11:00", "03:00-11:00", "03:00-11:00", "L", "19:00-27:00", "19:00-27:00"],
    ["03:00-11:00", "03:00-11:00", "03:00-11:00", "03:00-11:00", "L", "19:00-27:00", "19:00-27:00", "19:00-27:00", "19:00-27:00", "19:00-27:00", "19:00-27:00", "19:00-27:00", "L", "11:00-19:00"],
    ["19:00-27:00", "19:00-27:00", "19:00-27:00", "19:00-27:00", "19:00-27:00", "L", "11:00-19:00", "11:00-19:00", "11:00-19:00", "11:00-19:00", "11:00-19:00", "11:00-19:00", "11:00-19:00", "L"]
]

DIAS_SEMANA_ES = ["LUNES", "MARTES", "MIÉRCOLES", "JUEVES", "VIERNES", "SÁBADO", "DOMINGO"]

# ==========================================
# CONFIGURACIÓN GENERAL EN SIDEBAR
# ==========================================
st.sidebar.header("⚙️ Parámetros de Programación")
semanas = st.sidebar.slider("Semanas a generar", 1, 4, 2)
fecha_inicio_date = st.sidebar.date_input("Fecha de inicio", datetime.now())

st.sidebar.markdown("---")
st.sidebar.header("🏖️ Configuración de Descansos")
libres_por_semana_base = st.sidebar.number_input("Días libres por semana (estricto)", min_value=1, max_value=3, value=1)
tiene_festivo = st.sidebar.checkbox("¿Hay día festivo en el periodo?", value=False)

fechas_festivas_sel = []
if tiene_festivo:
    dias_totales_temp = semanas * 7
    fechas_posibles = [(fecha_inicio_date + timedelta(days=i)) for i in range(dias_totales_temp)]
    fechas_festivas_sel = st.sidebar.multiselect(
        "Días festivos:", options=fechas_posibles, format_func=lambda x: x.strftime("%d-%b-%Y")
    )

# ==========================================
# FUNCIONES AUXILIARES DE TIEMPO Y FORMATO
# ==========================================
def obtener_iniciales_cargo(nombre_cargo: str) -> str:
    nombre = str(nombre_cargo).strip().upper()
    if "TÉCNICO" in nombre or "TECNICO" in nombre:
        return "TO"
    elif "ANALISTA" in nombre or "AUXILIAR DE OPERACIONES" in nombre:
        return "AO"
    elif "LÍDER" in nombre or "LIDER" in nombre:
        return "OL"
    elif "ALISTAMIENTO" in nombre:
        return "AA"
    else:
        palabras = [p for p in nombre.split() if p not in ["DE", "DEL", "LA", "EL"]]
        return "".join([p[0] for p in palabras])

def extraer_horas(texto_turno: str):
    coincidencia = re.search(r"(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})", str(texto_turno))
    if coincidencia:
        return map(int, coincidencia.groups())
    return None

def calcular_datetimes_turno(fecha_base_dt: datetime, texto_turno: str):
    """Calcula las fechas/horas exactas de inicio y fin manejando horas extendidas (>24)."""
    parsed = extraer_horas(texto_turno)
    if not parsed:
        return None, None
    h_i, m_i, h_f, m_f = parsed
    
    dt_inicio = fecha_base_dt + timedelta(hours=h_i, minutes=m_i)
    dt_fin = fecha_base_dt + timedelta(hours=h_f, minutes=m_f)
    return dt_inicio, dt_fin

def clasificar_franja(texto_turno: str) -> str:
    parsed = extraer_horas(texto_turno)
    if not parsed:
        return "MAÑANA"
    h_ini = list(parsed)[0]
    if 0 <= h_ini < 11:
        return "MAÑANA"
    elif 11 <= h_ini < 18:
        return "TARDE"
    return "NOCHE"

def obtener_siguiente_franja_permitida(franja_actual: str) -> list:
    mapeo = {
        "NOCHE": ["TARDE", "NOCHE"],
        "TARDE": ["MAÑANA", "TARDE"],
        "MAÑANA": ["NOCHE", "MAÑANA"]
    }
    return mapeo.get(franja_actual, ["NOCHE", "TARDE", "MAÑANA"])

def calcular_descanso_suficiente(salida_previa_dt: datetime, entrada_actual_dt: datetime, min_horas: float = 12.0) -> bool:
    if salida_previa_dt is None or entrada_actual_dt is None:
        return True
    diferencia_horas = (entrada_actual_dt - salida_previa_dt).total_seconds() / 3600.0
    return diferencia_horas >= min_horas

def parsear_fecha_incidencia(val_fecha, anio_referencia: int):
    if pd.isna(val_fecha) or str(val_fecha).strip() == "":
        return None
    try:
        dt = pd.to_datetime(val_fecha, dayfirst=True)
        if dt.year == 1970 or dt.year != anio_referencia:
            dt = dt.replace(year=anio_referencia)
        return dt.date()
    except Exception:
        return None

# ==========================================
# 1. CARGA INICIAL DE ARCHIVOS Y TAREAS
# ==========================================
st.subheader("1. Carga Inicial de Datos de Entrada")

col_f1, col_f2, col_f3 = st.columns(3)
df_empleados = None
df_semana_anterior = None
df_tareas_req = None

with col_f1:
    uploaded_file = st.file_uploader("1. Lista de Personal (Excel/CSV)", type=["xlsx", "csv"], key="file_personal")

with col_f2:
    uploaded_prev_file = st.file_uploader("2. Malla Semana Anterior (Opcional)", type=["xlsx", "csv"], key="file_prev")

with col_f3:
    uploaded_tareas_file = st.file_uploader("3. Matriz Tareas por Cargo (Opcional)", type=["xlsx", "csv"], key="file_tareas_ini")

if uploaded_tareas_file is not None:
    try:
        df_tareas_req = pd.read_csv(uploaded_tareas_file) if uploaded_tareas_file.name.endswith(".csv") else pd.read_excel(uploaded_tareas_file)
        df_tareas_req.columns = [str(c).strip().upper() for c in df_tareas_req.columns]
        if "CARGO" in df_tareas_req.columns:
            df_tareas_req["CARGO"] = df_tareas_req["CARGO"].astype(str).str.strip().str.upper()
        st.success("✅ Matriz de Tareas cargada.")
    except Exception as e:
        st.error(f"Error en archivo de tareas: {e}")
else:
    filas_default = []
    for c, t_dict in TURNOS_DEFAULT_POR_CARGO.items():
        max_len = max(len(t_dict["habil"]), len(t_dict["sabado"]), len(t_dict["domingo"]))
        for i in range(max_len):
            filas_default.append({
                "CARGO": c.upper(),
                "HABIL": t_dict["habil"][i] if i < len(t_dict["habil"]) else "",
                "SABADO": t_dict["sabado"][i] if i < len(t_dict["sabado"]) else "",
                "DOMINGO": t_dict["domingo"][i] if i < len(t_dict["domingo"]) else ""
            })
    df_tareas_req = pd.DataFrame(filas_default)

with st.expander("👁️ Ver / Editar Tareas Requeridas por Cargo (Habil, Sábado, Domingo)", expanded=False):
    st.dataframe(df_tareas_req, use_container_width=True)

if uploaded_file is not None:
    try:
        df_empleados = pd.read_csv(uploaded_file) if uploaded_file.name.endswith(".csv") else pd.read_excel(uploaded_file)
        df_empleados.columns = [str(c).upper().strip() for c in df_empleados.columns]

        if not {"CODIGO", "NOMBRE", "CARGO"}.issubset(set(df_empleados.columns)):
            st.error("El archivo de personal debe contener las columnas: CODIGO, NOMBRE, CARGO")
            df_empleados = None
        else:
            df_empleados["CARGO"] = df_empleados["CARGO"].astype(str).str.strip().str.upper()
            if "ESTADO" not in df_empleados.columns:
                df_empleados["ESTADO"] = "ACTIVO"
            col_inc = [c for c in df_empleados.columns if "INCIDENCIA" in c]
            col_f_ini = [c for c in df_empleados.columns if "FECHA INICI" in c or "FECHA_INICI" in c]
            col_f_fin = [c for c in df_empleados.columns if "FECHA FIN" in c or "FECHA_FIN" in c]

            df_empleados["INCIDENCIA_TIPO"] = df_empleados[col_inc[0]] if col_inc else None
            df_empleados["INCIDENCIA_INI"] = df_empleados[col_f_ini[0]] if col_f_ini else None
            df_empleados["INCIDENCIA_FIN"] = df_empleados[col_f_fin[0]] if col_f_fin else None
            st.success(f"¡Se cargaron {len(df_empleados)} empleados exitosamente!")
    except Exception as e:
        st.error(f"Error al procesar personal: {e}")

if uploaded_prev_file is not None:
    try:
        df_semana_anterior = pd.read_csv(uploaded_prev_file) if uploaded_prev_file.name.endswith(".csv") else pd.read_excel(uploaded_prev_file)
        df_semana_anterior.columns = [str(c).upper().strip() for c in df_semana_anterior.columns]
        if "CARGO" in df_semana_anterior.columns:
            df_semana_anterior["CARGO"] = df_semana_anterior["CARGO"].astype(str).str.strip().str.upper()
        st.success("✅ Malla anterior cargada correctamente para empalmar rotación.")
    except Exception as e:
        st.warning(f"No se pudo leer la semana anterior: {e}")

# ==========================================
# SECCIÓN MULTI-SELECCIÓN DE REEMPLAZOS
# ==========================================
reemplazos_config = {}

if df_empleados is not None:
    st.sidebar.markdown("---")
    st.sidebar.header("🔄 Cobertura Inter-Cargo / Reemplazos")
    activa_reemplazo = st.sidebar.checkbox("¿Asignar colaboradores a otro cargo?", value=False)
    
    if activa_reemplazo:
        if "lista_reemplazos" not in st.session_state:
            st.session_state.lista_reemplazos = []

        lista_empleados_nombres = df_empleados["NOMBRE"].tolist()
        cargos_disponibles = sorted(list(set(df_empleados["CARGO"].tolist() + list(TURNOS_DEFAULT_POR_CARGO.keys()))))

        emps_sel = st.sidebar.multiselect("1. Seleccionar Colaborador(es):", lista_empleados_nombres)
        cargo_destino_sel = st.sidebar.selectbox("2. Cargo secundario a cubrir:", cargos_disponibles)

        if st.sidebar.button("➕ Agregar Polivalencia"):
            for emp_nombre in emps_sel:
                row_emp = df_empleados[df_empleados["NOMBRE"] == emp_nombre].iloc[0]
                cod = str(row_emp["CODIGO"]).strip()
                c_orig = row_emp["CARGO"]

                if not any(item["CODIGO"] == cod for item in st.session_state.lista_reemplazos):
                    st.session_state.lista_reemplazos.append({
                        "CODIGO": cod,
                        "NOMBRE": emp_nombre,
                        "CARGO ORIGEN": c_orig,
                        "CARGO A CUBRIR": cargo_destino_sel
                    })

        if st.session_state.lista_reemplazos:
            st.sidebar.subheader("📋 Reemplazos Programados")
            df_temp_reemplazos = pd.DataFrame(st.session_state.lista_reemplazos)
            st.sidebar.dataframe(df_temp_reemplazos[["NOMBRE", "CARGO A CUBRIR"]], use_container_width=True)

            if st.sidebar.button("🗑️ Limpiar Reemplazos"):
                st.session_state.lista_reemplazos = []

            for item in st.session_state.lista_reemplazos:
                reemplazos_config[item["CODIGO"]] = item["CARGO A CUBRIR"]

# ==========================================
# 2. CONSTRUCCIÓN DE REGLAS DE DEMANDA
# ==========================================
matriz_demanda = {}

if df_empleados is not None:
    cargos_unicos = df_empleados["CARGO"].dropna().unique().tolist()
    for cargo in cargos_unicos:
        cargo_clean = str(cargo).strip().upper()
        matriz_demanda[cargo_clean] = {d: [] for d in DIAS_SEMANA_ES}
        sub_mat = df_tareas_req[df_tareas_req["CARGO"] == cargo_clean] if "CARGO" in df_tareas_req.columns else pd.DataFrame()
        
        col_habil = next((c for c in sub_mat.columns if "HABIL" in c), None)
        col_sab = next((c for c in sub_mat.columns if "SAB" in c), None)
        col_dom = next((c for c in sub_mat.columns if "DOM" in c), None)

        req_habil = [str(x).strip() for x in sub_mat[col_habil].dropna().tolist() if str(x).strip() != ""] if col_habil else []
        req_sab = [str(x).strip() for x in sub_mat[col_sab].dropna().tolist() if str(x).strip() != ""] if col_sab else []
        req_dom = [str(x).strip() for x in sub_mat[col_dom].dropna().tolist() if str(x).strip() != ""] if col_dom else []

        if not req_habil and not req_sab and not req_dom:
            defaults = TURNOS_DEFAULT_POR_CARGO.get(cargo_clean, {})
            req_habil = defaults.get("habil", [])
            req_sab = defaults.get("sabado", [])
            req_dom = defaults.get("domingo", [])

        for dh in ["LUNES", "MARTES", "MIÉRCOLES", "JUEVES", "VIERNES"]:
            matriz_demanda[cargo_clean][dh] = list(req_habil)
        matriz_demanda[cargo_clean]["SÁBADO"] = list(req_sab)
        matriz_demanda[cargo_clean]["DOMINGO"] = list(req_dom)

# ==========================================
# 3. GENERACIÓN DE MALLA OPTIMIZADA
# ==========================================
def generar_malla_matriz(df_personal, semanas_count, fecha_base_date, reglas_demanda, libres_base, festivos_list, df_prev=None, mapa_reemplazos=None):
    if mapa_reemplazos is None:
        mapa_reemplazos = {}

    dias_totales = semanas_count * 7
    fecha_base = datetime.combine(fecha_base_date, time.min)
    anio_ref = fecha_base_date.year
    columnas_fechas, fechas_dt = [], []
    
    for i in range(dias_totales):
        f_actual = fecha_base + timedelta(days=i)
        fechas_dt.append(f_actual)
        columnas_fechas.append(f"{DIAS_SEMANA_ES[f_actual.weekday()]}\n{f_actual.strftime('%d-%b')}")

    info_historial = {}
    if df_prev is not None and "CODIGO" in df_prev.columns:
        cols_dias_prev = [c for c in df_prev.columns if c not in ["CODIGO", "NOMBRE", "CARGO", "ESTADO"]]
        if cols_dias_prev:
            for _, fila in df_prev.iterrows():
                c_cod = str(fila["CODIGO"]).strip()
                val_ult = str(fila[cols_dias_prev[-1]]).strip()
                val_ult_limpio = val_ult.split()[0] if " " in val_ult else val_ult
                es_descanso_o_inc = val_ult_limpio.upper() in ["L", "AO", "VACACIONES", "LICENCIA", "INCAPACIDAD", "PERMISO"]
                
                info_historial[c_cod] = {
                    "ultimo_turno": None if es_descanso_o_inc else val_ult_limpio,
                    "ultima_franja": None if es_descanso_o_inc else clasificar_franja(val_ult_limpio),
                    "termino_en_descanso": es_descanso_o_inc
                }

    analistas = [emp for _, emp in df_personal.iterrows() if "ANALISTA" in str(emp["CARGO"]).upper()]
    total_analistas = len(analistas)
    patrones_analistas = PATRONES_ANALISTAS_5 if total_analistas >= 5 else PATRONES_ANALISTAS_4

    programacion_matriz, dias_libres_emp = {}, {}
    idx_patron_analistas_counter = 0

    for _, emp in df_personal.iterrows():
        cod = str(emp["CODIGO"]).strip()
        cargo_original = str(emp["CARGO"]).strip().upper()
        cargo_secundario = mapa_reemplazos.get(cod, None)
        hist = info_historial.get(cod, {})

        programacion_matriz[cod] = {
            "CODIGO": cod, "NOMBRE": emp["NOMBRE"],
            "CARGO_ORIGINAL": cargo_original,
            "CARGO_SECUNDARIO": cargo_secundario,
            "CARGO": cargo_original,
            "INCIDENCIA_TIPO": emp.get("INCIDENCIA_TIPO"),
            "INCIDENCIA_INI": parsear_fecha_incidencia(emp.get("INCIDENCIA_INI"), anio_ref),
            "INCIDENCIA_FIN": parsear_fecha_incidencia(emp.get("INCIDENCIA_FIN"), anio_ref),
            "TURNO_FIJO_BLOQUE": hist.get("ultimo_turno") if not hist.get("termino_en_descanso", True) else None,
            "ULTIMA_FRANJA": hist.get("ultima_franja"),
            "SALIDA_PREVIA_DT": None,
            "HISTORIAL_CARGOS_DIARIOS": {},
            "HISTORIAL_TURNOS_LIMPIOS": {}
        }

        if "ANALISTA" in cargo_original:
            p_idx_base = idx_patron_analistas_counter % len(patrones_analistas)
            idx_patron_analistas_counter += 1
            programacion_matriz[cod]["PATRON_BASE"] = p_idx_base

    lideres = [cod for cod, d in programacion_matriz.items() if "LÍDER" in d["CARGO_ORIGINAL"] or "LIDER" in d["CARGO_ORIGINAL"]]
    tecnicos = [cod for cod, d in programacion_matriz.items() if "TÉCNICO" in d["CARGO_ORIGINAL"] or "TECNICO" in d["CARGO_ORIGINAL"]]

    conteo_libres_tecnicos = {i: 0 for i in range(dias_totales)}

    # ASIGNACIÓN DE DÍAS LIBRES
    for s in range(semanas_count):
        indices_semana = [s * 7 + i for i in range(7)]
        carga_diaria = {idx: 0 for idx in indices_semana}
        for idx in indices_semana:
            fecha_temp = fechas_dt[idx]
            dia_nom = DIAS_SEMANA_ES[fecha_temp.weekday()]
            for c_k, v_dict in reglas_demanda.items():
                if "ANALISTA" not in c_k:
                    carga_diaria[idx] += len(v_dict.get(dia_nom, []))

        for _, emp in df_personal.iterrows():
            cod = str(emp["CODIGO"]).strip()
            cargo_emp = programacion_matriz[cod]["CARGO_ORIGINAL"]

            if "ANALISTA" not in cargo_emp:
                if cod not in dias_libres_emp:
                    dias_libres_emp[cod] = set()

                inicio_sem = fecha_base + timedelta(days=s * 7)
                fin_sem = inicio_sem + timedelta(days=6)
                hay_festivo_sem = any(inicio_sem.date() <= f <= fin_sem.date() for f in festivos_list)
                cant_libres = libres_base + (1 if hay_festivo_sem else 0)

                if cod in lideres:
                    idx_lider = lideres.index(cod)
                    dias_libres_emp[cod].add(s * 7 + (4 if idx_lider % 2 == 0 else 5))

                elif cod in tecnicos:
                    candidatos = [idx_d for idx_d in indices_semana if conteo_libres_tecnicos[idx_d] < 1]
                    elegidos = candidatos[:cant_libres] if len(candidatos) >= cant_libres else sorted(indices_semana, key=lambda x: conteo_libres_tecnicos[x])[:cant_libres]
                    for el in elegidos:
                        dias_libres_emp[cod].add(el)
                        conteo_libres_tecnicos[el] += 1
                else:
                    dias_ordenados_por_carga = sorted(indices_semana, key=lambda idx: carga_diaria[idx])
                    libres_elegidos = dias_ordenados_por_carga[:cant_libres]
                    dias_libres_emp[cod].update(libres_elegidos)
                    for el in libres_elegidos:
                        carga_diaria[el] += 2

    # PROCESAMIENTO DÍA A DÍA
    for idx_dia, col_nombre in enumerate(columnas_fechas):
        fecha_col = fechas_dt[idx_dia]
        fecha_actual_date = fecha_col.date()
        nombre_dia_semana = DIAS_SEMANA_ES[fecha_col.weekday()]
        dia_matriz_14 = idx_dia % 14
        dia_semana_idx = idx_dia % 7

        demandas_dia_actual = {c_k: list(v_dict.get(nombre_dia_semana, [])) for c_k, v_dict in reglas_demanda.items()}
        turnos_usados_analistas_hoy = set()
        libres_analistas_hoy = 0

        empleados_ordenados = sorted(
            programacion_matriz.keys(),
            key=lambda c: 0 if ("TÉCNICO" in programacion_matriz[c]["CARGO_ORIGINAL"] or "TECNICO" in programacion_matriz[c]["CARGO_ORIGINAL"]) else 1
        )

        # PASADA 1: Asignación Preferente
        for cod_emp in empleados_ordenados:
            d_emp = programacion_matriz[cod_emp]
            cargo_orig = d_emp["CARGO_ORIGINAL"]
            cargo_sec = d_emp["CARGO_SECUNDARIO"]

            # Gestión de Incidencias
            if d_emp["INCIDENCIA_TIPO"] and d_emp["INCIDENCIA_INI"] and d_emp["INCIDENCIA_FIN"]:
                if d_emp["INCIDENCIA_INI"] <= fecha_actual_date <= d_emp["INCIDENCIA_FIN"]:
                    programacion_matriz[cod_emp][col_nombre] = d_emp["INCIDENCIA_TIPO"]
                    d_emp["TURNO_FIJO_BLOQUE"] = None
                    d_emp["SALIDA_PREVIA_DT"] = None
                    d_emp["HISTORIAL_CARGOS_DIARIOS"][col_nombre] = cargo_orig
                    d_emp["HISTORIAL_TURNOS_LIMPIOS"][col_nombre] = d_emp["INCIDENCIA_TIPO"]
                    continue

            # Gestión de Analistas (Matriz Rotativa Fija)
            if "ANALISTA" in cargo_orig:
                p_base = d_emp["PATRON_BASE"]
                turno_sugerido = patrones_analistas[p_base][dia_matriz_14]

                if turno_sugerido == "L" and dia_semana_idx <= 4:
                    if libres_analistas_hoy >= 1:
                        candidatos = [t for t in patrones_analistas[p_base] if t not in ["L", "vacaciones"] and t not in turnos_usados_analistas_hoy]
                        turno_sugerido = candidatos[0] if candidatos else "AO"
                    else:
                        libres_analistas_hoy += 1

                if turno_sugerido not in ["L", "vacaciones"]:
                    if turno_sugerido in turnos_usados_analistas_hoy:
                        alt_turnos = [t for t in demandas_dia_actual.get(cargo_orig, []) if t not in turnos_usados_analistas_hoy]
                        turno_sugerido = alt_turnos[0] if alt_turnos else "AO"
                    if turno_sugerido != "AO":
                        turnos_usados_analistas_hoy.add(turno_sugerido)

                programacion_matriz[cod_emp][col_nombre] = turno_sugerido
                d_emp["HISTORIAL_CARGOS_DIARIOS"][col_nombre] = cargo_orig
                d_emp["HISTORIAL_TURNOS_LIMPIOS"][col_nombre] = turno_sugerido

            # Gestión de Otros Cargos Operativos
            else:
                if idx_dia in dias_libres_emp.get(cod_emp, set()):
                    programacion_matriz[cod_emp][col_nombre] = "L"
                    d_emp["TURNO_FIJO_BLOQUE"] = None
                    d_emp["SALIDA_PREVIA_DT"] = None
                    d_emp["ULTIMA_FRANJA"] = None
                    d_emp["HISTORIAL_CARGOS_DIARIOS"][col_nombre] = cargo_orig
                    d_emp["HISTORIAL_TURNOS_LIMPIOS"][col_nombre] = "L"
                    continue

                turno_a_asignar = None
                cargo_cubierto_hoy = cargo_orig
                cargos_a_evaluar = [cargo_orig] + ([cargo_sec] if cargo_sec else [])

                for c_eval in cargos_a_evaluar:
                    turnos_disp_cargo = demandas_dia_actual.get(c_eval, [])
                    if not turnos_disp_cargo:
                        continue

                    es_tecnico_planta = ("TÉCNICO" in cargo_orig or "TECNICO" in cargo_orig)
                    es_cubriendo_tecnico = ("TÉCNICO" in c_eval or "TECNICO" in c_eval)

                    cand_list = list(turnos_disp_cargo)
                    if es_cubriendo_tecnico and dia_semana_idx <= 5:
                        if es_tecnico_planta:
                            cand_list.sort(key=lambda t: (0 if any(str(t).startswith(p) for p in ["03:00", "11:00", "19:00"]) else 1))
                        else:
                            cand_list.sort(key=lambda t: (0 if turnos_disp_cargo.count(t) > 1 else 1))

                    # 1. Verificar Mantenimiento de Bloque
                    cand_bloque = d_emp.get("TURNO_FIJO_BLOQUE")
                    if cand_bloque and cand_bloque in cand_list:
                        dt_ent, _ = calcular_datetimes_turno(fecha_col, cand_bloque)
                        if calcular_descanso_suficiente(d_emp["SALIDA_PREVIA_DT"], dt_ent, min_horas=12):
                            turno_a_asignar = cand_bloque
                            turnos_disp_cargo.remove(cand_bloque)
                            cargo_cubierto_hoy = c_eval
                            break

                    # 2. Selección del Mejor Turno Disponible
                    franja_ult = d_emp["ULTIMA_FRANJA"]
                    franjas_permitidas = obtener_siguiente_franja_permitida(franja_ult) if franja_ult else ["NOCHE", "TARDE", "MAÑANA"]
                    
                    if not (es_cubriendo_tecnico and dia_semana_idx <= 5):
                        cand_list.sort(key=lambda t: 0 if clasificar_franja(t) in franjas_permitidas else 1)

                    for cand_t in cand_list:
                        dt_ent, _ = calcular_datetimes_turno(fecha_col, cand_t)
                        if calcular_descanso_suficiente(d_emp["SALIDA_PREVIA_DT"], dt_ent, min_horas=12):
                            turno_a_asignar = cand_t
                            turnos_disp_cargo.remove(cand_t)
                            cargo_cubierto_hoy = c_eval
                            break
                    
                    if turno_a_asignar:
                        break

                # Asignación final y actualización de estado
                if turno_a_asignar is None:
                    programacion_matriz[cod_emp][col_nombre] = "AO"
                    d_emp["TURNO_FIJO_BLOQUE"] = None
                    d_emp["SALIDA_PREVIA_DT"] = None
                    d_emp["HISTORIAL_CARGOS_DIARIOS"][col_nombre] = cargo_orig
                    d_emp["HISTORIAL_TURNOS_LIMPIOS"][col_nombre] = "AO"
                else:
                    d_emp["TURNO_FIJO_BLOQUE"] = turno_a_asignar
                    d_emp["ULTIMA_FRANJA"] = clasificar_franja(turno_a_asignar)
                    d_emp["HISTORIAL_CARGOS_DIARIOS"][col_nombre] = cargo_cubierto_hoy
                    d_emp["HISTORIAL_TURNOS_LIMPIOS"][col_nombre] = turno_a_asignar

                    programacion_matriz[cod_emp][col_nombre] = (
                        f"{turno_a_asignar} {obtener_iniciales_cargo(cargo_cubierto_hoy)}" 
                        if cargo_cubierto_hoy != cargo_orig else turno_a_asignar
                    )

                    _, dt_salida = calcular_datetimes_turno(fecha_col, turno_a_asignar)
                    d_emp["SALIDA_PREVIA_DT"] = dt_salida

        # PASADA 2: Recuperación de Puestos "AO"
        for cod_emp in empleados_ordenados:
            d_emp = programacion_matriz[cod_emp]
            cargo_orig = d_emp["CARGO_ORIGINAL"]
            cargo_sec = d_emp["CARGO_SECUNDARIO"]

            if "ANALISTA" in cargo_orig or programacion_matriz[cod_emp][col_nombre] != "AO":
                continue

            cargos_a_evaluar = [cargo_orig] + ([cargo_sec] if cargo_sec else [])
            for c_eval in cargos_a_evaluar:
                turnos_disp_cargo = demandas_dia_actual.get(c_eval, [])
                if not turnos_disp_cargo:
                    continue

                asignado_f2 = False
                for cand_t in list(turnos_disp_cargo):
                    dt_ent, dt_salida = calcular_datetimes_turno(fecha_col, cand_t)
                    if calcular_descanso_suficiente(d_emp["SALIDA_PREVIA_DT"], dt_ent, min_horas=12):
                        d_emp["TURNO_FIJO_BLOQUE"] = cand_t
                        d_emp["ULTIMA_FRANJA"] = clasificar_franja(cand_t)
                        d_emp["HISTORIAL_CARGOS_DIARIOS"][col_nombre] = c_eval
                        d_emp["HISTORIAL_TURNOS_LIMPIOS"][col_nombre] = cand_t
                        d_emp["SALIDA_PREVIA_DT"] = dt_salida
                        turnos_disp_cargo.remove(cand_t)

                        programacion_matriz[cod_emp][col_nombre] = (
                            f"{cand_t} {obtener_iniciales_cargo(c_eval)}" 
                            if c_eval != cargo_orig else cand_t
                        )
                        asignado_f2 = True
                        break

                if asignado_f2:
                    break

    # Construcción de DataFrame Limpio
    columnas_excluir = {
        "INCIDENCIA_TIPO", "INCIDENCIA_INI", "INCIDENCIA_FIN", "PATRON_BASE", 
        "TURNO_FIJO_BLOQUE", "ULTIMA_FRANJA", "CARGO_ORIGINAL", "CARGO_SECUNDARIO", 
        "SALIDA_PREVIA_DT", "HISTORIAL_CARGOS_DIARIOS", "HISTORIAL_TURNOS_LIMPIOS"
    }
    df_resultado = pd.DataFrame([
        {k: v for k, v in datos.items() if k not in columnas_excluir} 
        for datos in programacion_matriz.values()
    ])

    return df_resultado, programacion_matriz

# ==========================================
# 4. EJECUCIÓN Y VISUALIZACIÓN
# ==========================================
if df_empleados is not None:
    if st.button("⚡ Generar Horarios y Auditar Tareas"):
        df_res, dict_matriz = generar_malla_matriz(
            df_empleados, semanas, fecha_inicio_date, matriz_demanda, libres_por_semana_base, fechas_festivas_sel, df_semana_anterior, reemplazos_config
        )
        st.session_state.df_resultado = df_res
        st.session_state.dict_matriz = dict_matriz

if "df_resultado" in st.session_state:
    df_resultado = st.session_state.df_resultado
    dict_matriz = st.session_state.dict_matriz

    st.subheader("2. Malla Horaria Generada")
    st.dataframe(df_resultado, use_container_width=True)

    # REPORTE DE COBERTURA INTER-CARGO
    st.markdown("---")
    st.subheader("🔄 Reporte de Cobertura de Turnos de Diferente Cargo")

    reporte_coberturas = []
    cols_fechas_malla = [c for c in df_resultado.columns if c not in ["CODIGO", "NOMBRE", "CARGO"]]

    for cod_emp, d_emp in dict_matriz.items():
        cargo_orig = d_emp["CARGO_ORIGINAL"]
        cargo_sec = d_emp["CARGO_SECUNDARIO"]
        historial_cargos = d_emp.get("HISTORIAL_CARGOS_DIARIOS", {})
        historial_turnos_limpios = d_emp.get("HISTORIAL_TURNOS_LIMPIOS", {})

        if cargo_sec:
            for s in range(semanas):
                cols_semana = cols_fechas_malla[s * 7 : (s + 1) * 7]
                dias_cubiertos = []
                for col in cols_semana:
                    if historial_cargos.get(col) == cargo_sec:
                        val_t = historial_turnos_limpios.get(col, "")
                        if val_t not in ["L", "AO", "VACACIONES", "LICENCIA", "INCAPACIDAD", "PERMISO"]:
                            dias_cubiertos.append(f"{col.split()[0]}: {val_t}")

                reporte_coberturas.append({
                    "SEMANA": f"Semana {s + 1}",
                    "CÓDIGO": cod_emp,
                    "NOMBRE": d_emp["NOMBRE"],
                    "CARGO PERTENECIENTE": cargo_orig,
                    "CARGO CUBIERTO": cargo_sec,
                    "DETALLE DE TURNOS APORTADOS": ", ".join(dias_cubiertos) if dias_cubiertos else "Trabajó en su cargo original / Libre"
                })

    df_coberturas = pd.DataFrame(reporte_coberturas)
    if not df_coberturas.empty:
        st.info("ℹ️ A continuación se detalla el personal que realizó coberturas en un cargo secundario:")
        st.dataframe(df_coberturas, use_container_width=True)

    # RESUMEN DE FALTANTES
    st.markdown("---")
    st.subheader("🚨 Resumen Semanal de Tareas Desatendidas / Faltantes")

    for s in range(semanas):
        cols_semana = cols_fechas_malla[s * 7 : (s + 1) * 7]
        cols_semana_limpias = [c.replace("\n", " ") for c in cols_semana]
        
        datos_resumen_semana = []
        cargos_evaluar = df_empleados["CARGO"].unique().tolist()

        for cargo in cargos_evaluar:
            cargo_clean = str(cargo).strip().upper()
            sub_mat = df_tareas_req[df_tareas_req["CARGO"] == cargo_clean] if "CARGO" in df_tareas_req.columns else pd.DataFrame()

            for col_f in cols_semana:
                dia_nombre_ext = col_f.split("\n")[0].upper()
                tipo_col_mat = "SABADO" if dia_nombre_ext in ["SÁBADO", "SABADO"] else ("DOMINGO" if dia_nombre_ext == "DOMINGO" else "HABIL")
                col_target_mat = next((c for c in sub_mat.columns if tipo_col_mat in c), None)

                turnos_req = [str(x).strip() for x in sub_mat[col_target_mat].dropna().tolist() if str(x).strip() != ""] if col_target_mat else []

                turnos_cubiertos = [
                    d_e.get("HISTORIAL_TURNOS_LIMPIOS", {}).get(col_f, "")
                    for cod_e, d_e in dict_matriz.items()
                    if d_e.get("HISTORIAL_CARGOS_DIARIOS", {}).get(col_f) == cargo_clean
                    and d_e.get("HISTORIAL_TURNOS_LIMPIOS", {}).get(col_f, "") not in ["L", "AO", "VACACIONES", "LICENCIA", "INCAPACIDAD", "PERMISO"]
                ]

                for tr in set(turnos_req):
                    cant_req_t = turnos_req.count(tr)
                    cant_cub_t = turnos_cubiertos.count(tr)
                    faltante_t = max(0, cant_req_t - cant_cub_t)

                    datos_resumen_semana.append({
                        "CARGO": cargo_clean,
                        "TURNO": tr,
                        "DÍA": col_f.replace("\n", " "),
                        "FALTAN POR ASIGNAR": faltante_t
                    })

        df_sem_res = pd.DataFrame(datos_resumen_semana)

        if not df_sem_res.empty:
            tabla_pivot = df_sem_res.pivot_table(
                index=["CARGO", "TURNO"],
                columns="DÍA",
                values="FALTAN POR ASIGNAR",
                aggfunc="sum",
                fill_value=0
            )

            cols_existentes_ordenadas = [c for c in cols_semana_limpias if c in tabla_pivot.columns]
            tabla_pivot = tabla_pivot.reindex(columns=cols_existentes_ordenadas)

            def resaltar_faltantes_rojo(val):
                if isinstance(val, (int, float)) and val > 0:
                    return "background-color: #ff4b4b; color: white; font-weight: bold;"
                return "background-color: #e6ffed; color: #0d5a22;"

            st.markdown(f"##### 📌 Semana {s + 1}")
            st.dataframe(tabla_pivot.style.map(resaltar_faltantes_rojo), use_container_width=True)

    # EXPORTACIÓN
    st.markdown("---")
    st.subheader("📥 Exportación de Reportes Operativos")
    
    col_exp1, col_exp2 = st.columns(2)
    buffer_excel = io.BytesIO()

    with pd.ExcelWriter(buffer_excel) as writer:
        df_resultado.to_excel(writer, sheet_name="Malla Horaria", index=False)
        if not df_coberturas.empty:
            df_coberturas.to_excel(writer, sheet_name="Reemplazos Inter-Cargo", index=False)

    buffer_excel.seek(0)

    with col_exp1:
        st.download_button(
            label="📥 Descargar Malla Horaria Completa (Excel)",
            data=buffer_excel,
            file_name=f"Malla_Horaria_{fecha_inicio_date.strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with col_exp2:
        st.download_button(
            label="📄 Descargar Malla Simplificada (CSV)",
            data=df_resultado.to_csv(index=False).encode('utf-8'),
            file_name=f"Malla_Horaria_{fecha_inicio_date.strftime('%Y%m%d')}.csv",
            mime="text/csv"
        )
