from datetime import datetime, timedelta
import io
import re
import random
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

dias_semana_es = ["LUNES", "MARTES", "MIÉRCOLES", "JUEVES", "VIERNES", "SÁBADO", "DOMINGO"]

# ==========================================
# FUNCIONES AUXILIARES DE TIEMPO Y ROTACIÓN
# ==========================================
def extraer_horas(texto_turno):
    coincidencia = re.search(r"(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})", str(texto_turno))
    if coincidencia:
        h_ini, m_ini, h_fin, m_fin = map(int, coincidencia.groups())
        return h_ini, m_ini, h_fin, m_fin
    return None

def clasificar_franja(texto_turno):
    parsed = extraer_horas(texto_turno)
    if not parsed:
        return "MAÑANA"
    h_ini = parsed[0]
    if 0 <= h_ini < 11:
        return "MAÑANA"
    elif 11 <= h_ini < 18:
        return "TARDE"
    else:
        return "NOCHE"

def obtener_siguiente_franja_permitida(franja_actual):
    if franja_actual == "NOCHE":
        return ["TARDE", "NOCHE"]
    elif franja_actual == "TARDE":
        return ["MAÑANA", "TARDE"]
    elif franja_actual == "MAÑANA":
        return ["NOCHE", "MAÑANA"]
    return ["NOCHE", "TARDE", "MAÑANA"]

def parsear_fecha_incidencia(val_fecha, anio_referencia):
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
# CONFIGURACIÓN DE COBERTURAS INTER-CARGO
# ==========================================
reemplazos_config = {}
if df_empleados is not None:
    st.sidebar.markdown("---")
    st.sidebar.header("🔄 Cobertura Inter-Cargo / Reemplazos")
    activa_reemplazo = st.sidebar.checkbox("¿Asignar un colaborador a un cargo/turno diferente?", value=False)
    
    if activa_reemplazo:
        lista_empleados_nombres = df_empleados["NOMBRE"].tolist()
        cargos_disponibles = list(set(df_empleados["CARGO"].tolist() + list(TURNOS_DEFAULT_POR_CARGO.keys())))
        
        emp_sel = st.sidebar.selectbox("Seleccionar Colaborador:", lista_empleados_nombres)
        cargo_destino_sel = st.sidebar.selectbox("Cargo a cubrir:", cargos_disponibles)
        
        if emp_sel and cargo_destino_sel:
            row_emp = df_empleados[df_empleados["NOMBRE"] == emp_sel].iloc[0]
            cod_sel = str(row_emp["CODIGO"]).strip()
            reemplazos_config[cod_sel] = cargo_destino_sel
            st.sidebar.info(f"💡 {emp_sel} ({row_emp['CARGO']}) cubrirá tareas de {cargo_destino_sel}")

# ==========================================
# 2. CONSTRUCCIÓN DE REGLAS DE DEMANDA
# ==========================================
matriz_demanda = {}

if df_empleados is not None:
    cargos_unicos = df_empleados["CARGO"].dropna().unique().tolist()
    for cargo in cargos_unicos:
        cargo_clean = str(cargo).strip().upper()
        matriz_demanda[cargo_clean] = {d: [] for d in dias_semana_es}
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
# 3. GENERACIÓN DE MALLA INTEGRADA
# ==========================================
def generar_malla_matriz(df_personal, semanas_count, fecha_base_date, reglas_demanda, libres_base, festivos_list, df_prev=None, mapa_reemplazos=None):
    if mapa_reemplazos is None:
        mapa_reemplazos = {}

    dias_totales = semanas_count * 7
    fecha_base = datetime.combine(fecha_base_date, datetime.min.time())
    anio_ref = fecha_base_date.year
    columnas_fechas, fechas_dt = [], []
    
    for i in range(dias_totales):
        f_actual = fecha_base + timedelta(days=i)
        fechas_dt.append(f_actual)
        columnas_fechas.append(f"{dias_semana_es[f_actual.weekday()]}\n{f_actual.strftime('%d-%b')}")

    info_historial = {}
    if df_prev is not None and "CODIGO" in df_prev.columns:
        cols_dias_prev = [c for c in df_prev.columns if c not in ["CODIGO", "NOMBRE", "CARGO", "ESTADO"]]
        if cols_dias_prev:
            for _, fila in df_prev.iterrows():
                c_cod = str(fila["CODIGO"]).strip()
                val_ult = str(fila[cols_dias_prev[-1]]).strip()
                es_descanso_o_inc = val_ult.upper() in ["L", "AO", "VACACIONES", "LICENCIA", "INCAPACIDAD", "PERMISO"]
                
                info_historial[c_cod] = {
                    "ultimo_turno": None if es_descanso_o_inc else val_ult,
                    "ultima_franja": None if es_descanso_o_inc else clasificar_franja(val_ult),
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
        
        cargo_operativo = mapa_reemplazos.get(cod, cargo_original)
        hist = info_historial.get(cod, {})

        programacion_matriz[cod] = {
            "CODIGO": cod, "NOMBRE": emp["NOMBRE"],
            "CARGO_ORIGINAL": cargo_original,
            "CARGO": cargo_operativo,
            "INCIDENCIA_TIPO": emp.get("INCIDENCIA_TIPO"),
            "INCIDENCIA_INI": parsear_fecha_incidencia(emp.get("INCIDENCIA_INI"), anio_ref),
            "INCIDENCIA_FIN": parsear_fecha_incidencia(emp.get("INCIDENCIA_FIN"), anio_ref),
            "TURNO_FIJO_BLOQUE": hist.get("ultimo_turno") if not hist.get("termino_en_descanso", True) else None,
            "ULTIMA_FRANJA": hist.get("ultima_franja")
        }

        if "ANALISTA" in cargo_operativo:
            p_idx_base = idx_patron_analistas_counter % len(patrones_analistas)
            idx_patron_analistas_counter += 1
            programacion_matriz[cod]["PATRON_BASE"] = p_idx_base

    lideres = [cod for cod, d in programacion_matriz.items() if "LÍDER" in d["CARGO"] or "LIDER" in d["CARGO"]]
    
    for _, emp in df_personal.iterrows():
        cod = str(emp["CODIGO"]).strip()
        cargo_emp = programacion_matriz[cod]["CARGO"]

        if "ANALISTA" not in cargo_emp:
            libres_indices_emp = set()
            for s in range(semanas_count):
                inicio_sem = fecha_base + timedelta(days=s * 7)
                fin_sem = inicio_sem + timedelta(days=6)
                hay_festivo_sem = any(inicio_sem.date() <= f <= fin_sem.date() for f in festivos_list)
                cant_libres = libres_base + (1 if hay_festivo_sem else 0)

                if cod in lideres:
                    idx_lider = lideres.index(cod)
                    if idx_lider % 2 == 0:
                        libres_indices_emp.add(s * 7 + 4)
                    else:
                        libres_indices_emp.add(s * 7 + 5)
                else:
                    dias_permitidos = list(range(s * 7, (s + 1) * 7))
                    libres_indices_emp.update(random.sample(dias_permitidos, cant_libres))

            dias_libres_emp[cod] = libres_indices_emp

    for idx_dia, col_nombre in enumerate(columnas_fechas):
        fecha_col = fechas_dt[idx_dia]
        fecha_actual_date = fecha_col.date()
        nombre_dia_semana = dias_semana_es[fecha_col.weekday()]
        dia_matriz_14 = idx_dia % 14
        dia_semana_idx = idx_dia % 7

        turnos_usados_analistas_hoy = set()
        libres_analistas_hoy = 0

        cargos_activos = set([d["CARGO"] for d in programacion_matriz.values()])

        for cargo in cargos_activos:
            cargo_clean = str(cargo).strip().upper()
            is_analista = "ANALISTA" in cargo_clean
            turnos_disponibles_dia = list(reglas_demanda.get(cargo_clean, {}).get(nombre_dia_semana, []))
            
            empleados_cargo = [cod for cod, d in programacion_matriz.items() if d["CARGO"] == cargo_clean]
            random.shuffle(empleados_cargo)

            for cod_emp in empleados_cargo:
                d_emp = programacion_matriz[cod_emp]

                if d_emp["INCIDENCIA_TIPO"] and d_emp["INCIDENCIA_INI"] and d_emp["INCIDENCIA_FIN"]:
                    if d_emp["INCIDENCIA_INI"] <= fecha_actual_date <= d_emp["INCIDENCIA_FIN"]:
                        programacion_matriz[cod_emp][col_nombre] = d_emp["INCIDENCIA_TIPO"]
                        d_emp["TURNO_FIJO_BLOQUE"] = None
                        continue

                if is_analista:
                    p_base = d_emp["PATRON_BASE"]
                    turno_sugerido = patrones_analistas[p_base][dia_matriz_14]

                    if turno_sugerido == "L" and dia_semana_idx <= 4:
                        if libres_analistas_hoy >= 1:
                            candidatos = [t for t in patrones_analistas[p_base] if t not in ["L", "vacaciones"] and t not in turnos_usados_analistas_hoy]
                            turno_sugerido = candidatos[0] if candidatos else "08:00-17:00"
                        else:
                            libres_analistas_hoy += 1

                    if turno_sugerido not in ["L", "vacaciones"]:
                        if turno_sugerido in turnos_usados_analistas_hoy:
                            alt_turnos = [t for t in CATALOGO_TURNOS if t not in turnos_usados_analistas_hoy]
                            turno_sugerido = alt_turnos[0] if alt_turnos else "AO"
                        if turno_sugerido != "AO":
                            turnos_usados_analistas_hoy.add(turno_sugerido)

                    programacion_matriz[cod_emp][col_nombre] = turno_sugerido

                else:
                    if idx_dia in dias_libres_emp.get(cod_emp, set()):
                        programacion_matriz[cod_emp][col_nombre] = "L"
                        d_emp["TURNO_FIJO_BLOQUE"] = None
                        continue

                    turno_a_asignar = None

                    if d_emp["TURNO_FIJO_BLOQUE"] is not None:
                        mismo_turno = d_emp["TURNO_FIJO_BLOQUE"]
                        if mismo_turno in turnos_disponibles_dia:
                            turno_a_asignar = mismo_turno
                            turnos_disponibles_dia.remove(mismo_turno)
                        else:
                            franja_buscada = clasificar_franja(mismo_turno)
                            candidatos_misma_franja = [t for t in turnos_disponibles_dia if clasificar_franja(t) == franja_buscada]
                            if candidatos_misma_franja:
                                turno_a_asignar = candidatos_misma_franja[0]
                                turnos_disponibles_dia.remove(turno_a_asignar)

                    if turno_a_asignar is None and turnos_disponibles_dia:
                        franja_ult = d_emp["ULTIMA_FRANJA"]
                        franjas_permitidas = obtener_siguiente_franja_permitida(franja_ult) if franja_ult else ["NOCHE", "TARDE", "MAÑANA"]

                        turnos_candidatos = list(turnos_disponibles_dia)
                        turnos_candidatos.sort(key=lambda t: 0 if clasificar_franja(t) in franjas_permitidas else 1)

                        if turnos_candidatos:
                            turno_a_asignar = turnos_candidatos[0]
                            turnos_disponibles_dia.remove(turno_a_asignar)

                    if turno_a_asignar is None:
                        programacion_matriz[cod_emp][col_nombre] = "AO"
                        d_emp["TURNO_FIJO_BLOQUE"] = None
                    else:
                        programacion_matriz[cod_emp][col_nombre] = turno_a_asignar
                        d_emp["TURNO_FIJO_BLOQUE"] = turno_a_asignar
                        d_emp["ULTIMA_FRANJA"] = clasificar_franja(turno_a_asignar)

    return pd.DataFrame([{k: v for k, v in datos.items() if k not in ["INCIDENCIA_TIPO", "INCIDENCIA_INI", "INCIDENCIA_FIN", "PATRON_BASE", "TURNO_FIJO_BLOQUE", "ULTIMA_FRANJA", "CARGO_ORIGINAL"]} for datos in programacion_matriz.values()]), programacion_matriz

# ==========================================
# 4. GENERACIÓN DE RESULTADOS Y AUDITORÍA
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

    # ==========================================
    # 5. REPORTE DE REEMPLAZOS INTER-CARGO
    # ==========================================
    st.markdown("---")
    st.subheader("🔄 Reporte de Cobertura de Turnos de Diferente Cargo")

    reporte_coberturas = []
    cols_fechas_malla = [c for c in df_resultado.columns if c not in ["CODIGO", "NOMBRE", "CARGO"]]

    for cod_emp, d_emp in dict_matriz.items():
        cargo_orig = d_emp["CARGO_ORIGINAL"]
        cargo_asig = d_emp["CARGO"]

        if cargo_orig != cargo_asig:
            for s in range(semanas):
                cols_semana = cols_fechas_malla[s * 7 : (s + 1) * 7]
                turnos_semana = [d_emp[col] for col in cols_semana if col in d_emp]
                turnos_filtrados = [t for t in turnos_semana if t not in ["L", "AO", "VACACIONES", "LICENCIA", "INCAPACIDAD", "PERMISO"]]

                reporte_coberturas.append({
                    "SEMANA": f"Semana {s + 1}",
                    "CÓDIGO": cod_emp,
                    "NOMBRE": d_emp["NOMBRE"],
                    "CARGO PERTENECIENTE": cargo_orig,
                    "CARGO CUBIERTO": cargo_asig,
                    "TURNOS REALIZADOS EN LA SEMANA": ", ".join(set(turnos_filtrados)) if turnos_filtrados else "Solo descanso/AO"
                })

    df_coberturas = pd.DataFrame(reporte_coberturas)

    if not df_coberturas.empty:
        st.info("ℹ️ A continuación se detalla el personal que realizó tareas o cubrió puestos de un cargo distinto al suyo:")
        st.dataframe(df_coberturas, use_container_width=True)
    else:
        st.success("✅ Todo el personal trabajó al 100% en tareas correspondientes a su cargo original.")

    st.markdown("---")

    # ==========================================
    # 6. REPORTE DE TAREAS NO CUBIERTAS
    # ==========================================
    st.subheader("🚨 Reporte de Tareas NO CUBIERTAS por Faltante de Personal")

    reporte_incompletos = []

    for col_f in cols_fechas_malla:
        dia_nombre_ext = col_f.split("\n")[0].upper()
        tipo_col_mat = "SABADO" if dia_nombre_ext in ["SÁBADO", "SABADO"] else ("DOMINGO" if dia_nombre_ext == "DOMINGO" else "HABIL")

        cargos_evaluar = df_empleados["CARGO"].unique().tolist()
        for cargo in cargos_evaluar:
            cargo_clean = str(cargo).strip().upper()
            sub_mat = df_tareas_req[df_tareas_req["CARGO"] == cargo_clean] if "CARGO" in df_tareas_req.columns else pd.DataFrame()
            col_target_mat = next((c for c in sub_mat.columns if tipo_col_mat in c), None)
            
            turnos_req = [str(x).strip() for x in sub_mat[col_target_mat].dropna().tolist() if str(x).strip() != ""] if col_target_mat else []
            sub_malla = df_resultado[df_resultado["CARGO"] == cargo_clean]
            turnos_disp = [str(x).strip() for x in sub_malla[col_f].tolist() if str(x).strip().upper() not in ["L", "AO", "VACACIONES", "LICENCIA", "INCAPACIDAD", "PERMISO"]]

            cant_req = len(turnos_req)
            cant_disp = len(turnos_disp)
            diferencia = cant_disp - cant_req

            if diferencia < 0:
                faltantes = abs(diferencia)
                copia_disp = list(turnos_disp)
                turnos_sin_cubrir = []
                for tr in turnos_req:
                    if tr in copia_disp:
                        copia_disp.remove(tr)
                    else:
                        turnos_sin_cubrir.append(tr)
                
                texto_turnos_faltantes = ", ".join(turnos_sin_cubrir) if turnos_sin_cubrir else "Variación en capacidad"

                reporte_incompletos.append({
                    "FECHA": col_f.replace("\n", " "),
                    "CARGO": cargo_clean,
                    "TIPO DÍA": tipo_col_mat.capitalize(),
                    "TURNO NO CUBIERTO": texto_turnos_faltantes,
                    "TAREAS REQUERIDAS": cant_req,
                    "PERSONAL ASIGNADO": cant_disp,
                    "PERSONAS FALTANTES": faltantes,
                    "DETALLE DEL FALTANTE": f"Faltan {faltantes} personas para cubrir el turno: {texto_turnos_faltantes}"
                })

    df_incumplidas = pd.DataFrame(reporte_incompletos)

    if not df_incumplidas.empty:
        st.error(f"⚠️ Atención: Hay {len(df_incumplidas)} días/cargos donde la operación NO se cubre completamente.")
        m1, m2 = st.columns(2)
        with m1:
            st.metric("Total de Días/Cargos Incompletos", len(df_incumplidas))
        with m2:
            st.metric("Total de Turnos/Tareas Desatendidas", df_incumplidas["PERSONAS FALTANTES"].sum())
        
        st.dataframe(df_incumplidas, use_container_width=True)
    else:
        st.success("🎉 ¡Todas las tareas operativas requeridas quedan cubiertas al 100%!")
