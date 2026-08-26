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
    "Analista de Operaciones": {
        "habil": ["03:00-11:00", "11:00-19:00", "06:00-15:00", "19:00-27:00"],
        "sabado": ["03:00-11:00", "11:00-19:00", "19:00-27:00", "06:00-15:00"],
        "domingo": ["03:00-11:00", "11:00-19:00", "19:00-27:00"]
    },
    "Auxiliar de Operaciones": {
        "habil": ["03:00-11:00", "03:00-11:00", "11:00-19:00", "11:00-19:00", "19:00-27:00", "20:00-28:00"],
        "sabado": ["03:00-11:00", "03:00-11:00", "11:00-19:00", "11:00-19:00", "19:00-27:00", "22:00-30:00"],
        "domingo": ["03:00-11:00", "03:00-11:00", "11:00-19:00", "11:00-19:00", "19:00-27:00", "20:00-28:00"]
    },
    "Operador Líder": {
        "habil": ["20:00-28:00", "20:00-28:00"],
        "sabado": ["20:00-28:00"],
        "domingo": ["20:00-28:00", "20:00-28:00"]
    },
    "Técnico de Operaciones": {
        "habil": ["03:00-11:00", "11:00-19:00", "19:00-27:00", "06:00-15:00"],
        "sabado": ["03:00-11:00", "11:00-19:00", "19:00-27:00", "06:00-15:00"],
        "domingo": ["03:00-11:00", "11:00-19:00", "19:00-27:00", "06:00-15:00"]
    },
    "Auxiliar de Alistamiento": {
        "habil": ["08:00-17:00", "22:00-30:00", "20:00-28:00", "20:00-28:00"],
        "sabado": ["08:00-17:00", "20:00-28:00", "22:00-30:00"],
        "domingo": ["08:00-17:00", "20:00-28:00", "22:00-30:00", "20:00-28:00"]
    }
}

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
# FUNCIONES AUXILIARES
# ==========================================
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

def generar_dias_libres_exactos(semanas_count, libres_base, festivos_list, fecha_inicio_dt):
    """Garantiza la cantidad EXACTA de días libres asignados por semana sin exceder."""
    dias_libres_indices = set()
    for s in range(semanas_count):
        inicio_sem = fecha_inicio_dt + timedelta(days=s * 7)
        fin_sem = inicio_sem + timedelta(days=6)
        hay_festivo_sem = any(inicio_sem.date() <= f <= fin_sem.date() for f in festivos_list)
        cant_libres_semana = libres_base + (1 if hay_festivo_sem else 0)
        
        dias_semana = list(range(s * 7, (s + 1) * 7))
        dias_libres_indices.update(random.sample(dias_semana, cant_libres_semana))
    return dias_libres_indices

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
        df_tareas_req.columns = [str(c).strip() for c in df_tareas_req.columns]
        st.success("✅ Matriz de Tareas cargada.")
    except Exception as e:
        st.error(f"Error en archivo de tareas: {e}")
else:
    filas_default = []
    for c, t_dict in TURNOS_DEFAULT_POR_CARGO.items():
        max_len = max(len(t_dict["habil"]), len(t_dict["sabado"]), len(t_dict["domingo"]))
        for i in range(max_len):
            filas_default.append({
                "CARGO": c,
                "Habil": t_dict["habil"][i] if i < len(t_dict["habil"]) else "",
                "sabado": t_dict["sabado"][i] if i < len(t_dict["sabado"]) else "",
                "Domingo": t_dict["domingo"][i] if i < len(t_dict["domingo"]) else ""
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

# ==========================================
# 2. CONSTRUCCIÓN DE REGLAS DE DEMANDA
# ==========================================
matriz_demanda = {}

if df_empleados is not None:
    cargos_unicos = df_empleados["CARGO"].dropna().unique().tolist()
    for cargo in cargos_unicos:
        matriz_demanda[cargo] = {d: [] for d in dias_semana_es}
        sub_mat = df_tareas_req[df_tareas_req["CARGO"] == cargo] if "CARGO" in df_tareas_req.columns else pd.DataFrame()
        
        req_habil = [str(x).strip() for x in sub_mat["Habil"].dropna().tolist() if str(x).strip() != ""] if "Habil" in sub_mat.columns else []
        req_sab = [str(x).strip() for x in sub_mat["sabado"].dropna().tolist() if str(x).strip() != ""] if "sabado" in sub_mat.columns else []
        req_dom = [str(x).strip() for x in sub_mat["Domingo"].dropna().tolist() if str(x).strip() != ""] if "Domingo" in sub_mat.columns else []

        for dh in ["LUNES", "MARTES", "MIÉRCOLES", "JUEVES", "VIERNES"]:
            matriz_demanda[cargo][dh] = list(req_habil)
        matriz_demanda[cargo]["SÁBADO"] = list(req_sab)
        matriz_demanda[cargo]["DOMINGO"] = list(req_dom)

# ==========================================
# 3. GENERACIÓN DE MALLA HORARIA CON CONTROL RIGUROSO DE LIBRES
# ==========================================
def generar_malla_matriz(df_personal, semanas_count, fecha_base_date, reglas_demanda, libres_base, festivos_list):
    dias_totales = semanas_count * 7
    fecha_base = datetime.combine(fecha_base_date, datetime.min.time())
    anio_ref = fecha_base_date.year
    columnas_fechas, fechas_dt = [], []
    
    for i in range(dias_totales):
        f_actual = fecha_base + timedelta(days=i)
        fechas_dt.append(f_actual)
        columnas_fechas.append(f"{dias_semana_es[f_actual.weekday()]}\n{f_actual.strftime('%d-%b')}")

    programacion_matriz, dias_libres_emp = {}, {}

    # Pre-asignar días libres EXACTOS por empleado
    for _, emp in df_personal.iterrows():
        cod = str(emp["CODIGO"]).strip()
        programacion_matriz[cod] = {
            "CODIGO": cod, "NOMBRE": emp["NOMBRE"], "CARGO": emp["CARGO"],
            "INCIDENCIA_TIPO": emp.get("INCIDENCIA_TIPO"),
            "INCIDENCIA_INI": parsear_fecha_incidencia(emp.get("INCIDENCIA_INI"), anio_ref),
            "INCIDENCIA_FIN": parsear_fecha_incidencia(emp.get("INCIDENCIA_FIN"), anio_ref)
        }
        dias_libres_emp[cod] = generar_dias_libres_exactos(semanas_count, libres_base, festivos_list, fecha_base)

    # Recorrido día a día
    for idx_dia, col_nombre in enumerate(columnas_fechas):
        fecha_col = fechas_dt[idx_dia]
        fecha_actual_date = fecha_col.date()
        nombre_dia_semana = dias_semana_es[fecha_col.weekday()]

        for cargo in df_personal["CARGO"].unique():
            # Obtener lista de tareas/turnos a cubrir ese día
            turnos_disponibles_dia = list(reglas_demanda.get(cargo, {}).get(nombre_dia_semana, []))
            
            # Obtener empleados del cargo
            empleados_cargo = [cod for cod, d in programacion_matriz.items() if d["CARGO"] == cargo]
            random.shuffle(empleados_cargo)

            for cod_emp in empleados_cargo:
                d_emp = programacion_matriz[cod_emp]

                # 1. Verificar Incidencias (Vacaciones, Licencias, Incapacidades)
                if d_emp["INCIDENCIA_TIPO"] and d_emp["INCIDENCIA_INI"] and d_emp["INCIDENCIA_FIN"]:
                    if d_emp["INCIDENCIA_INI"] <= fecha_actual_date <= d_emp["INCIDENCIA_FIN"]:
                        programacion_matriz[cod_emp][col_nombre] = d_emp["INCIDENCIA_TIPO"]
                        continue

                # 2. Verificar Libre Programado (Estricto)
                if idx_dia in dias_libres_emp[cod_emp]:
                    programacion_matriz[cod_emp][col_nombre] = "L"
                    continue

                # 3. Asignar Turno si hay disponible en el catálogo del día
                if turnos_disponibles_dia:
                    turno_asignado = turnos_disponibles_dia.pop(0)
                    programacion_matriz[cod_emp][col_nombre] = turno_asignado
                else:
                    # Si ya no quedan tareas requeridas en la lista, se asigna un turno por defecto sin forzar un "L" extra
                    turnos_base_cargo = TURNOS_DEFAULT_POR_CARGO.get(cargo, {}).get("habil", ["08:00-17:00"])
                    programacion_matriz[cod_emp][col_nombre] = turnos_base_cargo[0]

    return pd.DataFrame([{k: v for k, v in datos.items() if k not in ["INCIDENCIA_TIPO", "INCIDENCIA_INI", "INCIDENCIA_FIN"]} for datos in programacion_matriz.values()])

# ==========================================
# 4. GENERACIÓN DE RESULTADOS
# ==========================================
if df_empleados is not None:
    if st.button("⚡ Generar Horarios y Auditar Tareas"):
        st.session_state.df_resultado = generar_malla_matriz(
            df_empleados, semanas, fecha_inicio_date, matriz_demanda, libres_por_semana_base, fechas_festivas_sel
        )

if "df_resultado" in st.session_state:
    df_resultado = st.session_state.df_resultado
    st.subheader("2. Malla Horaria Generada")
    st.dataframe(df_resultado, use_container_width=True)

    st.markdown("---")

    # ==========================================
    # 5. REPORTE EXCLUSIVO DE TAREAS NO CUBIERTAS
    # ==========================================
    st.subheader("🚨 Reporte de Tareas NO CUBIERTAS por Faltante de Personal")

    cols_fechas_malla = [c for c in df_resultado.columns if c not in ["CODIGO", "NOMBRE", "CARGO"]]
    reporte_incompletos = []

    for col_f in cols_fechas_malla:
        dia_nombre_ext = col_f.split("\n")[0].upper()
        tipo_col_mat = "sabado" if dia_nombre_ext in ["SÁBADO", "SABADO"] else ("Domingo" if dia_nombre_ext == "DOMINGO" else "Habil")

        for cargo in df_empleados["CARGO"].unique():
            sub_mat = df_tareas_req[df_tareas_req["CARGO"] == cargo] if "CARGO" in df_tareas_req.columns else pd.DataFrame()
            col_target_mat = next((c for c in sub_mat.columns if c.lower() == tipo_col_mat.lower()), None)
            
            turnos_req = [str(x).strip() for x in sub_mat[col_target_mat].dropna().tolist() if str(x).strip() != ""] if col_target_mat else []
            sub_malla = df_resultado[df_resultado["CARGO"] == cargo]
            turnos_disp = [str(x).strip() for x in sub_malla[col_f].tolist() if str(x).strip().upper() not in ["L", "VACACIONES", "LICENCIA", "INCAPACIDAD", "PERMISO"]]

            cant_req = len(turnos_req)
            cant_disp = len(turnos_disp)
            diferencia = cant_disp - cant_req

            if diferencia < 0:
                faltantes = abs(diferencia)
                reporte_incompletos.append({
                    "FECHA": col_f.replace("\n", " "),
                    "CARGO": cargo,
                    "TIPO DÍA": tipo_col_mat.capitalize(),
                    "TAREAS REQUERIDAS": cant_req,
                    "PERSONAL ASIGNADO": cant_disp,
                    "PERSONAS FALTANTES": faltantes,
                    "DETALLE DEL FALTANTE": f"Faltan {faltantes} personas para cubrir el total de tareas operativas"
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
