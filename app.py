from datetime import datetime, timedelta
import io
import re
import random
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Generador de Horarios Pro", layout="wide")
st.title("📅 Generador de Horarios Avanzado & Módulo de Tareas")

# ==========================================
# CATÁLOGO DE HORARIOS PREDEFINIDOS
# ==========================================
CATALOGO_TURNOS = [
    "03:00-11:00",
    "06:00-15:00",
    "07:00-16:00",
    "08:00-15:00 CAP",
    "08:00-17:00",
    "11:00-19:00",
    "11:00-19:00 AT",
    "11:00-27:00 AT",
    "19:00-27:00",
    "19:00-35:00 AT",
    "20:00-28:00",
    "22:00-30:00",
]

# ==========================================
# TURNOS DEFAULT SEGÚN TABLA DE CARGOS
# ==========================================
TURNOS_DEFAULT_POR_CARGO = {
    "Analista de Operaciones": {
        "habil": ["06:00-15:00", "11:00-19:00", "03:00-11:00", "19:00-27:00"],
        "sabado": ["06:00-15:00", "11:00-19:00", "03:00-11:00", "19:00-27:00"],
        "domingo": ["11:00-19:00", "03:00-11:00", "19:00-27:00"]
    },
    "Auxiliar de Operaciones": {
        "habil": ["11:00-19:00", "03:00-11:00", "19:00-27:00", "20:00-28:00"],
        "sabado": ["11:00-19:00", "03:00-11:00", "19:00-27:00", "22:00-30:00"],
        "domingo": ["11:00-19:00", "03:00-11:00", "19:00-27:00", "20:00-28:00"]
    },
    "Operador Líder": {
        "habil": ["20:00-28:00"],
        "sabado": ["20:00-28:00"],
        "domingo": ["20:00-28:00"]
    },
    "Técnico de Operaciones": {
        "habil": ["11:00-19:00", "03:00-11:00", "19:00-27:00"],
        "sabado": ["11:00-19:00", "03:00-11:00", "19:00-27:00"],
        "domingo": ["11:00-19:00", "03:00-11:00", "19:00-27:00"]
    },
    "Auxiliar de Alistamiento": {
        "habil": ["20:00-28:00", "22:00-30:00", "08:00-17:00"],
        "sabado": ["20:00-28:00", "22:00-30:00", "08:00-17:00"],
        "domingo": ["20:00-28:00", "22:00-30:00", "08:00-17:00"]
    }
}

# ==========================================
# CONFIGURACIÓN GENERAL
# ==========================================
st.sidebar.header("⚙️ Parámetros de Programación")
semanas = st.sidebar.slider("Semanas a generar", 1, 4, 2)
fecha_inicio_date = st.sidebar.date_input("Fecha de inicio", datetime.now())

st.sidebar.markdown("---")
st.sidebar.header("🏖️ Configuración de Descansos")
libres_por_semana_base = st.sidebar.number_input("Días libres por semana (normal)", min_value=1, max_value=3, value=1)
tiene_festivo = st.sidebar.checkbox("¿Hay día festivo en el periodo?", value=False)

fechas_festivas_sel = []
if tiene_festivo:
    dias_totales_temp = semanas * 7
    fechas_posibles = [(fecha_inicio_date + timedelta(days=i)) for i in range(dias_totales_temp)]
    fechas_festivas_sel = st.sidebar.multiselect(
        "Seleccionar día(s) festivo(s):",
        options=fechas_posibles,
        format_func=lambda x: x.strftime("%d-%b-%Y")
    )

dias_semana_es = ["LUNES", "MARTES", "MIÉRCOLES", "JUEVES", "VIERNES", "SÁBADO", "DOMINGO"]

# ==========================================
# FUNCIONES AUXILIARES DE TIEMPO Y TURNO
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

def calcular_descanso_suficiente(salida_previa_dt, entrada_actual_dt, min_horas=12):
    if salida_previa_dt is None:
        return True
    diferencia_horas = (entrada_actual_dt - salida_previa_dt).total_seconds() / 3600.0
    return diferencia_horas >= min_horas

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

def calcular_patron_dias_libres(semanas_count, libres_base, festivos_list, fecha_inicio_dt, dias_consecutivos_previos=0):
    dias_totales = semanas_count * 7
    intentos = 0
    while intentos < 500:
        intentos += 1
        dias_libres_indices = set()
        
        for s in range(semanas_count):
            inicio_sem = fecha_inicio_dt + timedelta(days=s * 7)
            fin_sem = inicio_sem + timedelta(days=6)
            hay_festivo_sem = any(inicio_sem.date() <= f <= fin_sem.date() for f in festivos_list)
            cant_libres_semana = libres_base + (1 if hay_festivo_sem else 0)
            
            dias_semana = list(range(s * 7, (s + 1) * 7))
            libres_sem = random.sample(dias_semana, cant_libres_semana)
            dias_libres_indices.update(libres_sem)
        
        valido = True
        consecutivos = dias_consecutivos_previos
        for dia_i in range(dias_totales):
            if dia_i in dias_libres_indices:
                consecutivos = 0
            else:
                consecutivos += 1
                if consecutivos > 10:
                    valido = False
                    break
        if valido:
            return dias_libres_indices

    return set([random.randint(s * 7, (s * 7) + 6) for s in range(semanas_count)])

# ==========================================
# 1. CARGA DE PERSONAL Y SEMANA ANTERIOR
# ==========================================
st.subheader("1. Cargar Lista de Personal y Semana Anterior")

col_f1, col_f2 = st.columns(2)
df_empleados = None
df_semana_anterior = None

with col_f1:
    uploaded_file = st.file_uploader("1. Subir Lista de Personal (Excel / CSV)", type=["xlsx", "csv"], key="file_personal")

with col_f2:
    uploaded_prev_file = st.file_uploader("2. Subir Malla de la Semana Anterior (Opcional)", type=["xlsx", "csv"], key="file_prev")

if uploaded_file is not None:
    try:
        df_empleados = pd.read_csv(uploaded_file) if uploaded_file.name.endswith(".csv") else pd.read_excel(uploaded_file)
        df_empleados.columns = [str(c).upper().strip() for c in df_empleados.columns]

        columnas_requeridas = {"CODIGO", "NOMBRE", "CARGO"}
        if not columnas_requeridas.issubset(set(df_empleados.columns)):
            st.error(f"El archivo debe contener las columnas: {columnas_requeridas}")
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
        st.error(f"Error al procesar archivo de personal: {e}")

if uploaded_prev_file is not None:
    try:
        df_semana_anterior = pd.read_csv(uploaded_prev_file) if uploaded_prev_file.name.endswith(".csv") else pd.read_excel(uploaded_prev_file)
        df_semana_anterior.columns = [str(c).upper().strip() for c in df_semana_anterior.columns]
        st.info("✅ Malla anterior cargada correctamente.")
    except Exception as e:
        st.warning(f"No se pudo leer la semana anterior: {e}")

# ==========================================
# 2. CONFIGURACIÓN COMPLETA DE REGLAS Y CUPOS
# ==========================================
matriz_demanda = {}

if df_empleados is not None:
    col_tulo, col_chk_guardar = st.columns([2.5, 1.5])
    with col_tulo:
        st.subheader("2. Configuración de Horarios y Requerimientos por Cargo")
    with col_chk_guardar:
        usar_default = st.checkbox(
            "📌 Mantener turnos default (Predeterminados)", 
            value=True, 
            help="Si está marcado, se aplicarán los turnos base automáticamente."
        )

    cargos_unicos = df_empleados["CARGO"].dropna().unique().tolist()

    for cargo in cargos_unicos:
        defaults_cargo = TURNOS_DEFAULT_POR_CARGO.get(
            cargo, 
            {"habil": ["03:00-11:00", "11:00-19:00", "19:00-27:00"], "sabado": ["08:00-15:00 CAP"], "domingo": ["08:00-15:00 CAP"]}
        )
        
        matriz_demanda[cargo] = {d: {} for d in dias_semana_es}
        dias_habiles = ["LUNES", "MARTES", "MIÉRCOLES", "JUEVES", "VIERNES"]
        
        for dh in dias_habiles:
            for t_def in defaults_cargo["habil"]:
                matriz_demanda[cargo][dh][t_def] = 1
        for t_def in defaults_cargo["sabado"]:
            matriz_demanda[cargo]["SÁBADO"][t_def] = 1
        for t_def in defaults_cargo["domingo"]:
            matriz_demanda[cargo]["DOMINGO"][t_def] = 1

        with st.expander(f"🔹 Configuración para: {cargo}", expanded=not usar_default):
            t_base_lv = st.multiselect("Turnos base Lunes-Viernes", options=CATALOGO_TURNOS, default=defaults_cargo["habil"], key=f"ms_base_{cargo}")
            txt_base_lv = st.text_input("Nuevo turno L-V (opcional)", key=f"tx_base_{cargo}")
            turnos_final_lv = list(t_base_lv) + ([txt_base_lv.strip()] if txt_base_lv.strip() else [])

            for d_reg in dias_habiles:
                matriz_demanda[cargo][d_reg] = {t: 1 for t in turnos_final_lv}

            col_sab, col_dom = st.columns(2)
            with col_sab:
                t_sab = st.multiselect("Turnos Sábado", options=CATALOGO_TURNOS, default=defaults_cargo["sabado"], key=f"ms_sab_{cargo}")
                matriz_demanda[cargo]["SÁBADO"] = {t: 1 for t in t_sab}
            with col_dom:
                t_dom = st.multiselect("Turnos Domingo", options=CATALOGO_TURNOS, default=defaults_cargo["domingo"], key=f"ms_dom_{cargo}")
                matriz_demanda[cargo]["DOMINGO"] = {t: 1 for t in t_dom}

# ==========================================
# 3. GENERACIÓN DE MALLA
# ==========================================
def generar_malla_matriz(df_personal, semanas_count, fecha_base_date, reglas_demanda, libres_base, festivos_list, df_prev=None):
    dias_totales = semanas_count * 7
    fecha_base = datetime.combine(fecha_base_date, datetime.min.time())
    anio_ref = fecha_base_date.year

    columnas_fechas = []
    fechas_dt = []
    for i in range(dias_totales):
        f_actual = fecha_base + timedelta(days=i)
        fechas_dt.append(f_actual)
        nombre_dia = dias_semana_es[f_actual.weekday()]
        fecha_fmt = f_actual.strftime("%d-%b")
        columnas_fechas.append(f"{nombre_dia}\n{fecha_fmt}")

    programacion_matriz = {}
    dias_libres_emp = {}
    info_historial = {}

    if df_prev is not None and "CODIGO" in df_prev.columns:
        cols_dias_prev = [c for c in df_prev.columns if c not in ["CODIGO", "NOMBRE", "CARGO", "ESTADO"]]
        for _, fila in df_prev.iterrows():
            c_cod = str(fila["CODIGO"]).strip()
            val_ult = str(fila[cols_dias_prev[-1]]).strip() if cols_dias_prev else "L"
            
            dias_consecutivos = 0
            for col_d in reversed(cols_dias_prev):
                val_dia = str(fila[col_d]).strip().upper()
                if val_dia in ["L", "L (DESCANSO)", "VACACIONES", "PERMISO", "LICENCIA", "INCAPACIDAD"]:
                    break
                dias_consecutivos += 1

            salida_calc = None
            franja_calc = None
            vino_desc = val_ult.upper() in ["L", "L (DESCANSO)", "VACACIONES", "PERMISO", "LICENCIA", "INCAPACIDAD"]
            
            if not vino_desc:
                parsed_h = extraer_horas(val_ult)
                if parsed_h:
                    h_i, m_i, h_f, m_f = parsed_h
                    dia_dom_prev = fecha_base - timedelta(days=1)
                    if h_f >= 24:
                        salida_calc = dia_dom_prev.replace(hour=h_f - 24, minute=m_f) + timedelta(days=1)
                    elif h_f < h_i:
                        salida_calc = dia_dom_prev.replace(hour=h_f, minute=m_f) + timedelta(days=1)
                    else:
                        salida_calc = dia_dom_prev.replace(hour=h_f, minute=m_f)
                    franja_calc = clasificar_franja(val_ult)

            info_historial[c_cod] = {
                "salida": salida_calc,
                "franja": franja_calc,
                "vino_descanso": vino_desc,
                "dias_consecutivos": dias_consecutivos
            }

    for _, emp in df_personal.iterrows():
        cod = str(emp["CODIGO"]).strip()
        hist = info_historial.get(cod, {})

        programacion_matriz[cod] = {
            "CODIGO": cod,
            "NOMBRE": emp["NOMBRE"],
            "CARGO": emp["CARGO"],
            "ESTADO": str(emp.get("ESTADO", "ACTIVO")).strip().upper(),
            "INCIDENCIA_TIPO": emp.get("INCIDENCIA_TIPO"),
            "INCIDENCIA_INI": parsear_fecha_incidencia(emp.get("INCIDENCIA_INI"), anio_ref),
            "INCIDENCIA_FIN": parsear_fecha_incidencia(emp.get("INCIDENCIA_FIN"), anio_ref),
            "SALIDA_PREVIA": hist.get("salida"),
            "ULTIMA_FRANJA": hist.get("franja"),
            "VINO_DE_DESCANSO": hist.get("vino_descanso", False),
        }

        dias_libres_emp[cod] = calcular_patron_dias_libres(
            semanas_count, libres_base, festivos_list, fecha_base, hist.get("dias_consecutivos", 0)
        )

    for idx_dia, col_nombre in enumerate(columnas_fechas):
        fecha_col = fechas_dt[idx_dia]
        fecha_actual_date = fecha_col.date()
        nombre_dia_semana = dias_semana_es[fecha_col.weekday()]
        
        empleados_bloqueados_hoy = set()
        for cod_emp, d in programacion_matriz.items():
            if d["INCIDENCIA_TIPO"] and d["INCIDENCIA_INI"] and d["INCIDENCIA_FIN"]:
                if d["INCIDENCIA_INI"] <= fecha_actual_date <= d["INCIDENCIA_FIN"]:
                    programacion_matriz[cod_emp][col_nombre] = d["INCIDENCIA_TIPO"]
                    programacion_matriz[cod_emp]["SALIDA_PREVIA"] = None
                    programacion_matriz[cod_emp]["VINO_DE_DESCANSO"] = True
                    empleados_bloqueados_hoy.add(cod_emp)

        for cod_emp in programacion_matriz.keys():
            if idx_dia in dias_libres_emp[cod_emp] and cod_emp not in empleados_bloqueados_hoy:
                programacion_matriz[cod_emp][col_nombre] = "L"
                programacion_matriz[cod_emp]["SALIDA_PREVIA"] = None
                programacion_matriz[cod_emp]["VINO_DE_DESCANSO"] = True
                empleados_bloqueados_hoy.add(cod_emp)

        for cargo, dem_dias in reglas_demanda.items():
            cupos_hoy = dem_dias.get(nombre_dia_semana, {})
            empleados_cargo = [cod for cod, d in programacion_matriz.items() if d["CARGO"] == cargo and cod not in empleados_bloqueados_hoy]
            random.shuffle(empleados_cargo)
            
            turnos_a_cubrir = []
            for t_nom, req in cupos_hoy.items():
                turnos_a_cubrir.extend([t_nom] * req)

            for cod_emp in empleados_cargo:
                salida_ant = programacion_matriz[cod_emp]["SALIDA_PREVIA"]
                vino_de_descanso = programacion_matriz[cod_emp]["VINO_DE_DESCANSO"]
                min_descanso = 24 if vino_de_descanso else 12

                turno_asignado = None
                nueva_salida = None

                for t_candidato in list(turnos_a_cubrir):
                    parsed_h = extraer_horas(t_candidato)
                    if parsed_h:
                        h_ini, m_ini, h_fin, m_fin = parsed_h
                        entrada_dt = fecha_col.replace(hour=h_ini, minute=m_ini)

                        if calcular_descanso_suficiente(salida_ant, entrada_dt, min_horas=min_descanso):
                            turno_asignado = t_candidato
                            turnos_a_cubrir.remove(t_candidato)
                            
                            if h_fin >= 24:
                                nueva_salida = (fecha_col + timedelta(days=1)).replace(hour=h_fin - 24, minute=m_fin)
                            elif h_fin < h_ini:
                                nueva_salida = (fecha_col + timedelta(days=1)).replace(hour=h_fin, minute=m_fin)
                            else:
                                nueva_salida = fecha_col.replace(hour=h_fin, minute=m_fin)
                            break

                if turno_asignado:
                    programacion_matriz[cod_emp][col_nombre] = turno_asignado
                    programacion_matriz[cod_emp]["SALIDA_PREVIA"] = nueva_salida
                    programacion_matriz[cod_emp]["ULTIMA_FRANJA"] = clasificar_franja(turno_asignado)
                    programacion_matriz[cod_emp]["VINO_DE_DESCANSO"] = False
                else:
                    programacion_matriz[cod_emp][col_nombre] = "L"
                    programacion_matriz[cod_emp]["VINO_DE_DESCANSO"] = True

    filas_finales = []
    for cod_emp, datos in programacion_matriz.items():
        d_limpio = {k: v for k, v in datos.items() if k not in ["ESTADO", "SALIDA_PREVIA", "INCIDENCIA_TIPO", "INCIDENCIA_INI", "INCIDENCIA_FIN", "ULTIMA_FRANJA", "VINO_DE_DESCANSO"]}
        filas_finales.append(d_limpio)

    return pd.DataFrame(filas_finales)

# ==========================================
# 4. PROCESAMIENTO Y EXPORTACIÓN
# ==========================================
if df_empleados is not None:
    if st.button("⚡ Generar Malla Horaria Completa"):
        st.session_state.df_resultado = generar_malla_matriz(
            df_empleados, semanas, fecha_inicio_date, matriz_demanda, libres_por_semana_base, fechas_festivas_sel, df_semana_anterior
        )

if "df_resultado" in st.session_state:
    df_resultado = st.session_state.df_resultado
    st.subheader("3. Malla Horaria Generada")
    st.dataframe(df_resultado, use_container_width=True)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_resultado.to_excel(writer, index=False, sheet_name="Programación")
        ws = writer.sheets["Programación"]
        for row in ws.iter_rows(min_row=2, max_col=len(df_resultado.columns), max_row=len(df_resultado) + 1):
            for cell in row:
                if str(cell.value).strip().upper() == "L":
                    cell.style = "Accent1"

    st.download_button(
        label="📥 Descargar Excel Matriz con Formato",
        data=output.getvalue(),
        file_name=f"Horario_{semanas}_semanas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.markdown("---")

    # ==========================================
    # 5. CARGA Y ANÁLISIS DE TAREAS OPERATIVAS
    # ==========================================
    st.subheader("📋 4. Asignación y Cobertura de Tareas Operativas")
    st.markdown("Sube un archivo de tareas para comparar las **horas/personas requeridas vs disponibles por cargo**.")

    def generar_plantilla_tareas():
        ejemplo = pd.DataFrame([
            {"TAREA": "Alistamiento de Pedidos", "CARGO": "Auxiliar de Alistamiento", "HORAS_REQUERIDAS": 8, "PRIORIDAD": "ALTA"},
            {"TAREA": "Supervisión de Turno Noche", "CARGO": "Operador Líder", "HORAS_REQUERIDAS": 8, "PRIORIDAD": "CRÍTICA"},
            {"TAREA": "Control de Inventario", "CARGO": "Analista de Operaciones", "HORAS_REQUERIDAS": 6, "PRIORIDAD": "MEDIA"}
        ])
        buf = io.BytesIO()
        ejemplo.to_excel(buf, index=False)
        return buf.getvalue()

    st.download_button(
        label="📄 Descargar Plantilla de Ejemplo para Tareas",
        data=generar_plantilla_tareas(),
        file_name="Plantilla_Tareas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    file_tareas = st.file_uploader("Subir Archivo de Tareas (Excel/CSV)", type=["xlsx", "csv"], key="file_tareas")

    if file_tareas is not None:
        try:
            df_tareas = pd.read_csv(file_tareas) if file_tareas.name.endswith(".csv") else pd.read_excel(file_tareas)
            df_tareas.columns = [str(c).upper().strip() for c in df_tareas.columns]

            req_cols_t = {"TAREA", "CARGO", "HORAS_REQUERIDAS"}
            if not req_cols_t.issubset(set(df_tareas.columns)):
                st.error(f"El archivo debe tener las columnas: {req_cols_t}")
            else:
                cols_fechas_malla = [c for c in df_resultado.columns if c not in ["CODIGO", "NOMBRE", "CARGO"]]
                
                disponibilidad_por_cargo = {}
                for cargo in df_resultado["CARGO"].unique():
                    sub_df = df_resultado[df_resultado["CARGO"] == cargo]
                    total_turnos_activos = 0
                    
                    for col in cols_fechas_malla:
                        turnos = sub_df[col].apply(lambda x: 1 if str(x).strip().upper() not in ["L", "VACACIONES", "LICENCIA", "INCAPACIDAD", "PERMISO"] else 0)
                        total_turnos_activos += turnos.sum()
                    
                    disponibilidad_por_cargo[cargo] = total_turnos_activos * 8

                tareas_resumen = df_tareas.groupby("CARGO")["HORAS_REQUERIDAS"].sum().reset_index()
                tareas_resumen["HORAS_DISPONIBLES"] = tareas_resumen["CARGO"].map(disponibilidad_por_cargo).fillna(0)
                tareas_resumen["DIFERENCIA (HORAS)"] = tareas_resumen["HORAS_DISPONIBLES"] - tareas_resumen["HORAS_REQUERIDAS"]
                tareas_resumen["ESTADO_COBERTURA"] = tareas_resumen["DIFERENCIA (HORAS)"].apply(
                    lambda x: "✅ CUBIERTO" if x >= 0 else "❌ FALTAN HORAS"
                )

                st.markdown("#### 📊 Balance de Capacidad por Cargo")
                st.dataframe(tareas_resumen, use_container_width=True)

                st.markdown("#### 🎯 Cobertura de Tareas Individuales")
                df_tareas["DISPONIBLE_EN_CARGO"] = df_tareas["CARGO"].map(disponibilidad_por_cargo).fillna(0)
                df_tareas["ESTADO"] = df_tareas.apply(
                    lambda row: "✅ Se puede cubrir" if row["DISPONIBLE_EN_CARGO"] >= row["HORAS_REQUERIDAS"] else "⚠️ Riesgo / Faltan Personas",
                    axis=1
                )
                
                st.dataframe(df_tareas[["TAREA", "CARGO", "HORAS_REQUERIDAS", "ESTADO"]], use_container_width=True)

        except Exception as e:
            st.error(f"Error al analizar las tareas: {e}")
