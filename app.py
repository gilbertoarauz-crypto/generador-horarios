from datetime import datetime, timedelta
import io
import re
import random
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Generador de Horarios Pro", layout="wide")
st.title("📅 Generador de Horarios Avanzado")

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

dias_semana_es = [
    "LUNES",
    "MARTES",
    "MIÉRCOLES",
    "JUEVES",
    "VIERNES",
    "SÁBADO",
    "DOMINGO",
]

if "matriz_demanda_guardada" not in st.session_state:
    st.session_state.matriz_demanda_guardada = {}

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

def obtener_siguiente_franja_deseada(franja_actual):
    if franja_actual == "NOCHE":
        return "TARDE"
    elif franja_actual == "TARDE":
        return "MAÑANA"
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

def calcular_patron_dias_libres(semanas_count, libres_base, festivos_list, fecha_inicio_dt):
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
        consecutivos = 0
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
        if uploaded_file.name.endswith(".csv"):
            df_empleados = pd.read_csv(uploaded_file)
        else:
            df_empleados = pd.read_excel(uploaded_file)

        df_empleados.columns = [str(c).upper().strip() for c in df_empleados.columns]

        columnas_requeridas = {"CODIGO", "NOMBRE", "CARGO"}
        if not columnas_requeridas.issubset(set(df_empleados.columns)):
            st.error(f"El archivo debe contener al menos las columnas: {columnas_requeridas}")
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
            st.dataframe(df_empleados, use_container_width=True)
    except Exception as e:
        st.error(f"Error al procesar el archivo de personal: {e}")

if uploaded_prev_file is not None:
    try:
        if uploaded_prev_file.name.endswith(".csv"):
            df_semana_anterior = pd.read_csv(uploaded_prev_file)
        else:
            df_semana_anterior = pd.read_excel(uploaded_prev_file)
        
        df_semana_anterior.columns = [str(c).upper().strip() for c in df_semana_anterior.columns]
        st.info("✅ Malla de la semana anterior cargada para empalmar descansos y rotaciones.")
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
            help="Si está marcado, se aplicarán los turnos precargados por catálogo para cada cargo."
        )

    cargos_unicos = df_empleados["CARGO"].dropna().unique().tolist()

    for cargo in cargos_unicos:
        defaults_cargo = TURNOS_DEFAULT_POR_CARGO.get(
            cargo, 
            {"habil": ["03:00-11:00", "11:00-19:00", "19:00-27:00"], "sabado": ["08:00-15:00 CAP"], "domingo": ["08:00-15:00 CAP"]}
        )
        
        with st.expander(f"🔹 Configuración para: {cargo}", expanded=not usar_default):
            matriz_demanda[cargo] = {d: {} for d in dias_semana_es}
            
            st.markdown("#### 📌 Semana Laboral Base (Lunes a Viernes)")
            col_lv1, col_lv2 = st.columns([3, 2])
            with col_lv1:
                t_base_lv = st.multiselect(
                    "Turnos base Lunes-Viernes",
                    options=CATALOGO_TURNOS,
                    default=defaults_cargo["habil"],
                    key=f"ms_base_{cargo}"
                )
            with col_lv2:
                txt_base_lv = st.text_input("Nuevo turno L-V (opcional)", key=f"tx_base_{cargo}")
            
            turnos_final_lv = list(t_base_lv) + ([txt_base_lv.strip()] if txt_base_lv.strip() else [])

            st.markdown("**Marcar si algún día de la semana laboral es diferente:**")
            cols_chk = st.columns(5)
            dias_excepcion = {}
            dias_laborales = ["LUNES", "MARTES", "MIÉRCOLES", "JUEVES", "VIERNES"]
            
            for idx_d, d_nom in enumerate(dias_laborales):
                with cols_chk[idx_d]:
                    dias_excepcion[d_nom] = st.checkbox(f"{d_nom}", key=f"chk_{cargo}_{d_nom}")

            turnos_por_dia_lv = {}
            for d_nom in dias_laborales:
                if dias_excepcion[d_nom]:
                    st.markdown("---")
                    st.markdown(f"⚙️ **Excepción y Cupos para {d_nom}:**")
                    c1, c2 = st.columns([3, 2])
                    with c1:
                        t_spec = st.multiselect(
                            f"Turnos {d_nom}", 
                            options=CATALOGO_TURNOS, 
                            default=turnos_final_lv if turnos_final_lv else ["20:00-28:00"], 
                            key=f"ms_{cargo}_{d_nom}"
                        )
                    with c2:
                        txt_spec = st.text_input(f"Nuevo turno {d_nom} (opcional)", key=f"tx_{cargo}_{d_nom}")
                    
                    turnos_por_dia_lv[d_nom] = list(t_spec) + ([txt_spec.strip()] if txt_spec.strip() else [])
                    
                    if turnos_por_dia_lv[d_nom]:
                        st.markdown(f"**Cantidad de personas requeridas para {d_nom}:**")
                        cols_exc = st.columns(len(turnos_por_dia_lv[d_nom]))
                        for idx_te, t_nom_exc in enumerate(turnos_por_dia_lv[d_nom]):
                            with cols_exc[idx_te]:
                                cant_e = st.number_input(
                                    f"{d_nom}: {t_nom_exc}", 
                                    min_value=1, max_value=20, value=1, 
                                    key=f"num_exc_{cargo}_{d_nom}_{t_nom_exc}"
                                )
                                matriz_demanda[cargo][d_nom][t_nom_exc] = cant_e
                else:
                    turnos_por_dia_lv[d_nom] = turnos_final_lv

            dias_regulares = [d for d in dias_laborales if not dias_excepcion[d]]
            if dias_regulares and turnos_final_lv:
                st.markdown("---")
                st.markdown("**Cantidad de personas requeridas por turno (Días Regulares L-V):**")
                cols_c_lv = st.columns(len(turnos_final_lv))
                for idx_t, t_nom in enumerate(turnos_final_lv):
                    with cols_c_lv[idx_t]:
                        cant_reg = st.number_input(
                            f"L-V Base: {t_nom}", 
                            min_value=1, max_value=20, value=1, 
                            key=f"num_lv_base_{cargo}_{t_nom}"
                        )
                        for d_reg in dias_regulares:
                            matriz_demanda[cargo][d_reg][t_nom] = cant_reg

            st.markdown("---")
            st.markdown("#### 🗓️ Fin de Semana (Sábado y Domingo)")
            col_sab, col_dom = st.columns(2)

            with col_sab:
                st.markdown("**SÁBADO**")
                t_sab = st.multiselect("Turnos Sábado", options=CATALOGO_TURNOS, default=defaults_cargo["sabado"], key=f"ms_sab_{cargo}")
                txt_sab = st.text_input("Nuevo turno Sábado (opcional)", key=f"tx_sab_{cargo}")
                list_sab = list(t_sab) + ([txt_sab.strip()] if txt_sab.strip() else [])
                for t_nom in list_sab:
                    cant_s = st.number_input(f"Sáb: {t_nom}", min_value=1, max_value=20, value=1, key=f"num_sab_{cargo}_{t_nom}")
                    matriz_demanda[cargo]["SÁBADO"][t_nom] = cant_s

            with col_dom:
                st.markdown("**DOMINGO**")
                t_dom = st.multiselect("Turnos Domingo", options=CATALOGO_TURNOS, default=defaults_cargo["domingo"], key=f"ms_dom_{cargo}")
                txt_dom = st.text_input("Nuevo turno Domingo (opcional)", key=f"tx_dom_{cargo}")
                list_dom = list(t_dom) + ([txt_dom.strip()] if txt_dom.strip() else [])
                for t_nom in list_dom:
                    cant_d = st.number_input(f"Dom: {t_nom}", min_value=1, max_value=20, value=1, key=f"num_dom_{cargo}_{t_nom}")
                    matriz_demanda[cargo]["DOMINGO"][t_nom] = cant_d

        st.session_state.matriz_demanda_guardada = matriz_demanda

# ==========================================
# 3. LÓGICA DE GENERACIÓN EMPALMADA (CORREGIDA)
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
        ult_col = df_prev.columns[-1]
        for _, fila in df_prev.iterrows():
            c_cod = fila["CODIGO"]
            val_ult = str(fila[ult_col]).strip()
            
            salida_calc = None
            franja_calc = None
            vino_desc = False
            
            if val_ult.upper() in ["L", "L (DESCANSO)", "VACACIONES", "PERMISO", "LICENCIA"]:
                vino_desc = True
            else:
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
                "vino_descanso": vino_desc
            }

    analistas_cods = [
        row["CODIGO"] for _, row in df_personal.iterrows() 
        if "ANALISTA" in str(row["CARGO"]).upper()
    ]

    for _, emp in df_personal.iterrows():
        cod = emp["CODIGO"]
        f_ini_inc = parsear_fecha_incidencia(emp.get("INCIDENCIA_INI"), anio_ref)
        f_fin_inc = parsear_fecha_incidencia(emp.get("INCIDENCIA_FIN"), anio_ref)
        tipo_inc = str(emp.get("INCIDENCIA_TIPO", "")).strip().upper()

        hist = info_historial.get(cod, {})

        programacion_matriz[cod] = {
            "CODIGO": cod,
            "NOMBRE": emp["NOMBRE"],
            "CARGO": emp["CARGO"],
            "ESTADO": str(emp.get("ESTADO", "ACTIVO")).strip().upper(),
            "INCIDENCIA_TIPO": tipo_inc if tipo_inc and tipo_inc != "NAN" else None,
            "INCIDENCIA_INI": f_ini_inc,
            "INCIDENCIA_FIN": f_fin_inc,
            "SALIDA_PREVIA": hist.get("salida"),
            "ULTIMA_FRANJA": hist.get("franja"),
            "VINO_DE_DESCANSO": hist.get("vino_descanso", False),
        }

        # Asignar patrones de libres asegurando máximo 1 analista libre en L-V
        libres_calculados = calcular_patron_dias_libres(semanas_count, libres_base, festivos_list, fecha_base)
        
        if cod in analistas_cods:
            for s in range(semanas_count):
                libres_sem = [idx for idx in libres_calculados if (s * 7) <= idx < ((s + 1) * 7)]
                for libre_idx in libres_sem:
                    # CORRECCIÓN DE ERROR EN FECHA INDEXADA
                    dt_libre = fecha_base + timedelta(days=libre_idx)
                    nombre_d = dias_semana_es[dt_libre.weekday()]
                    if nombre_d in ["LUNES", "MARTES", "MIÉRCOLES", "JUEVES", "VIERNES"]:
                        libres_ya = sum(1 for a_c in dias_libres_emp if a_c in analistas_cods and libre_idx in dias_libres_emp[a_c])
                        if libres_ya >= 1:
                            dias_posibles = [d_i for d_i in range(s * 7, (s + 1) * 7) if d_i not in libres_calculados]
                            if dias_posibles:
                                libres_calculados.remove(libre_idx)
                                libres_calculados.add(random.choice(dias_posibles))

        dias_libres_emp[cod] = libres_calculados

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
            
            empleados_cargo = [
                cod for cod, d in programacion_matriz.items()
                if d["CARGO"] == cargo and cod not in empleados_bloqueados_hoy
            ]
            
            random.shuffle(empleados_cargo)
            
            turnos_a_cubrir = []
            for t_nom, req in cupos_hoy.items():
                turnos_a_cubrir.extend([t_nom] * req)

            for cod_emp in empleados_cargo:
                salida_ant = programacion_matriz[cod_emp]["SALIDA_PREVIA"]
                vino_de_descanso = programacion_matriz[cod_emp]["VINO_DE_DESCANSO"]
                franja_ant = programacion_matriz[cod_emp]["ULTIMA_FRANJA"]
                
                min_descanso = 24 if vino_de_descanso else 12
                franja_deseada = obtener_siguiente_franja_deseada(franja_ant) if (vino_de_descanso and franja_ant) else None

                turno_asignado = None
                nueva_salida = None

                turnos_candidatos_ordenados = list(turnos_a_cubrir)
                if franja_deseada:
                    turnos_candidatos_ordenados.sort(
                        key=lambda x: 0 if clasificar_franja(x) == franja_deseada else 1
                    )

                for t_candidato in turnos_candidatos_ordenados:
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
                
                if turno_asignado is None:
                    turnos_respaldo = list(cupos_hoy.keys())
                    if franja_deseada:
                        turnos_respaldo.sort(key=lambda x: 0 if clasificar_franja(x) == franja_deseada else 1)

                    for t_alt in turnos_respaldo:
                        parsed_h = extraer_horas(t_alt)
                        if parsed_h:
                            h_ini, m_ini, h_fin, m_fin = parsed_h
                            entrada_dt = fecha_col.replace(hour=h_ini, minute=m_ini)
                            if calcular_descanso_suficiente(salida_ant, entrada_dt, min_horas=min_descanso):
                                turno_asignado = t_alt
                                if h_fin >= 24:
                                    nueva_salida = (fecha_col + timedelta(days=1)).replace(hour=h_fin - 24, minute=m_fin)
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

    # AUDITORÍA Y CONTROL FINAL DE DÍAS LIBRES OBLIGATORIOS POR SEMANA
    for cod_emp, datos in programacion_matriz.items():
        for s in range(semanas_count):
            cols_semana = columnas_fechas[s * 7 : (s + 1) * 7]
            valores_semana = [datos.get(c) for c in cols_semana]
            
            tiene_descanso = any(
                str(val).upper() in ["L", "VACACIONES", "PERMISO", "LICENCIA", "INCAPACIDAD"] or "INCIDENCIA" in str(val).upper()
                for val in valores_semana
            )
            
            if not tiene_descanso:
                col_forzada = cols_semana[-1]
                datos[col_forzada] = "L"

    filas_finales = []
    for cod_emp, datos in programacion_matriz.items():
        d_limpio = {
            k: v for k, v in datos.items() 
            if k not in ["ESTADO", "SALIDA_PREVIA", "INCIDENCIA_TIPO", "INCIDENCIA_INI", "INCIDENCIA_FIN", "ULTIMA_FRANJA", "VINO_DE_DESCANSO"]
        }
        filas_finales.append(d_limpio)

    return pd.DataFrame(filas_finales)

# ==========================================
# 4. PROCESAMIENTO Y EXPORTACIÓN
# ==========================================
if df_empleados is not None:
    if st.button("⚡ Generar Malla Horaria Completa"):
        df_resultado = generar_malla_matriz(
            df_empleados, semanas, fecha_inicio_date, matriz_demanda, libres_por_semana_base, fechas_festivas_sel, df_semana_anterior
        )

        st.subheader("3. Malla Horaria Generada")
        st.dataframe(df_resultado, use_container_width=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_resultado.to_excel(writer, index=False, sheet_name="Programación")

        st.download_button(
            label="📥 Descargar Excel Matriz",
            data=output.getvalue(),
            file_name=f"Horario_{semanas}_semanas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
