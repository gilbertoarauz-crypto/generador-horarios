from datetime import datetime, timedelta
import io
import re
import random
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Generador de Horarios Pro", layout="wide")
st.title("📅 Generador de Horarios con Continuidad e Histórico Semanal")

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
# CONFIGURACIÓN GENERAL
# ==========================================
st.sidebar.header("⚙️ Parámetros de Programación")
semanas = st.sidebar.slider("Semanas a generar", 1, 4, 2)
fecha_inicio_date = st.sidebar.date_input("Fecha de inicio", datetime.now())

dias_semana_es = [
    "LUNES",
    "MARTES",
    "MIÉRCOLES",
    "JUEVES",
    "VIERNES",
    "SÁBADO",
    "DOMINGO",
]

# ==========================================
# FUNCIONES AUXILIARES DE TIEMPO
# ==========================================
def extraer_horas(texto_turno):
    coincidencia = re.search(r"(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})", str(texto_turno))
    if coincidencia:
        h_ini, m_ini, h_fin, m_fin = map(int, coincidencia.groups())
        return h_ini, m_ini, h_fin, m_fin
    return None

def calcular_descanso_suficiente(salida_previa_dt, entrada_actual_dt, min_horas=12):
    if salida_previa_dt is None:
        return True
    diferencia_horas = (entrada_actual_dt - salida_previa_dt).total_seconds() / 3600.0
    return diferencia_horas >= min_horas

# ==========================================
# 1. CARGA DE ARCHIVOS
# ==========================================
st.subheader("1. Cargar Datos del Personal")
col_arch1, col_arch2 = st.columns(2)

with col_arch1:
    uploaded_file = st.file_uploader("1.1 Subir Lista de Personal (Obligatorio)", type=["xlsx", "csv"])

with col_arch2:
    uploaded_prev_file = st.file_uploader("1.2 Subir Malla Completa de Semana Anterior (Opcional)", type=["xlsx", "csv"])

df_empleados = None
historico_empleados = {} # Dict: {cod: {"salida_domingo": dt, "ultimo_turno": str, "ultimo_dia_libre": int}}

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
            st.success(f"¡Se cargaron {len(df_empleados)} empleados exitosamente!")
    except Exception as e:
        st.error(f"Error al procesar lista de personal: {e}")

# Analizar la malla semanal previa completa
if uploaded_prev_file is not None and df_empleados is not None:
    try:
        df_prev = pd.read_csv(uploaded_prev_file) if uploaded_prev_file.name.endswith(".csv") else pd.read_excel(uploaded_prev_file)
        df_prev.columns = [str(c).upper().strip() for c in df_prev.columns]
        
        # Identificar columnas de días
        cols_dias = [c for c in df_prev.columns if any(d in c for d in ["LUNES", "MARTES", "MIÉRCOLES", "JUEVES", "VIERNES", "SÁBADO", "DOMINGO"])]
        if not cols_dias:
            cols_dias = df_prev.columns[3:] # Si no detecta nombres, toma desde la cuarta columna

        for _, row in df_prev.iterrows():
            cod = str(row.get("CODIGO", "")).strip()
            
            # 1. Obtener último turno del domingo anterior para descanso de 12h
            turno_dom = str(row.get(cols_dias[-1], "")).strip()
            parsed_h = extraer_horas(turno_dom)
            dt_salida = None
            if parsed_h:
                h_ini, m_ini, h_fin, m_fin = parsed_h
                fecha_dom_prev = datetime.combine(fecha_inicio_date - timedelta(days=1), datetime.min.time())
                if h_fin >= 24:
                    dt_salida = (fecha_dom_prev + timedelta(days=1)).replace(hour=h_fin - 24, minute=m_fin)
                elif h_fin < h_ini:
                    dt_salida = (fecha_dom_prev + timedelta(days=1)).replace(hour=h_fin, minute=m_fin)
                else:
                    dt_salida = fecha_dom_prev.replace(hour=h_fin, minute=m_fin)
            
            # 2. Obtener último día libre (L)
            ultimo_libre_idx = -1
            for idx_d, col_d in enumerate(cols_dias):
                if str(row.get(col_d, "")).strip().upper() == "L":
                    ultimo_libre_idx = idx_d
            
            # 3. Guardar en estructura histórica
            historico_empleados[cod] = {
                "salida_domingo": dt_salida,
                "ultimo_turno": turno_dom,
                "ultimo_dia_libre": ultimo_libre_idx
            }
        
        st.info("🔄 Se procesó la semana previa: el algoritmo rotará los turnos trabajados y mantendrá el patrón de descansos.")
    except Exception as e:
        st.warning(f"No se pudo procesar la malla previa completa: {e}")

# ==========================================
# 2. CONFIGURACIÓN DE DEMANDA
# ==========================================
matriz_demanda = {}

if df_empleados is not None:
    st.subheader("2. Configuración de Horarios y Requerimientos por Cargo")
    cargos_unicos = df_empleados["CARGO"].dropna().unique().tolist()

    for cargo in cargos_unicos:
        with st.expander(f"🔹 Configuración para: {cargo}", expanded=True):
            matriz_demanda[cargo] = {d: {} for d in dias_semana_es}
            
            st.markdown("#### 📌 Semana Laboral (Lunes a Viernes)")
            c1, c2 = st.columns([3, 2])
            with c1:
                t_base_lv = st.multiselect("Turnos base L-V", options=CATALOGO_TURNOS, default=["03:00-11:00", "11:00-19:00", "19:00-27:00"], key=f"ms_{cargo}")
            with c2:
                txt_base_lv = st.text_input("Nuevo turno L-V (opcional)", key=f"tx_{cargo}")
            
            turnos_lv = list(t_base_lv) + ([txt_base_lv.strip()] if txt_base_lv.strip() else [])

            st.markdown("**Cantidad de personas requeridas por turno (L-V):**")
            if turnos_lv:
                cols_c = st.columns(len(turnos_lv))
                for idx_t, t_nom in enumerate(turnos_lv):
                    with cols_c[idx_t]:
                        cant = st.number_input(f"L-V: {t_nom}", min_value=1, max_value=20, value=1, key=f"n_lv_{cargo}_{t_nom}")
                        for d_nom in ["LUNES", "MARTES", "MIÉRCOLES", "JUEVES", "VIERNES"]:
                            matriz_demanda[cargo][d_nom][t_nom] = cant

            st.markdown("---")
            st.markdown("#### 🗓️ Fin de Semana (Sábado y Domingo)")
            col_sab, col_dom = st.columns(2)

            with col_sab:
                st.markdown("**SÁBADO**")
                t_sab = st.multiselect("Turnos Sábado", options=CATALOGO_TURNOS, default=["08:00-15:00 CAP"], key=f"ms_s_{cargo}")
                txt_sab = st.text_input("Nuevo turno Sábado", key=f"tx_s_{cargo}")
                list_sab = list(t_sab) + ([txt_sab.strip()] if txt_sab.strip() else [])
                for t_nom in list_sab:
                    cant_s = st.number_input(f"Sáb: {t_nom}", min_value=1, max_value=20, value=1, key=f"n_s_{cargo}_{t_nom}")
                    matriz_demanda[cargo]["SÁBADO"][t_nom] = cant_s

            with col_dom:
                st.markdown("**DOMINGO**")
                t_dom = st.multiselect("Turnos Domingo", options=CATALOGO_TURNOS, default=["08:00-15:00 CAP"], key=f"ms_d_{cargo}")
                txt_dom = st.text_input("Nuevo turno Domingo", key=f"tx_d_{cargo}")
                list_dom = list(t_dom) + ([txt_dom.strip()] if txt_dom.strip() else [])
                for t_nom in list_dom:
                    cant_d = st.number_input(f"Dom: {t_nom}", min_value=1, max_value=20, value=1, key=f"n_d_{cargo}_{t_nom}")
                    matriz_demanda[cargo]["DOMINGO"][t_nom] = cant_d

# ==========================================
# 3. GENERACIÓN DE LA MALLA
# ==========================================
def generar_malla_matriz(df_personal, semanas_count, fecha_base_date, reglas_demanda, historico):
    dias_totales = semanas_count * 7
    fecha_base = datetime.combine(fecha_base_date, datetime.min.time())

    columnas_fechas = []
    fechas_dt = []
    for i in range(dias_totales):
        f_actual = fecha_base + timedelta(days=i)
        fechas_dt.append(f_actual)
        nombre_dia = dias_semana_es[f_actual.weekday()]
        fecha_fmt = f_actual.strftime("%d-%b")
        columnas_fechas.append(f"{nombre_dia}\n{fecha_fmt}")

    programacion_matriz = {}
    for _, emp in df_personal.iterrows():
        cod = str(emp["CODIGO"]).strip()
        hist = historico.get(cod, {})
        
        programacion_matriz[cod] = {
            "CODIGO": cod,
            "NOMBRE": emp["NOMBRE"],
            "CARGO": emp["CARGO"],
            "ESTADO": str(emp.get("ESTADO", "ACTIVO")).strip().upper(),
            "SALIDA_PREVIA": hist.get("salida_domingo", None),
            "ULTIMO_TURNO": hist.get("ultimo_turno", "")
        }

    # Asignar Días Libres evitando repetir el mismo día de la semana pasada inmediatamente
    dias_libres_emp = {}
    for cod, datos in programacion_matriz.items():
        if datos["ESTADO"] == "VACACIONES":
            dias_libres_emp[cod] = list(range(dias_totales))
        else:
            hist = historico.get(cod, {})
            ult_l = hist.get("ultimo_dia_libre", -1)
            
            dias_libres_emp[cod] = []
            for s in range(semanas_count):
                opciones_libres = [d for d in range(7) if d != ult_l] # Evitar repetir el mismo día
                dia_elegido = random.choice(opciones_libres if opciones_libres else list(range(7)))
                dias_libres_emp[cod].append(dia_elegido + (s * 7))
                ult_l = dia_elegido

    for idx_dia, col_nombre in enumerate(columnas_fechas):
        fecha_col = fechas_dt[idx_dia]
        nombre_dia_semana = dias_semana_es[fecha_col.weekday()]
        
        for cargo, dem_dias in reglas_demanda.items():
            cupos_hoy = dem_dias.get(nombre_dia_semana, {})
            empleados_cargo = [
                cod for cod, d in programacion_matriz.items()
                if d["CARGO"] == cargo and idx_dia not in dias_libres_emp[cod]
            ]
            random.shuffle(empleados_cargo)
            
            turnos_a_cubrir = []
            for t_nom, req in cupos_hoy.items():
                turnos_a_cubrir.extend([t_nom] * req)

            for cod_emp in empleados_cargo:
                salida_ant = programacion_matriz[cod_emp]["SALIDA_PREVIA"]
                ult_turno_emp = programacion_matriz[cod_emp]["ULTIMO_TURNO"]
                
                # Intentar asignar un turno DIFERENTE al de la semana previa para favorecer la rotación
                turnos_ordenados = sorted(list(turnos_a_cubrir), key=lambda x: 1 if x == ult_turno_emp else 0)
                
                turno_asignado = None
                nueva_salida = None

                for t_candidato in turnos_ordenados:
                    parsed_h = extraer_horas(t_candidato)
                    if parsed_h:
                        h_ini, m_ini, h_fin, m_fin = parsed_h
                        entrada_dt = fecha_col.replace(hour=h_ini, minute=m_ini)

                        if calcular_descanso_suficiente(salida_ant, entrada_dt, min_horas=12):
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
                    programacion_matriz[cod_emp]["ULTIMO_TURNO"] = turno_asignado
                else:
                    programacion_matriz[cod_emp][col_nombre] = "L (Descanso)"
                    programacion_matriz[cod_emp]["SALIDA_PREVIA"] = None

        for cod_emp, d in programacion_matriz.items():
            if idx_dia in dias_libres_emp[cod_emp]:
                programacion_matriz[cod_emp][col_nombre] = "VACACIONES" if d["ESTADO"] == "VACACIONES" else "L"
                programacion_matriz[cod_emp]["SALIDA_PREVIA"] = None

    filas_finales = []
    for cod_emp, datos in programacion_matriz.items():
        filas_finales.append({k: v for k, v in datos.items() if k not in ["ESTADO", "SALIDA_PREVIA", "ULTIMO_TURNO"]})

    return pd.DataFrame(filas_finales)

# ==========================================
# 4. EXPORTACIÓN
# ==========================================
if df_empleados is not None:
    if st.button("⚡ Generar Malla Horaria Continuada y Rotativa"):
        df_resultado = generar_malla_matriz(
            df_empleados, semanas, fecha_inicio_date, matriz_demanda, historico_empleados
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
