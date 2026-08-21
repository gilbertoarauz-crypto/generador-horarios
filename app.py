from datetime import datetime, timedelta
import io
import re
import random
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Generador de Horarios por Cargo", layout="wide")
st.title("📅 Generador de Horarios con Excepciones Diarias por Cargo")

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
# 1. CARGA DE PERSONAL
# ==========================================
st.subheader("1. Cargar Lista de Personal")
uploaded_file = st.file_uploader("Subir archivo Excel o CSV", type=["xlsx", "csv"])

df_empleados = None

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith(".csv"):
            df_empleados = pd.read_csv(uploaded_file)
        else:
            df_empleados = pd.read_excel(uploaded_file)

        df_empleados.columns = [str(c).upper().strip() for c in df_empleados.columns]

        columnas_requeridas = {"CODIGO", "NOMBRE", "CARGO"}
        if not columnas_requeridas.issubset(set(df_empleados.columns)):
            st.error(f"El archivo debe contener las columnas: {columnas_requeridas}")
            df_empleados = None
        else:
            if "ESTADO" not in df_empleados.columns:
                df_empleados["ESTADO"] = "ACTIVO"
            st.success(f"¡Se cargaron {len(df_empleados)} empleados exitosamente!")
            st.dataframe(df_empleados, use_container_width=True)
    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")

# ==========================================
# 2. SELECTOR VISUAL CON EXCEPCIONES DIARIAS
# ==========================================
matriz_reglas = {}

if df_empleados is not None:
    st.subheader("2. Selección de Horarios por Cargo")
    st.info("Configura el horario base para la semana laboral y marca los días que requieren horarios diferentes.")

    cargos_unicos = df_empleados["CARGO"].dropna().unique().tolist()

    for cargo in cargos_unicos:
        with st.expander(f"🔹 Seleccionar turnos para: {cargo}", expanded=True):
            matriz_reglas[cargo] = {}
            
            # 1. Horario General Base L-V
            st.markdown("#### 📌 Horario Base (Lunes a Viernes)")
            col_gen_1, col_gen_2 = st.columns([3, 2])
            with col_gen_1:
                turnos_base_lv = st.multiselect(
                    "Turnos generales Lunes-Viernes",
                    options=CATALOGO_TURNOS,
                    default=["08:00-17:00", "11:00-19:00"],
                    key=f"ms_base_{cargo}"
                )
            with col_gen_2:
                txt_base_lv = st.text_input("Nuevo turno L-V (opcional)", key=f"tx_base_{cargo}")
            
            list_base_lv = list(turnos_base_lv) + ([txt_base_lv.strip()] if txt_base_lv.strip() else [])

            # 2. Casillas para seleccionar excepciones en la semana laboral
            st.markdown("#### ⚡ Marcar si algún día de la semana laboral NO es igual al horario base:")
            cols_chk = st.columns(5)
            dias_excepcion = {}
            dias_laborales = ["LUNES", "MARTES", "MIÉRCOLES", "JUEVES", "VIERNES"]
            
            for idx_d, d_nom in enumerate(dias_laborales):
                with cols_chk[idx_d]:
                    dias_excepcion[d_nom] = st.checkbox(f"{d_nom} es diferente", key=f"chk_diff_{cargo}_{d_nom}")

            # Desplegar campos para los días marcados como diferentes
            for d_nom in dias_laborales:
                if dias_excepcion[d_nom]:
                    st.markdown(f"**Horarios específicos para {d_nom}:**")
                    c1, c2 = st.columns([3, 2])
                    with c1:
                        t_spec = st.multiselect(
                            f"Turnos para {d_nom}",
                            options=CATALOGO_TURNOS,
                            default=["06:00-15:00"],
                            key=f"ms_spec_{cargo}_{d_nom}"
                        )
                    with c2:
                        txt_spec = st.text_input(f"Nuevo turno {d_nom} (opcional)", key=f"tx_spec_{cargo}_{d_nom}")
                    matriz_reglas[cargo][d_nom] = list(t_spec) + ([txt_spec.strip()] if txt_spec.strip() else [])
                else:
                    matriz_reglas[cargo][d_nom] = list_base_lv

            # 3. Fin de Semana (Sábado y Domingo)
            st.markdown("#### 🗓️ Fin de Semana")
            col_sab, col_dom = st.columns(2)
            
            with col_sab:
                st.markdown("**SÁBADO**")
                turnos_sab = st.multiselect(
                    "Turnos Sábado",
                    options=CATALOGO_TURNOS,
                    default=["08:00-15:00 CAP"],
                    key=f"ms_sab_{cargo}"
                )
                txt_sab = st.text_input("Nuevo turno Sábado (opcional)", key=f"tx_sab_{cargo}")
                matriz_reglas[cargo]["SÁBADO"] = list(turnos_sab) + ([txt_sab.strip()] if txt_sab.strip() else [])

            with col_dom:
                st.markdown("**DOMINGO**")
                turnos_dom = st.multiselect(
                    "Turnos Domingo",
                    options=CATALOGO_TURNOS,
                    default=["08:00-15:00 CAP"],
                    key=f"ms_dom_{cargo}"
                )
                txt_dom = st.text_input("Nuevo turno Domingo (opcional)", key=f"tx_dom_{cargo}")
                matriz_reglas[cargo]["DOMINGO"] = list(turnos_dom) + ([txt_dom.strip()] if txt_dom.strip() else [])

# ==========================================
# 3. LÓGICA DE GENERACIÓN
# ==========================================
def generar_malla_matriz(df_personal, semanas_count, fecha_base_date, reglas_cargos):
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

    filas_horario = []

    for _, emp in df_personal.iterrows():
        codigo_emp = emp.get("CODIGO", "")
        nombre_emp = emp.get("NOMBRE", "")
        cargo_emp = emp.get("CARGO", "")
        estado_emp = str(emp.get("ESTADO", "ACTIVO")).strip().upper()

        fila = {"CODIGO": codigo_emp, "NOMBRE": nombre_emp, "CARGO": cargo_emp}

        dias_libres = [random.randint(0, 6) + (s * 7) for s in range(semanas_count)]
        es_vacaciones = estado_emp == "VACACIONES"

        salida_anterior_dt = None

        for idx, col_nombre in enumerate(columnas_fechas):
            fecha_col = fechas_dt[idx]
            nombre_dia_semana = dias_semana_es[fecha_col.weekday()]

            if es_vacaciones:
                fila[col_nombre] = "VACACIONES"
                salida_anterior_dt = None
            elif idx in dias_libres:
                fila[col_nombre] = "L"
                salida_anterior_dt = None
            else:
                opciones_turnos = reglas_cargos.get(cargo_emp, {}).get(nombre_dia_semana, [])
                opciones_turnos = list(opciones_turnos)

                random.shuffle(opciones_turnos)
                turno_seleccionado = None
                nueva_salida_dt = None

                for t_candidato in opciones_turnos:
                    parsed_h = extraer_horas(t_candidato)
                    if parsed_h:
                        h_ini, m_ini, h_fin, m_fin = parsed_h
                        entrada_dt = fecha_col.replace(hour=h_ini, minute=m_ini)

                        if calcular_descanso_suficiente(salida_anterior_dt, entrada_dt, min_horas=12):
                            turno_seleccionado = t_candidato
                            
                            if h_fin >= 24:
                                nueva_salida_dt = (fecha_col + timedelta(days=1)).replace(hour=h_fin - 24, minute=m_fin)
                            elif h_fin < h_ini:
                                nueva_salida_dt = (fecha_col + timedelta(days=1)).replace(hour=h_fin, minute=m_fin)
                            else:
                                nueva_salida_dt = fecha_col.replace(hour=h_fin, minute=m_fin)
                            break

                if turno_seleccionado is None:
                    if opciones_turnos:
                        fila[col_nombre] = "L (Descanso)"
                        salida_anterior_dt = None
                    else:
                        fila[col_nombre] = "08:00-17:00"
                        salida_anterior_dt = fecha_col.replace(hour=17, minute=0)
                else:
                    fila[col_nombre] = turno_seleccionado
                    salida_anterior_dt = nueva_salida_dt

        filas_horario.append(fila)

    return pd.DataFrame(filas_horario)

# ==========================================
# 4. PROCESAMIENTO Y EXPORTACIÓN
# ==========================================
if df_empleados is not None:
    if st.button("⚡ Generar Malla Horaria Personalizada"):
        df_resultado = generar_malla_matriz(
            df_empleados, semanas, fecha_inicio_date, matriz_reglas
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
