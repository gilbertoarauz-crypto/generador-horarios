from datetime import datetime, timedelta
import io
import re
import random
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Generador de Horarios por Cargo", layout="wide")
st.title("📅 Generador de Horarios con Selector por Días y Semana Laboral")

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
# 2. SELECTOR VISUAL DE HORARIOS POR CARGO
# ==========================================
matriz_reglas = {}

if df_empleados is not None:
    st.subheader("2. Selección de Horarios por Cargo")
    st.info("Puedes usar una regla unificada para Lunes a Viernes o personalizar cada día de la semana.")

    cargos_unicos = df_empleados["CARGO"].dropna().unique().tolist()

    for cargo in cargos_unicos:
        with st.expander(f"🔹 Seleccionar turnos para: {cargo}", expanded=True):
            matriz_reglas[cargo] = {}
            
            # Opción para repetición rápida Lunes a Viernes
            repetir_lv = st.checkbox(
                "Usar mismo horario para toda la semana laboral (Lunes a Viernes)",
                value=True,
                key=f"chk_lv_{cargo}"
            )

            if repetir_lv:
                col_lv, col_sab, col_dom = st.columns(3)
                
                with col_lv:
                    st.markdown("**LUNES A VIERNES**")
                    turnos_lv = st.multiselect(
                        "Turnos Lunes-Viernes",
                        options=CATALOGO_TURNOS,
                        default=["08:00-17:00", "11:00-19:00"],
                        key=f"ms_lv_{cargo}"
                    )
                    txt_lv = st.text_input("Nuevo turno L-V (opcional)", key=f"tx_lv_{cargo}")
                    list_lv = list(turnos_lv) + ([txt_lv.strip()] if txt_lv.strip() else [])
                    
                    # Asignar a los 5 días de la semana laboral
                    for d in ["LUNES", "MARTES", "MIÉRCOLES", "JUEVES", "VIERNES"]:
                        matriz_reglas[cargo][d] = list_lv

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

            else:
                # Vista detallada día por día
                cols = st.columns(7)
                for idx_dia, dia_nombre in enumerate(dias_semana_es):
                    defecto = ["08:00-17:00", "11:00-19:00"] if idx_dia < 5 else ["08:00-15:00 CAP"]
                    
                    with cols[idx_dia]:
                        st.markdown(f"**{dia_nombre}**")
                        turnos_sel = st.multiselect(
                            "Turnos base",
                            options=CATALOGO_TURNOS,
                            default=[t for t in defecto if t in CATALOGO_TURNOS],
                            key=f"ms_{cargo}_{dia_nombre}"
                        )
                        turno_extra = st.text_input(
                            "Nuevo turno (opcional)",
                            key=f"tx_{cargo}_{dia_nombre}"
                        )
                        
                        turnos_totales = list(turnos_sel)
                        if turno_extra.strip():
                            turnos_totales.append(turno_extra.strip())
                            
                        matriz_reglas[cargo][dia_nombre] = turnos_totales

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
