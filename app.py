from datetime import datetime, timedelta
import io
import random
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Generador de Horarios", layout="wide")
st.title("📅 Generador Automático de Horarios (Malla Horizontal)")

# Configuración interactiva
semanas = st.sidebar.slider("Semanas a generar", 1, 4, 2)
fecha_inicio_date = st.sidebar.date_input("Fecha de inicio", datetime.now())

# Formulario para agregar empleados
st.subheader("1. Configuración de Personal")
if "empleados" not in st.session_state:
    st.session_state.empleados = []

with st.form("form_empleado"):
    col1, col2, col3, col4 = st.columns(4)
    codigo = col1.text_input("CÓDIGO")
    nombre = col2.text_input("NOMBRE")
    cargo = col3.text_input("CARGO")
    tarea = col4.text_input("TAREA / ESTADO", value="07:00-16:00")
    agregar = st.form_submit_button("Agregar Empleado")

    if agregar and codigo and nombre:
        st.session_state.empleados.append(
            {
                "CODIGO": codigo,
                "NOMBRE": nombre,
                "CARGO": cargo,
                "TAREA": tarea,
            }
        )
        st.success(f"Empleado {nombre} agregado.")

# Mostrar lista de empleados
if st.session_state.empleados:
    st.write("Empleados registrados:", pd.DataFrame(st.session_state.empleados))


# Generador en formato matriz (horizontal)
def generar_malla_horizontal(empleados_list, semanas_count, fecha_base_date):
    dias_totales = semanas_count * 7
    fecha_base = datetime.combine(fecha_base_date, datetime.min.time())

    # Nombres de días en español
    dias_semana_es = [
        "LUNES",
        "MARTES",
        "MIÉRCOLES",
        "JUEVES",
        "VIERNES",
        "SÁBADO",
        "DOMINGO",
    ]

    # Crear encabezados con estructura de Fecha y Día
    columnas_fechas = []
    fechas_dt = []
    for i in range(dias_totales):
        f_actual = fecha_base + timedelta(days=i)
        fechas_dt.append(f_actual)
        nombre_dia = dias_semana_es[f_actual.weekday()]
        fecha_fmt = f_actual.strftime("%d-%b")
        columnas_fechas.append(f"{nombre_dia}\n{fecha_fmt}")

    filas_horario = []

    for emp in empleados_list:
        fila = {
            "CÓDIGO": emp["CODIGO"],
            "NOMBRE": emp["NOMBRE"],
            "CARGO": emp["CARGO"],
        }

        # Días libres (1 por semana)
        dias_libres = [
            random.randint(0, 6) + (s * 7) for s in range(semanas_count)
        ]

        # Si el estado del empleado es VACACIONES, se asigna a toda la matriz
        es_vacaciones = emp["TAREA"].strip().upper() == "VACACIONES"

        for idx, col_nombre in enumerate(columnas_fechas):
            if es_vacaciones:
                fila[col_nombre] = "VACACIONES"
            elif idx in dias_libres:
                fila[col_nombre] = "L"
            else:
                fila[col_nombre] = emp["TAREA"]

        filas_horario.append(fila)

    return pd.DataFrame(filas_horario)


# Generación y visualización
if (
    st.button("⚡ Generar Horarios")
    and st.session_state.empleados
):
    df_matriz = generar_malla_horizontal(
        st.session_state.empleados, semanas, fecha_inicio_date
    )

    st.subheader("2. Malla Horaria")
    st.dataframe(df_matriz, use_container_width=True)

    # Exportación a Excel con formato horizontal
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_matriz.to_excel(writer, index=False, sheet_name="Programación")

    st.download_button(
        label="📥 Descargar Excel Matriz",
        data=output.getvalue(),
        file_name="horario_matriz.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
