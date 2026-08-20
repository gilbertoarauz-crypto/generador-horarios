from datetime import datetime, timedelta
import random
import io
import pandas as pd
import streamlit as st

st.title("📅 Generador Automático de Horarios")

# Configuración interactiva en la web
semanas = st.sidebar.slider("Semanas a generar", 1, 4, 1)
fecha_inicio = st.sidebar.date_input("Fecha de inicio", datetime.now())

# Formulario para agregar empleados
st.subheader("1. Configuración de Personal")
if "empleados" not in st.session_state:
    st.session_state.empleados = []

with st.form("form_empleado"):
    col1, col2, col3, col4 = st.columns(4)
    codigo = col1.text_input("Código")
    nombre = col2.text_input("Nombre")
    cargo = col3.text_input("Cargo")
    tarea = col4.text_input("Tarea")
    agregar = st.form_submit_button("Agregar Empleado")

    if agregar and codigo and nombre:
        st.session_state.empleados.append(
            {
                "codigo": codigo,
                "nombre": nombre,
                "cargo": cargo,
                "tarea": tarea,
            }
        )
        st.success(f"Empleado {nombre} agregado.")

# Mostrar lista de empleados actual
if st.session_state.empleados:
    st.write("Empleados registrados:", pd.DataFrame(st.session_state.empleados))


# Lógica de generación de horarios
def generar_malla(empleados_list, semanas_count, fecha_base):
    turnos = {
        "Mañana": {"inicio": 7, "fin": 15},
        "Tarde": {"inicio": 15, "fin": 23},
        "Noche": {"inicio": 23, "fin": 7},
    }
    programacion = []
    dias_totales = semanas_count * 7

    for emp in empleados_list:
        dias_libres = [
            random.randint(0, 6) + (s * 7) for s in range(semanas_count)
        ]
        ultimo_fin = None

        for dia_idx in range(dias_totales):
            fecha_actual = fecha_base + timedelta(days=dia_idx)

            if dia_idx in dias_libres:
                programacion.append(
                    {
                        "Código": emp["codigo"],
                        "Nombre": emp["nombre"],
                        "Cargo": emp["cargo"],
                        "Tarea": "N/A",
                        "Fecha": fecha_actual.strftime("%Y-%m-%d"),
                        "Turno": "LIBRE",
                        "Entrada": "OFF",
                        "Salida": "OFF",
                    }
                )
                ultimo_fin = None
                continue

            turnos_validos = []
            for n_turno, horas in turnos.items():
                h_entrada = fecha_actual.replace(
                    hour=horas["inicio"], minute=0, second=0
                )
                if (
                    ultimo_fin is None
                    or (h_entrada - ultimo_fin).total_seconds() / 3600 >= 12
                ):
                    turnos_validos.append((n_turno, horas))

            n_turno, horas = (
                random.choice(turnos_validos)
                if turnos_validos
                else ("Mañana", turnos["Mañana"])
            )

            h_entrada = fecha_actual.replace(hour=horas["inicio"], minute=0)
            h_salida = (
                (fecha_actual + timedelta(days=1)).replace(
                    hour=horas["fin"], minute=0
                )
                if horas["fin"] < horas["inicio"]
                else fecha_actual.replace(hour=horas["fin"], minute=0)
            )

            ultimo_fin = h_salida
            programacion.append(
                {
                    "Código": emp["codigo"],
                    "Nombre": emp["nombre"],
                    "Cargo": emp["cargo"],
                    "Tarea": emp["tarea"],
                    "Fecha": fecha_actual.strftime("%Y-%m-%d"),
                    "Turno": n_turno,
                    "Entrada": f"{horas['inicio']:02d}:00",
                    "Salida": f"{horas['fin']:02d}:00",
                }
            )

    return pd.DataFrame(programacion)


# Botón para procesar y descargar
if (
    st.button("⚡ Generar Horarios")
    and st.session_state.empleados
):
    df_resultado = generar_malla(
        st.session_state.empleados, semanas, fecha_inicio
    )
    st.subheader("2. Horario Generado")
    st.dataframe(df_resultado)

    # Convertir DataFrame a Excel para descarga en web
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_resultado.to_excel(writer, index=False)

    st.download_button(
        label="📥 Descargar Excel",
        data=output.getvalue(),
        file_name="horario_generado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
