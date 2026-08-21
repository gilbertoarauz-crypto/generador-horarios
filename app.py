from datetime import datetime, timedelta
import io
import random
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Generador de Horarios", layout="wide")
st.title("📅 Generador Automático de Horarios Malla Matriz")

# ==========================================
# CONFIGURACIÓN LATERAL (SIDEBAR)
# ==========================================
st.sidebar.header("⚙️ Parámetros de Programación")
semanas = st.sidebar.slider("Semanas a generar", 1, 4, 2)
fecha_inicio_date = st.sidebar.date_input("Fecha de inicio", datetime.now())

# Entrada para la lista de turnos disponibles
st.sidebar.subheader("🕒 Turnos Disponibles por Semana")
st.sidebar.caption("Ingresa los turnos separados por comas o saltos de línea.")
turnos_input = st.sidebar.text_area(
    "Lista de Horarios/Turnos",
    value="07:00-16:00\n06:00-15:00\n11:00-19:00\n19:00-27:00\n03:00-11:00\n22:00-30:00\n08:00-17:00",
    height=150,
)

# Convertir el texto de turnos en una lista limpia
lista_turnos = [
    t.strip() for t in turnos_input.replace(",", "\n").split("\n") if t.strip()
]

# ==========================================
# 1. CARGA DE PERSONAL (EXCEL / CSV)
# ==========================================
st.subheader("1. Cargar Personal desde Archivo (Excel o CSV)")

uploaded_file = st.file_uploader(
    "Subir archivo con la lista de empleados", type=["xlsx", "csv"]
)

df_empleados = None

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith(".csv"):
            df_empleados = pd.read_csv(uploaded_file)
        else:
            df_empleados = pd.read_excel(uploaded_file)

        # Normalizar nombres de columnas a mayúsculas
        df_empleados.columns = [str(c).upper().strip() for c in df_empleados.columns]

        # Validar campos mínimos obligatorios
        columnas_requeridas = {"CODIGO", "NOMBRE", "CARGO"}
        if not columnas_requeridas.issubset(set(df_empleados.columns)):
            st.error(
                f"El archivo debe contener al menos las siguientes columnas: {columnas_requeridas}"
            )
            df_empleados = None
        else:
            # Si no existe la columna TAREA o ESTADO, se crea por defecto opcional
            if "ESTADO" not in df_empleados.columns and "TAREA" not in df_empleados.columns:
                df_empleados["ESTADO"] = "ROTATIVO"
            elif "TAREA" in df_empleados.columns and "ESTADO" not in df_empleados.columns:
                df_empleados["ESTADO"] = df_empleados["TAREA"]

            st.success(f"¡Se cargaron {len(df_empleados)} empleados con éxito!")
            st.dataframe(df_empleados, use_container_width=True)

    except Exception as e:
        st.error(f"Error al leer el archivo: {e}")

# ==========================================
# 2. GENERADOR DE MALLA HORIZONTAL
# ==========================================
def generar_malla_horizontal(df_personal, semanas_count, fecha_base_date, turnos_rotacion):
    dias_totales = semanas_count * 7
    fecha_base = datetime.combine(fecha_base_date, datetime.min.time())

    dias_semana_es = [
        "LUNES",
        "MARTES",
        "MIÉRCOLES",
        "JUEVES",
        "VIERNES",
        "SÁBADO",
        "DOMINGO",
    ]

    # Crear encabezados con estructura de Fecha y Día (Ej: LUNES 17-ago)
    columnas_fechas = []
    for i in range(dias_totales):
        f_actual = fecha_base + timedelta(days=i)
        nombre_dia = dias_semana_es[f_actual.weekday()]
        fecha_fmt = f_actual.strftime("%d-%b")
        columnas_fechas.append(f"{nombre_dia}\n{fecha_fmt}")

    filas_horario = []

    for _, emp in df_personal.iterrows():
        codigo_emp = emp.get("CODIGO", "")
        nombre_emp = emp.get("NOMBRE", "")
        cargo_emp = emp.get("CARGO", "")
        estado_emp = str(emp.get("ESTADO", "ROTATIVO")).strip().upper()

        fila = {
            "CODIGO": codigo_emp,
            "NOMBRE": nombre_emp,
            "CARGO": cargo_emp,
        }

        # Días libres (1 por cada 7 días / semana)
        dias_libres = [
            random.randint(0, 6) + (s * 7) for s in range(semanas_count)
        ]

        es_vacaciones = estado_emp == "VACACIONES"

        # Asignar turno aleatorio o específico de la lista para cada día
        turno_asignado_semana = random.choice(turnos_rotacion) if turnos_rotacion else "08:00-17:00"

        for idx, col_nombre in enumerate(columnas_fechas):
            # Cambiar turno cada semana (cada 7 días) si es rotativo
            if idx % 7 == 0 and turnos_rotacion:
                turno_asignado_semana = random.choice(turnos_rotacion)

            if es_vacaciones:
                fila[col_nombre] = "VACACIONES"
            elif idx in dias_libres:
                fila[col_nombre] = "L"
            elif estado_emp not in ["ROTATIVO", "NAN", "NONE", ""]:
                # Si se especificó un horario o estado fijo directo en el archivo subido
                fila[col_nombre] = estado_emp
            else:
                fila[col_nombre] = turno_asignado_semana

        filas_horario.append(fila)

    return pd.DataFrame(filas_horario)

# ==========================================
# 3. PROCESAMIENTO Y DESCARGA
# ==========================================
if df_empleados is not None:
    if st.button("⚡ Generar Malla Horaria"):
        if not lista_turnos:
            st.warning("Añade al menos un turno en el menú lateral.")
        else:
            df_resultado = generar_malla_horizontal(
                df_empleados, semanas, fecha_inicio_date, lista_turnos
            )

            st.subheader("2. Malla Horaria Generada")
            st.dataframe(df_resultado, use_container_width=True)

            # Exportación a Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df_resultado.to_excel(writer, index=False, sheet_name="Programación")

            st.download_button(
                label="📥 Descargar Excel Matriz",
                data=output.getvalue(),
                file_name=f"Horario_{semanas}_semanas.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
