from datetime import datetime, timedelta
import io
import random
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Generador de Horarios por Cargo", layout="wide")
st.title("📅 Generador de Horarios Personalizados por Cargo y Día")

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
# 1. CARGA DE PERSONAL (EXCEL / CSV)
# ==========================================
st.subheader("1. Cargar Lista de Personal")
uploaded_file = st.file_uploader(
    "Subir archivo Excel o CSV", type=["xlsx", "csv"]
)

df_empleados = None

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith(".csv"):
            df_empleados = pd.read_csv(uploaded_file)
        else:
            df_empleados = pd.read_excel(uploaded_file)

        df_empleados.columns = [
            str(c).upper().strip() for c in df_empleados.columns
        ]

        columnas_requeridas = {"CODIGO", "NOMBRE", "CARGO"}
        if not columnas_requeridas.issubset(set(df_empleados.columns)):
            st.error(
                f"El archivo debe contener las columnas: {columnas_requeridas}"
            )
            df_empleados = None
        else:
            if "ESTADO" not in df_empleados.columns:
                df_empleados["ESTADO"] = "ACTIVO"
            st.success(
                f"¡Se cargaron {len(df_empleados)} empleados exitosamente!"
            )
            st.dataframe(df_empleados, use_container_width=True)
    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")

# ==========================================
# 2. CONFIGURACIÓN DE HORARIOS POR CARGO Y DÍA
# ==========================================
matriz_reglas = {}

if df_empleados is not None:
    st.subheader("2. Matriz de Horarios por Cargo y Tipología de Día")
    st.info(
        "Define los horarios permitidos (separados por coma) para cada día de la semana según el cargo."
    )

    cargos_unicos = df_empleados["CARGO"].dropna().unique().tolist()

    for cargo in cargos_unicos:
        with st.expander(f"🔹 Configurar horarios para: {cargo}", expanded=True):
            matriz_reglas[cargo] = {}
            cols = st.columns(7)
            for idx_dia, dia_nombre in enumerate(dias_semana_es):
                val_defecto = (
                    "08:00-17:00, 11:00-19:00"
                    if idx_dia < 5
                    else "08:00-15:00 CAP"
                )
                matriz_reglas[cargo][dia_nombre] = cols[idx_dia].text_area(
                    label=dia_nombre,
                    value=val_defecto,
                    key=f"{cargo}_{dia_nombre}",
                    height=80,
                )

# ==========================================
# 3. LÓGICA DE GENERACIÓN
# ==========================================
def generar_malla_matriz(
    df_personal, semanas_count, fecha_base_date, reglas_cargos
):
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

        fila = {
            "CODIGO": codigo_emp,
            "NOMBRE": nombre_emp,
            "CARGO": cargo_emp,
        }

        # 1 día libre por semana
        dias_libres = [
            random.randint(0, 6) + (s * 7) for s in range(semanas_count)
        ]

        es_vacaciones = estado_emp == "VACACIONES"

        for idx, col_nombre in enumerate(columnas_fechas):
            fecha_col = fechas_dt[idx]
            nombre_dia_semana = dias_semana_es[fecha_col.weekday()]

            if es_vacaciones:
                fila[col_nombre] = "VACACIONES"
            elif idx in dias_libres:
                fila[col_nombre] = "L"
            else:
                # Obtener horarios permitidos para ese cargo y día específico
                turnos_txt = reglas_cargos.get(cargo_emp, {}).get(
                    nombre_dia_semana, ""
                )
                opciones_turnos = [
                    t.strip() for t in turnos_txt.split(",") if t.strip()
                ]

                if opciones_turnos:
                    fila[col_nombre] = random.choice(opciones_turnos)
                else:
                    fila[col_nombre] = "08:00-17:00"

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
            df_resultado.to_excel(
                writer, index=False, sheet_name="Programación"
            )

        st.download_button(
            label="📥 Descargar Excel Matriz",
            data=output.getvalue(),
            file_name=f"Horario_{semanas}_semanas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
