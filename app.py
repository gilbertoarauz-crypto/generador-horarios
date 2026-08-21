from datetime import datetime, timedelta
import io
import re
import random
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Generador de Horarios por Demanda", layout="wide")
st.title("📅 Generador de Horarios con Cupos por Turno y Cargo")

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
# 2. CONFIGURACIÓN DE CUPOS POR TURNO Y CARGO
# ==========================================
matriz_cupos = {}

if df_empleados is not None:
    st.subheader("2. Requerimientos de Personal (Cupos por Turno)")
    st.info("Selecciona los turnos para el cargo e indica la cantidad de personas necesarias por turno.")

    cargos_unicos = df_empleados["CARGO"].dropna().unique().tolist()

    for cargo in cargos_unicos:
        with st.expander(f"🔹 Configurar requirimientos para: {cargo}", expanded=True):
            matriz_cupos[cargo] = {}
            
            # Selector de turnos base para el cargo
            turnos_sel = st.multiselect(
                f"Turnos habilitados para {cargo}",
                options=CATALOGO_TURNOS,
                default=["03:00-11:00", "11:00-19:00", "19:00-27:00"],
                key=f"ms_turnos_{cargo}"
            )
            
            if turnos_sel:
                st.markdown("**Cantidad de personas requeridas por turno (Lunes a Domingo):**")
                cols_cupos = st.columns(len(turnos_sel))
                
                for idx_t, turno in enumerate(turnos_sel):
                    with cols_cupos[idx_t]:
                        # Definir por defecto 1 o 2 personas según el turno
                        cant = st.number_input(
                            label=f"Turno {turno}",
                            min_value=1,
                            max_value=20,
                            value=1,
                            key=f"num_{cargo}_{turno}"
                        )
                        matriz_cupos[cargo][turno] = cant

# ==========================================
# 3. LÓGICA DE GENERACIÓN CON CUPOS Y RESTRICCIONES
# ==========================================
def generar_malla_matriz(df_personal, semanas_count, fecha_base_date, reglas_cupos):
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

    # Estructura base para guardar resultados
    programacion_matriz = {emp["CODIGO"]: {
        "CODIGO": emp["CODIGO"],
        "NOMBRE": emp["NOMBRE"],
        "CARGO": emp["CARGO"],
        "ESTADO": str(emp.get("ESTADO", "ACTIVO")).strip().upper(),
        "SALIDA_PREVIA": None
    } for _, emp in df_personal.iterrows()}

    # Asignar Días Libres (1 por semana) y Vacaciones
    dias_libres_emp = {}
    for cod, datos in programacion_matriz.items():
        if datos["ESTADO"] == "VACACIONES":
            dias_libres_emp[cod] = list(range(dias_totales))
        else:
            dias_libres_emp[cod] = [random.randint(0, 6) + (s * 7) for s in range(semanas_count)]

    # Recorrido día por día (cobertura equitativa de turnos)
    for idx_dia, col_nombre in enumerate(columnas_fechas):
        fecha_col = fechas_dt[idx_dia]
        
        # Agrupar empleados disponibles por cargo para este día
        cargos = list(reglas_cupos.keys())
        
        for cargo in cargos:
            cupos_cargo = reglas_cupos[cargo] # Dict {"03:00-11:00": 2, ...}
            
            # Buscar empleados de este cargo activos y que no tengan día libre hoy
            empleados_cargo = [
                cod for cod, d in programacion_matriz.items()
                if d["CARGO"] == cargo and idx_dia not in dias_libres_emp[cod]
            ]
            
            random.shuffle(empleados_cargo)
            
            # Crear lista extendida de turnos a cubrir según la demanda (ej: ["03:00-11:00", "03:00-11:00", "11:00-19:00", ...])
            turnos_a_cubrir = []
            for t_nom, requeridos in cupos_cargo.items():
                turnos_a_cubrir.extend([t_nom] * requeridos)

            # Asignar turnos a los empleados disponibles
            for cod_emp in empleados_cargo:
                salida_ant = programacion_matriz[cod_emp]["SALIDA_PREVIA"]
                turno_asignado = None
                nueva_salida = None

                # Intentar asignar de la lista de requeridos
                for t_candidato in list(turnos_a_cubrir):
                    parsed_h = extraer_horas(t_candidato)
                    if parsed_h:
                        h_ini, m_ini, h_fin, m_fin = parsed_h
                        entrada_dt = fecha_col.replace(hour=h_ini, minute=m_ini)

                        if calcular_descanso_suficiente(salida_ant, entrada_dt, min_horas=12):
                            turno_asignado = t_candidato
                            turnos_a_cubrir.remove(t_candidato) # Consumir el cupo
                            
                            if h_fin >= 24:
                                nueva_salida = (fecha_col + timedelta(days=1)).replace(hour=h_fin - 24, minute=m_fin)
                            elif h_fin < h_ini:
                                nueva_salida = (fecha_col + timedelta(days=1)).replace(hour=h_fin, minute=m_fin)
                            else:
                                nueva_salida = fecha_col.replace(hour=h_fin, minute=m_fin)
                            break
                
                # Si no había requerimiento pendiente o no cumplía descanso, dar turno alternativo válido
                if turno_asignado is None:
                    for t_alt in cupos_cargo.keys():
                        parsed_h = extraer_horas(t_alt)
                        if parsed_h:
                            h_ini, m_ini, h_fin, m_fin = parsed_h
                            entrada_dt = fecha_col.replace(hour=h_ini, minute=m_ini)
                            if calcular_descanso_suficiente(salida_ant, entrada_dt, min_horas=12):
                                turno_asignado = t_alt
                                if h_fin >= 24:
                                    nueva_salida = (fecha_col + timedelta(days=1)).replace(hour=h_fin - 24, minute=m_fin)
                                else:
                                    nueva_salida = fecha_col.replace(hour=h_fin, minute=m_fin)
                                break

                if turno_asignado:
                    programacion_matriz[cod_emp][col_nombre] = turno_asignado
                    programacion_matriz[cod_emp]["SALIDA_PREVIA"] = nueva_salida
                else:
                    programacion_matriz[cod_emp][col_nombre] = "L (Descanso)"
                    programacion_matriz[cod_emp]["SALIDA_PREVIA"] = None

        # Marcar días libres y vacaciones en el resultado
        for cod_emp, d in programacion_matriz.items():
            if idx_dia in dias_libres_emp[cod_emp]:
                if d["ESTADO"] == "VACACIONES":
                    programacion_matriz[cod_emp][col_nombre] = "VACACIONES"
                else:
                    programacion_matriz[cod_emp][col_nombre] = "L"
                programacion_matriz[cod_emp]["SALIDA_PREVIA"] = None

    # Formatear la matriz final limpiando columnas auxiliares
    filas_finales = []
    for cod_emp, datos in programacion_matriz.items():
        d_limpio = {k: v for k, v in datos.items() if k not in ["ESTADO", "SALIDA_PREVIA"]}
        filas_finales.append(d_limpio)

    return pd.DataFrame(filas_finales)

# ==========================================
# 4. PROCESAMIENTO Y EXPORTACIÓN
# ==========================================
if df_empleados is not None:
    if st.button("⚡ Generar Malla Horaria con Demanda de Personal"):
        df_resultado = generar_malla_matriz(
            df_empleados, semanas, fecha_inicio_date, matriz_cupos
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
