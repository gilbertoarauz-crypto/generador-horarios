from datetime import datetime, time, timedelta
import io
import re
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Generador de Horarios Pro", layout="wide")
st.title("📅 Generador de Horarios & Control de Tareas Operativas")

# ==========================================
# CONSTANTES Y CATÁLOGOS
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
    "23:00-31:00",
]

TURNOS_DEFAULT_POR_CARGO = {
    "ANALISTA DE OPERACIONES": {
        "habil": [
            "03:00-11:00",
            "11:00-19:00",
            "06:00-15:00",
            "19:00-27:00",
        ],
        "sabado": [
            "03:00-11:00",
            "11:00-19:00",
            "19:00-27:00",
            "06:00-15:00",
        ],
        "domingo": ["03:00-11:00", "11:00-19:00", "19:00-27:00"],
    },
    "AUXILIAR DE OPERACIONES": {
        "habil": [
            "03:00-11:00",
            "03:00-11:00",
            "11:00-19:00",
            "11:00-19:00",
            "19:00-27:00",
            "20:00-28:00",
        ],
        "sabado": [
            "03:00-11:00",
            "03:00-11:00",
            "11:00-19:00",
            "11:00-19:00",
            "19:00-27:00",
            "22:00-30:00",
        ],
        "domingo": [
            "03:00-11:00",
            "03:00-11:00",
            "11:00-19:00",
            "11:00-19:00",
            "19:00-27:00",
            "20:00-28:00",
        ],
    },
    "OPERADOR LÍDER": {
        "habil": ["20:00-28:00", "20:00-28:00"],
        "sabado": ["20:00-28:00"],
        "domingo": ["20:00-28:00", "20:00-28:00"],
    },
    "TÉCNICO DE OPERACIONES": {
        "habil": [
            "03:00-11:00",
            "11:00-19:00",
            "19:00-27:00",
            "06:00-15:00",
        ],
        "sabado": [
            "03:00-11:00",
            "11:00-19:00",
            "19:00-27:00",
            "06:00-15:00",
        ],
        "domingo": [
            "03:00-11:00",
            "11:00-19:00",
            "19:00-27:00",
            "06:00-15:00",
        ],
    },
    "AUXILIAR DE ALISTAMIENTO": {
        "habil": [
            "08:00-17:00",
            "22:00-30:00",
            "20:00-28:00",
            "20:00-28:00",
        ],
        "sabado": ["08:00-17:00", "20:00-28:00", "22:00-30:00"],
        "domingo": [
            "08:00-17:00",
            "20:00-28:00",
            "22:00-30:00",
            "20:00-28:00",
        ],
    },
}

PATRONES_ANALISTAS_5 = [
    [
        "06:00-15:00",
        "06:00-15:00",
        "06:00-15:00",
        "L",
        "03:00-11:00",
        "03:00-11:00",
        "08:00-17:00",
        "06:00-15:00",
        "06:00-15:00",
        "L",
        "03:00-11:00",
        "03:00-11:00",
        "03:00-11:00",
        "06:00-15:00",
    ],
    [
        "08:00-17:00",
        "08:00-17:00",
        "08:00-17:00",
        "08:00-17:00",
        "08:00-17:00",
        "11:00-19:00",
        "L",
        "08:00-17:00",
        "08:00-17:00",
        "06:00-15:00",
        "06:00-15:00",
        "06:00-15:00",
        "L",
        "03:00-11:00",
    ],
    [
        "11:00-19:00",
        "11:00-19:00",
        "11:00-19:00",
        "11:00-19:00",
        "11:00-19:00",
        "L",
        "03:00-11:00",
        "03:00-11:00",
        "03:00-11:00",
        "03:00-11:00",
        "03:00-11:00",
        "L",
        "19:00-27:00",
        "19:00-27:00",
    ],
    [
        "03:00-11:00",
        "03:00-11:00",
        "03:00-11:00",
        "03:00-11:00",
        "L",
        "19:00-27:00",
        "19:00-27:00",
        "19:00-27:00",
        "19:00-27:00",
        "19:00-27:00",
        "19:00-27:00",
        "19:00-27:00",
        "L",
        "11:00-19:00",
    ],
    [
        "19:00-27:00",
        "19:00-27:00",
        "19:00-27:00",
        "19:00-27:00",
        "19:00-27:00",
        "L",
        "11:00-19:00",
        "11:00-19:00",
        "11:00-19:00",
        "11:00-19:00",
        "11:00-19:00",
        "11:00-19:00",
        "11:00-19:00",
        "L",
    ],
]

PATRONES_ANALISTAS_4 = [
    [
        "06:00-15:00",
        "06:00-15:00",
        "06:00-15:00",
        "L",
        "03:00-11:00",
        "03:00-11:00",
        "06:00-15:00",
        "06:00-15:00",
        "06:00-15:00",
        "06:00-15:00",
        "L",
        "03:00-11:00",
        "03:00-11:00",
        "03:00-11:00",
    ],
    [
        "vacaciones",
        "vacaciones",
        "vacaciones",
        "vacaciones",
        "vacaciones",
        "vacaciones",
        "vacaciones",
        "vacaciones",
        "vacaciones",
        "vacaciones",
        "vacaciones",
        "vacaciones",
        "vacaciones",
        "vacaciones",
    ],
    [
        "11:00-19:00",
        "11:00-19:00",
        "11:00-19:00",
        "11:00-19:00",
        "11:00-19:00",
        "L",
        "03:00-11:00",
        "03:00-11:00",
        "03:00-11:00",
        "03:00-11:00",
        "03:00-11:00",
        "L",
        "19:00-27:00",
        "19:00-27:00",
    ],
    [
        "03:00-11:00",
        "03:00-11:00",
        "03:00-11:00",
        "03:00-11:00",
        "L",
        "19:00-27:00",
        "19:00-27:00",
        "19:00-27:00",
        "19:00-27:00",
        "19:00-27:00",
        "19:00-27:00",
        "19:00-27:00",
        "L",
        "11:00-19:00",
    ],
    [
        "19:00-27:00",
        "19:00-27:00",
        "19:00-27:00",
        "19:00-27:00",
        "19:00-27:00",
        "L",
        "11:00-19:00",
        "11:00-19:00",
        "11:00-19:00",
        "11:00-19:00",
        "11:00-19:00",
        "11:00-19:00",
        "11:00-19:00",
        "L",
    ],
]

DIAS_SEMANA_ES = [
    "LUNES",
    "MARTES",
    "MIÉRCOLES",
    "JUEVES",
    "VIERNES",
    "SÁBADO",
    "DOMINGO",
]
NO_WORKING_TERMS = {
    "L",
    "AO",
    "VACACIONES",
    "LICENCIA",
    "INCAPACIDAD",
    "PERMISO",
    "SUSPEND",
}

# ==========================================
# SIDEBAR - CONFIGURACIÓN GENERAL
# ==========================================
st.sidebar.header("⚙️ Parámetros de Programación")
semanas = st.sidebar.slider("Semanas a generar", 1, 4, 2)
fecha_inicio_date = st.sidebar.date_input("Fecha de inicio", datetime.now())

st.sidebar.markdown("---")
st.sidebar.header("🏖️ Configuración de Descansos")
libres_por_semana_base = st.sidebar.number_input(
    "Días libres por semana (base)", min_value=1, max_value=3, value=1
)
tiene_festivo = st.sidebar.checkbox(
    "¿Hay día festivo en el periodo?", value=False
)

fechas_festivas_sel = []
if tiene_festivo:
    dias_totales_temp = semanas * 7
    fechas_posibles = [
        (fecha_inicio_date + timedelta(days=i))
        for i in range(dias_totales_temp)
    ]
    fechas_festivas_sel = st.sidebar.multiselect(
        "Días festivos (otorgan día libre compensatorio adicional en su semana):",
        options=fechas_posibles,
        format_func=lambda x: x.strftime("%d-%b-%Y"),
    )

# ==========================================
# FUNCIONES ROBUSTAS DE TIEMPO Y PARSEO
# ==========================================
def extraer_horas(texto_turno: str):
    if not texto_turno or str(texto_turno).strip().upper() in NO_WORKING_TERMS:
        return None
    match = re.search(
        r"(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})", str(texto_turno)
    )
    if match:
        return tuple(map(int, match.groups()))
    return None


def calcular_datetimes_turno(fecha_base_dt: datetime, texto_turno: str):
    parsed = extraer_horas(texto_turno)
    if not parsed:
        return None, None
    h_i, m_i, h_f, m_f = parsed
    dt_inicio = fecha_base_dt + timedelta(hours=h_i, minutes=m_i)
    dt_fin = fecha_base_dt + timedelta(hours=h_f, minutes=m_f)
    return dt_inicio, dt_fin


def clasificar_franja(texto_turno: str) -> str:
    parsed = extraer_horas(texto_turno)
    if not parsed:
        return "MAÑANA"
    h_ini = parsed[0]
    if 0 <= h_ini < 11:
        return "MAÑANA"
    elif 11 <= h_ini < 18:
        return "TARDE"
    return "NOCHE"


def obtener_siguiente_franja_permitida(
    franja_actual: str, viene_de_descanso: bool = False
) -> list:
    """
    Si viene de descanso (L) después de un turno NOCTURNO, prioriza TARDE
    (Noche -> Libre -> Tarde).
    """
    if viene_de_descanso and franja_actual == "NOCHE":
        return ["TARDE", "MAÑANA", "NOCHE"]
    if viene_de_descanso and franja_actual == "TARDE":
        return ["MAÑANA", "TARDE", "NOCHE"]

    mapeo = {
        "NOCHE": ["NOCHE", "TARDE"],
        "TARDE": ["TARDE", "MAÑANA"],
        "MAÑANA": ["MAÑANA", "TARDE", "NOCHE"],
    }
    return mapeo.get(franja_actual, ["MAÑANA", "TARDE", "NOCHE"])


def calcular_descanso_suficiente(
    salida_previa_dt: datetime,
    entrada_actual_dt: datetime,
    min_horas: float = 12.0,
) -> bool:
    if salida_previa_dt is None or entrada_actual_dt is None:
        return True
    diferencia_horas = (
        entrada_actual_dt - salida_previa_dt
    ).total_seconds() / 3600.0
    return diferencia_horas >= min_horas


def prioridad_por_restriccion_horaria(cod: str, matriz: dict):
    """
    Ordena primero a quienes tienen la hora de salida previa más tardía
    (mayor restricción de 12h/24h), y luego a quienes vienen de L tras noche
    para que capturen su turno de TARDE.
    """
    d = matriz[cod]
    sal = d.get("SALIDA_PREVIA_DT")
    if sal is not None:
        return (0, -sal.timestamp())
    if d.get("VIENE_DE_DESCANSO"):
        franja_ant = d.get("FRANJA_PREVIA_DESCANSO")
        peso_franja = {"NOCHE": 0, "TARDE": 1, "MAÑANA": 2}.get(franja_ant, 3)
        return (1, peso_franja)
    return (2, 0)


def parsear_fecha_incidencia(val_fecha, anio_referencia: int):
    if pd.isna(val_fecha) or str(val_fecha).strip() == "":
        return None
    try:
        dt = pd.to_datetime(val_fecha, dayfirst=True)
        if dt.year == 1970 or dt.year != anio_referencia:
            dt = dt.replace(year=anio_referencia)
        return dt.date()
    except Exception:
        return None


def encontrar_posicion_secuencia_analista(
    historial_turnos_ultimos, matriz_patrones
):
    if not historial_turnos_ultimos:
        return None, 0

    len_hist = len(historial_turnos_ultimos)
    for idx_patron, patron in enumerate(matriz_patrones):
        for pos_fin in range(14):
            coincide = True
            for i in range(min(len_hist, 14)):
                pos_eval = (pos_fin - i) % 14
                turno_hist = historial_turnos_ultimos[-(i + 1)].strip().upper()
                turno_patron = patron[pos_eval].strip().upper()

                if turno_hist != turno_patron:
                    coincide = False
                    break
            if coincide:
                return idx_patron, (pos_fin + 1) % 14

    return None, 0


# ==========================================
# 1. CARGA DE ARCHIVOS
# ==========================================
st.subheader("1. Carga Inicial de Datos de Entrada")
col_f1, col_f2, col_f3 = st.columns(3)
df_empleados, df_semana_anterior, df_tareas_req = None, None, None

with col_f1:
    uploaded_file = st.file_uploader(
        "1. Lista de Personal (Excel/CSV)",
        type=["xlsx", "csv"],
        key="file_personal",
    )

with col_f2:
    uploaded_prev_file = st.file_uploader(
        "2. Malla Semana Anterior (Opcional)",
        type=["xlsx", "csv"],
        key="file_prev",
    )

with col_f3:
    uploaded_tareas_file = st.file_uploader(
        "3. Matriz Tareas por Cargo (Opcional)",
        type=["xlsx", "csv"],
        key="file_tareas_ini",
    )

if uploaded_tareas_file is not None:
    try:
        df_tareas_req = (
            pd.read_csv(uploaded_tareas_file)
            if uploaded_tareas_file.name.endswith(".csv")
            else pd.read_excel(uploaded_tareas_file)
        )
        df_tareas_req.columns = [
            str(c).strip().upper() for c in df_tareas_req.columns
        ]
        if "CARGO" in df_tareas_req.columns:
            df_tareas_req["CARGO"] = (
                df_tareas_req["CARGO"].astype(str).str.strip().str.upper()
            )
        st.success("✅ Matriz de Tareas cargada.")
    except Exception as e:
        st.error(f"Error en archivo de tareas: {e}")
else:
    filas_default = []
    for c, t_dict in TURNOS_DEFAULT_POR_CARGO.items():
        max_len = max(
            len(t_dict["habil"]),
            len(t_dict["sabado"]),
            len(t_dict["domingo"]),
        )
        for i in range(max_len):
            filas_default.append(
                {
                    "CARGO": c.upper(),
                    "HABIL": t_dict["habil"][i]
                    if i < len(t_dict["habil"])
                    else "",
                    "SABADO": t_dict["sabado"][i]
                    if i < len(t_dict["sabado"])
                    else "",
                    "DOMINGO": t_dict["domingo"][i]
                    if i < len(t_dict["domingo"])
                    else "",
                }
            )
    df_tareas_req = pd.DataFrame(filas_default)

with st.expander(
    "👁️ Ver / Editar Tareas Requeridas por Cargo", expanded=False
):
    st.dataframe(df_tareas_req, use_container_width=True)

if uploaded_file is not None:
    try:
        df_empleados = (
            pd.read_csv(uploaded_file)
            if uploaded_file.name.endswith(".csv")
            else pd.read_excel(uploaded_file)
        )
        df_empleados.columns = [
            str(c).upper().strip() for c in df_empleados.columns
        ]

        if not {"CODIGO", "NOMBRE", "CARGO"}.issubset(
            set(df_empleados.columns)
        ):
            st.error(
                "El archivo debe incluir las columnas: CODIGO, NOMBRE, CARGO"
            )
            df_empleados = None
        else:
            df_empleados["CARGO"] = (
                df_empleados["CARGO"].astype(str).str.strip().str.upper()
            )
            if "ESTADO" not in df_empleados.columns:
                df_empleados["ESTADO"] = "ACTIVO"

            col_inc = [c for c in df_empleados.columns if "INCIDENCIA" in c]
            col_f_ini = [
                c
                for c in df_empleados.columns
                if "FECHA INICI" in c or "FECHA_INICI" in c
            ]
            col_f_fin = [
                c
                for c in df_empleados.columns
                if "FECHA FIN" in c or "FECHA_FIN" in c
            ]

            df_empleados["INCIDENCIA_TIPO"] = (
                df_empleados[col_inc[0]] if col_inc else None
            )
            df_empleados["INCIDENCIA_INI"] = (
                df_empleados[col_f_ini[0]] if col_f_ini else None
            )
            df_empleados["INCIDENCIA_FIN"] = (
                df_empleados[col_f_fin[0]] if col_f_fin else None
            )
            st.success(
                f"¡Se cargaron {len(df_empleados)} empleados exitosamente!"
            )
    except Exception as e:
        st.error(f"Error al procesar personal: {e}")

if uploaded_prev_file is not None:
    try:
        df_semana_anterior = (
            pd.read_csv(uploaded_prev_file)
            if uploaded_prev_file.name.endswith(".csv")
            else pd.read_excel(uploaded_prev_file)
        )
        df_semana_anterior.columns = [
            str(c).upper().strip() for c in df_semana_anterior.columns
        ]
        if "CARGO" in df_semana_anterior.columns:
            df_semana_anterior["CARGO"] = (
                df_semana_anterior["CARGO"].astype(str).str.strip().str.upper()
            )
        st.success("✅ Malla anterior cargada correctamente.")
    except Exception as e:
        st.warning(f"No se pudo leer la semana anterior: {e}")

# ==========================================
# REEMPLAZOS INTER-CARGO Y SOBRE TIEMPO
# ==========================================
reemplazos_config = {}
sobretiempo_config = {}

if df_empleados is not None:
    st.sidebar.markdown("---")
    st.sidebar.header("🔄 Cobertura Inter-Cargo / Reemplazos")
    activa_reemplazo = st.sidebar.checkbox(
        "¿Asignar colaboradores a otro cargo?", value=False
    )

    if activa_reemplazo:
        if "lista_reemplazos" not in st.session_state:
            st.session_state.lista_reemplazos = []

        lista_empleados_nombres = df_empleados["NOMBRE"].tolist()
        cargos_disponibles = sorted(
            list(
                set(
                    df_empleados["CARGO"].tolist()
                    + list(TURNOS_DEFAULT_POR_CARGO.keys())
                )
            )
        )

        emps_sel = st.sidebar.multiselect(
            "1. Seleccionar Colaborador(es):",
            lista_empleados_nombres,
            key="sel_remplazo",
        )
        cargo_destino_sel = st.sidebar.selectbox(
            "2. Cargo secundario a cubrir:", cargos_disponibles
        )

        if st.sidebar.button("➕ Agregar Polivalencia"):
            for emp_nombre in emps_sel:
                row_emp = df_empleados[
                    df_empleados["NOMBRE"] == emp_nombre
                ].iloc[0]
                cod = str(row_emp["CODIGO"]).strip()
                c_orig = row_emp["CARGO"]

                if not any(
                    item["CODIGO"] == cod
                    for item in st.session_state.lista_reemplazos
                ):
                    st.session_state.lista_reemplazos.append(
                        {
                            "CODIGO": cod,
                            "NOMBRE": emp_nombre,
                            "CARGO ORIGEN": c_orig,
                            "CARGO A CUBRIR": cargo_destino_sel,
                        }
                    )

        if st.session_state.lista_reemplazos:
            st.sidebar.subheader("📋 Reemplazos Programados")
            df_temp_reemplazos = pd.DataFrame(st.session_state.lista_reemplazos)
            st.sidebar.dataframe(
                df_temp_reemplazos[["NOMBRE", "CARGO A CUBRIR"]],
                use_container_width=True,
            )

            if st.sidebar.button("🗑️ Limpiar Reemplazos"):
                st.session_state.lista_reemplazos = []

            for item in st.session_state.lista_reemplazos:
                reemplazos_config[item["CODIGO"]] = item["CARGO A CUBRIR"]

    # SECCIÓN SOBRE TIEMPO
    st.sidebar.markdown("---")
    st.sidebar.header("⏰ Configuración de Sobre Tiempo (Horas Extras)")
    activa_sobretiempo = st.sidebar.checkbox(
        "¿Habilitar sobre tiempo para colaboradores?", value=False
    )

    if activa_sobretiempo:
        if "lista_sobretiempo" not in st.session_state:
            st.session_state.lista_sobretiempo = []

        lista_empleados_nombres_st = df_empleados["NOMBRE"].tolist()

        emps_st_sel = st.sidebar.multiselect(
            "1. Seleccionar Colaborador(es):",
            lista_empleados_nombres_st,
            key="sel_st",
        )
        max_horas_st = st.sidebar.number_input(
            "2. Máximo de Horas Extras por Día:",
            min_value=1,
            max_value=8,
            value=4,
            step=1,
        )

        if st.sidebar.button("➕ Habilitar Sobre Tiempo"):
            for emp_nombre in emps_st_sel:
                row_emp = df_empleados[
                    df_empleados["NOMBRE"] == emp_nombre
                ].iloc[0]
                cod = str(row_emp["CODIGO"]).strip()

                idx_existente = next(
                    (
                        i
                        for i, item in enumerate(
                            st.session_state.lista_sobretiempo
                        )
                        if item["CODIGO"] == cod
                    ),
                    None,
                )
                if idx_existente is not None:
                    st.session_state.lista_sobretiempo[idx_existente][
                        "MAX_HORAS_EXTRA"
                    ] = max_horas_st
                else:
                    st.session_state.lista_sobretiempo.append(
                        {
                            "CODIGO": cod,
                            "NOMBRE": emp_nombre,
                            "CARGO": row_emp["CARGO"],
                            "MAX_HORAS_EXTRA": max_horas_st,
                        }
                    )

        if st.session_state.lista_sobretiempo:
            st.sidebar.subheader("📋 Sobre Tiempo Habilitado")
            df_temp_st = pd.DataFrame(st.session_state.lista_sobretiempo)
            st.sidebar.dataframe(
                df_temp_st[["NOMBRE", "MAX_HORAS_EXTRA"]],
                use_container_width=True,
            )

            if st.sidebar.button("🗑️ Limpiar Sobre Tiempo"):
                st.session_state.lista_sobretiempo = []

            for item in st.session_state.lista_sobretiempo:
                sobretiempo_config[item["CODIGO"]] = item["MAX_HORAS_EXTRA"]

# ==========================================
# CONSTRUCCIÓN DE DEMANDA Y GENERACIÓN
# ==========================================
matriz_demanda = {}
if df_empleados is not None:
    cargos_unicos = df_empleados["CARGO"].dropna().unique().tolist()
    for cargo in cargos_unicos:
        cargo_clean = str(cargo).strip().upper()
        matriz_demanda[cargo_clean] = {d: [] for d in DIAS_SEMANA_ES}
        sub_mat = (
            df_tareas_req[df_tareas_req["CARGO"] == cargo_clean]
            if "CARGO" in df_tareas_req.columns
            else pd.DataFrame()
        )

        col_habil = next((c for c in sub_mat.columns if "HABIL" in c), None)
        col_sab = next((c for c in sub_mat.columns if "SAB" in c), None)
        col_dom = next((c for c in sub_mat.columns if "DOM" in c), None)

        req_habil = (
            [
                str(x).strip()
                for x in sub_mat[col_habil].dropna().tolist()
                if str(x).strip() != ""
            ]
            if col_habil
            else []
        )
        req_sab = (
            [
                str(x).strip()
                for x in sub_mat[col_sab].dropna().tolist()
                if str(x).strip() != ""
            ]
            if col_sab
            else []
        )
        req_dom = (
            [
                str(x).strip()
                for x in sub_mat[col_dom].dropna().tolist()
                if str(x).strip() != ""
            ]
            if col_dom
            else []
        )

        if not req_habil and not req_sab and not req_dom:
            defaults = TURNOS_DEFAULT_POR_CARGO.get(cargo_clean, {})
            req_habil = defaults.get("habil", [])
            req_sab = defaults.get("sabado", [])
            req_dom = defaults.get("domingo", [])

        for dh in ["LUNES", "MARTES", "MIÉRCOLES", "JUEVES", "VIERNES"]:
            matriz_demanda[cargo_clean][dh] = list(req_habil)
        matriz_demanda[cargo_clean]["SÁBADO"] = list(req_sab)
        matriz_demanda[cargo_clean]["DOMINGO"] = list(req_dom)

# ==========================================
# FUNCION DE PRE-AUDITORÍA Y BALANCE
# ==========================================
def calcular_balance_capacidad(
    df_personal,
    semanas_count,
    fecha_base_date,
    demandas_matriz,
    libres_base,
    festivos_list=None,
):
    if festivos_list is None:
        festivos_list = []

    dias_totales = semanas_count * 7
    fecha_base = datetime.combine(fecha_base_date, time.min)
    cols_dias = []

    for i in range(dias_totales):
        f = fecha_base + timedelta(days=i)
        cols_dias.append(
            f"{DIAS_SEMANA_ES[f.weekday()]}\n{f.strftime('%d-%b')}"
        )

    resumen_balance = []
    cargos = df_personal["CARGO"].unique()

    for cargo in cargos:
        cargo_clean = str(cargo).strip().upper()
        tot_emp = len(df_personal[df_personal["CARGO"] == cargo_clean])

        for idx_d, col_d in enumerate(cols_dias):
            f_act = fecha_base + timedelta(days=idx_d)
            sem_idx = idx_d // 7
            f_sem_ini = fecha_base_date + timedelta(days=sem_idx * 7)
            f_sem_fin = f_sem_ini + timedelta(days=6)
            festivos_sem = sum(
                1 for ff in festivos_list if f_sem_ini <= ff <= f_sem_fin
            )

            libres_sem_total = min(6, libres_base + festivos_sem)
            efectivo_estimado = max(
                0, tot_emp - round(tot_emp * (libres_sem_total / 7.0))
            )

            nom_dia = DIAS_SEMANA_ES[f_act.weekday()]
            req_turnos = demandas_matriz.get(cargo_clean, {}).get(nom_dia, [])
            cant_req = len(req_turnos)
            diferencia = efectivo_estimado - cant_req

            resumen_balance.append(
                {
                    "CARGO": cargo_clean,
                    "DÍA": col_d.replace("\n", " "),
                    "FECHA_ORDEN": f_act,
                    "TAREAS REQUERIDAS": cant_req,
                    "PERSONAL DISPONIBLE": efectivo_estimado,
                    "BALANCE (+ SOBRA / - FALTA)": diferencia,
                }
            )

    return pd.DataFrame(resumen_balance)


# ==========================================
# MOTOR PRINCIPAL DE GENERACIÓN DE MALLA
# ==========================================
def generar_malla_matriz(
    df_personal,
    semanas_count,
    fecha_base_date,
    reglas_demanda,
    libres_base,
    festivos_list,
    df_prev=None,
    mapa_reemplazos=None,
    mapa_sobretiempo=None,
):
    if mapa_reemplazos is None:
        mapa_reemplazos = {}
    if mapa_sobretiempo is None:
        mapa_sobretiempo = {}
    if festivos_list is None:
        festivos_list = []

    dias_totales = semanas_count * 7
    fecha_base = datetime.combine(fecha_base_date, time.min)
    anio_ref = fecha_base_date.year
    columnas_fechas, fechas_dt = [], []

    for i in range(dias_totales):
        f_actual = fecha_base + timedelta(days=i)
        fechas_dt.append(f_actual)
        columnas_fechas.append(
            f"{DIAS_SEMANA_ES[f_actual.weekday()]}\n{f_actual.strftime('%d-%b')}"
        )

    # 1. PROCESAR MALLA ANTERIOR
    info_historial = {}
    dias_trabajados_previos = {}
    historial_analistas_malla_prev = {}

    if df_prev is not None and "CODIGO" in df_prev.columns:
        cols_dias_prev = [
            c
            for c in df_prev.columns
            if c not in ["CODIGO", "NOMBRE", "CARGO", "ESTADO"]
        ]
        if cols_dias_prev:
            fecha_ultimo_dia_prev = fecha_base - timedelta(days=1)

            for _, fila in df_prev.iterrows():
                c_cod = str(fila["CODIGO"]).strip()
                val_ult_raw = str(fila[cols_dias_prev[-1]]).strip()
                val_ult_limpio = (
                    val_ult_raw.split()[0] if " " in val_ult_raw else val_ult_raw
                )
                es_descanso_o_inc = val_ult_limpio.upper() in NO_WORKING_TERMS

                secuencia_prev = [
                    str(fila[col_p]).strip().split()[0]
                    for col_p in cols_dias_prev
                ]
                historial_analistas_malla_prev[c_cod] = secuencia_prev

                conteo_cont = 0
                for col_prev in reversed(cols_dias_prev):
                    v_p = str(fila[col_prev]).strip().split()[0].upper()
                    if v_p not in NO_WORKING_TERMS:
                        conteo_cont += 1
                    else:
                        break

                dias_trabajados_previos[c_cod] = conteo_cont

                salida_dt_prev = None
                ultimo_turno_previo_desc = None
                franja_previa_desc = None

                if not es_descanso_o_inc:
                    _, salida_dt_prev = calcular_datetimes_turno(
                        fecha_ultimo_dia_prev, val_ult_raw
                    )
                else:
                    for col_prev in reversed(cols_dias_prev):
                        v_p = str(fila[col_prev]).strip().split()[0]
                        if v_p.upper() not in NO_WORKING_TERMS:
                            ultimo_turno_previo_desc = v_p
                            franja_previa_desc = clasificar_franja(v_p)
                            break

                info_historial[c_cod] = {
                    "ultimo_turno": None
                    if es_descanso_o_inc
                    else val_ult_limpio,
                    "ultima_franja": None
                    if es_descanso_o_inc
                    else clasificar_franja(val_ult_limpio),
                    "termino_en_descanso": es_descanso_o_inc,
                    "salida_previa_dt": salida_dt_prev,
                    "turno_previo_desc": ultimo_turno_previo_desc,
                    "franja_previa_desc": franja_previa_desc,
                }

    # 2. DETECTAR VACACIONES DE ANALISTAS
    hay_analista_vacaciones = False
    f_inicio_malla = fecha_base_date
    f_fin_malla = fecha_base_date + timedelta(days=dias_totales - 1)

    for _, emp in df_personal.iterrows():
        if "ANALISTA" in str(emp["CARGO"]).upper():
            inc_t = str(emp.get("INCIDENCIA_TIPO", "")).upper()
            f_ini = parsear_fecha_incidencia(emp.get("INCIDENCIA_INI"), anio_ref)
            f_fin = parsear_fecha_incidencia(emp.get("INCIDENCIA_FIN"), anio_ref)

            if "VACACIONES" in inc_t and f_ini and f_fin:
                if not (f_fin < f_inicio_malla or f_ini > f_fin_malla):
                    hay_analista_vacaciones = True
                    break

    patrones_analistas = (
        PATRONES_ANALISTAS_4 if hay_analista_vacaciones else PATRONES_ANALISTAS_5
    )

    programacion_matriz = {}
    idx_patron_analistas_counter = 0

    for _, emp in df_personal.iterrows():
        cod = str(emp["CODIGO"]).strip()
        cargo_original = str(emp["CARGO"]).strip().upper()
        cargo_secundario = mapa_reemplazos.get(cod, None)
        max_st = mapa_sobretiempo.get(cod, 0)
        hist = info_historial.get(cod, {})

        programacion_matriz[cod] = {
            "CODIGO": cod,
            "NOMBRE": emp["NOMBRE"],
            "CARGO_ORIGINAL": cargo_original,
            "CARGO_SECUNDARIO": cargo_secundario,
            "MAX_HORAS_EXTRA_DIA": max_st,
            "CARGO": cargo_original,
            "INCIDENCIA_TIPO": emp.get("INCIDENCIA_TIPO"),
            "INCIDENCIA_INI": parsear_fecha_incidencia(
                emp.get("INCIDENCIA_INI"), anio_ref
            ),
            "INCIDENCIA_FIN": parsear_fecha_incidencia(
                emp.get("INCIDENCIA_FIN"), anio_ref
            ),
            "TURNO_FIJO_BLOQUE": hist.get("ultimo_turno")
            if not hist.get("termino_en_descanso", True)
            else None,
            "TURNO_PREVIO_DESCANSO": hist.get("turno_previo_desc"),
            "FRANJA_PREVIA_DESCANSO": hist.get("franja_previa_desc"),
            "VIENE_DE_DESCANSO": hist.get("termino_en_descanso", False),
            "ULTIMA_FRANJA": hist.get("ultima_franja"),
            "SALIDA_PREVIA_DT": hist.get("salida_previa_dt"),
            "DIAS_SEGUIDOS_TRABAJADOS": dias_trabajados_previos.get(cod, 0),
            "POSICION_SECUENCIA_ANALISTA": 0,
            "HISTORIAL_CARGOS_DIARIOS": {},
            "HISTORIAL_TURNOS_LIMPIOS": {},
            "HISTORIAL_SOBRETIEMPO": {},
            "LIBRES_POR_SEMANA": {s: 0 for s in range(semanas_count)},
        }

        if "ANALISTA" in cargo_original:
            sec_previa = historial_analistas_malla_prev.get(cod, [])
            idx_patron, pos_siguiente = encontrar_posicion_secuencia_analista(
                sec_previa, patrones_analistas
            )

            if idx_patron is not None:
                programacion_matriz[cod]["PATRON_BASE"] = idx_patron
                programacion_matriz[cod][
                    "POSICION_SECUENCIA_ANALISTA"
                ] = pos_siguiente
            else:
                programacion_matriz[cod]["PATRON_BASE"] = (
                    idx_patron_analistas_counter % len(patrones_analistas)
                )
                programacion_matriz[cod]["POSICION_SECUENCIA_ANALISTA"] = 0
                idx_patron_analistas_counter += 1

    tecnicos_planta_cods = [
        cod
        for cod, d in programacion_matriz.items()
        if "TÉCNICO" in d["CARGO_ORIGINAL"] or "TECNICO" in d["CARGO_ORIGINAL"]
    ]

    # Conteo total por cargo para evitar libres simultáneos
    total_empleados_cargo = {}
    for d in programacion_matriz.values():
        c_o = d["CARGO_ORIGINAL"]
        total_empleados_cargo[c_o] = total_empleados_cargo.get(c_o, 0) + 1

    # 3. GENERACIÓN DÍA A DÍA
    for idx_dia, col_nombre in enumerate(columnas_fechas):
        fecha_col = fechas_dt[idx_dia]
        fecha_actual_date = fecha_col.date()
        nombre_dia_semana = DIAS_SEMANA_ES[fecha_col.weekday()]
        semana_actual_idx = idx_dia // 7
        dias_restantes_semana = 6 - (idx_dia % 7)

        f_sem_ini = fecha_base_date + timedelta(days=semana_actual_idx * 7)
        f_sem_fin = f_sem_ini + timedelta(days=6)
        festivos_en_semana = sum(
            1 for ff in festivos_list if f_sem_ini <= ff <= f_sem_fin
        )
        cuota_libres_semana = min(6, libres_base + festivos_en_semana)

        demandas_dia_actual = {
            c_k: list(v_dict.get(nombre_dia_semana, []))
            for c_k, v_dict in reglas_demanda.items()
        }

        libres_asignados_hoy_por_cargo = {}

        # ETAPA 1: INCIDENCIAS, TOPE 10 DÍAS Y DESCANSOS ESCALONADOS (NOCHE -> L -> TARDE)
        disponibles_hoy = []
        pendientes_evaluar_libre = []

        for cod_e, d_e in programacion_matriz.items():
            cargo_orig = d_e["CARGO_ORIGINAL"]
            es_analista = "ANALISTA" in cargo_orig

            # 1A. Incidencias
            if (
                d_e["INCIDENCIA_TIPO"]
                and d_e["INCIDENCIA_INI"]
                and d_e["INCIDENCIA_FIN"]
            ):
                if (
                    d_e["INCIDENCIA_INI"]
                    <= fecha_actual_date
                    <= d_e["INCIDENCIA_FIN"]
                ):
                    programacion_matriz[cod_e][col_nombre] = d_e[
                        "INCIDENCIA_TIPO"
                    ]
                    d_e.update(
                        {
                            "TURNO_FIJO_BLOQUE": None,
                            "SALIDA_PREVIA_DT": None,
                            "DIAS_SEGUIDOS_TRABAJADOS": 0,
                        }
                    )
                    d_e["LIBRES_POR_SEMANA"][semana_actual_idx] += 1
                    d_e["HISTORIAL_CARGOS_DIARIOS"][col_nombre] = cargo_orig
                    d_e["HISTORIAL_TURNOS_LIMPIOS"][col_nombre] = d_e[
                        "INCIDENCIA_TIPO"
                    ]
                    continue

            if es_analista:
                disponibles_hoy.append(cod_e)
                continue

            pendientes_evaluar_libre.append(cod_e)

        # Ordenar candidatos a libre: primero los que llegaron a 10 días, luego los que tienen pocos días restantes
        # y prioridad a quienes vienen de TURNO NOCTURNO (Noche -> Libre -> Tarde)
        pendientes_evaluar_libre.sort(
            key=lambda c: (
                0
                if programacion_matriz[c]["DIAS_SEGUIDOS_TRABAJADOS"] >= 10
                else 1,
                0
                if (
                    cuota_libres_semana
                    - programacion_matriz[c]["LIBRES_POR_SEMANA"][
                        semana_actual_idx
                    ]
                )
                > dias_restantes_semana
                else 1,
                0 if programacion_matriz[c]["ULTIMA_FRANJA"] == "NOCHE" else 1,
                -programacion_matriz[c]["DIAS_SEGUIDOS_TRABAJADOS"],
            )
        )

        for cod_e in pendientes_evaluar_libre:
            d_e = programacion_matriz[cod_e]
            cargo_orig = d_e["CARGO_ORIGINAL"]
            forzar_libre_10_dias = d_e["DIAS_SEGUIDOS_TRABAJADOS"] >= 10

            libres_tomados = d_e["LIBRES_POR_SEMANA"][semana_actual_idx]
            libres_faltantes = max(0, cuota_libres_semana - libres_tomados)
            obligado_por_fin_semana = libres_faltantes > dias_restantes_semana

            # Máximo de libres simultáneos permitidos hoy para este cargo
            tot_c = total_empleados_cargo.get(cargo_orig, 1)
            max_libres_simultaneos = max(1, (tot_c + 6) // 7)
            ya_libres_hoy_cargo = libres_asignados_hoy_por_cargo.get(
                cargo_orig, 0
            )

            # Si aún no cumple su cuota semanal y hay cupo de descanso hoy en su cargo:
            # Dar libre inmediatamente si viene de NOCHE (para rotarlo a TARDE) o si lleva >= 4 días
            quiere_rotar_noche_a_tarde = (
                libres_faltantes > 0
                and d_e["ULTIMA_FRANJA"] == "NOCHE"
                and d_e["DIAS_SEGUIDOS_TRABAJADOS"] >= 2
                and ya_libres_hoy_cargo < max_libres_simultaneos
            )

            toma_libre_escalonado = (
                libres_faltantes > 0
                and d_e["DIAS_SEGUIDOS_TRABAJADOS"] >= 4
                and ya_libres_hoy_cargo < max_libres_simultaneos
            )

            if (
                forzar_libre_10_dias
                or obligado_por_fin_semana
                or quiere_rotar_noche_a_tarde
                or toma_libre_escalonado
            ):
                programacion_matriz[cod_e][col_nombre] = "L"
                d_e.update(
                    {
                        "FRANJA_PREVIA_DESCANSO": d_e.get("ULTIMA_FRANJA"),
                        "VIENE_DE_DESCANSO": True,
                        "TURNO_FIJO_BLOQUE": None,
                        "SALIDA_PREVIA_DT": None,
                        "DIAS_SEGUIDOS_TRABAJADOS": 0,
                    }
                )
                d_e["LIBRES_POR_SEMANA"][semana_actual_idx] += 1
                libres_asignados_hoy_por_cargo[cargo_orig] = (
                    ya_libres_hoy_cargo + 1
                )
                d_e["HISTORIAL_CARGOS_DIARIOS"][col_nombre] = cargo_orig
                d_e["HISTORIAL_TURNOS_LIMPIOS"][col_nombre] = "L"
            else:
                disponibles_hoy.append(cod_e)

        # ETAPA 2: ANALISTAS DE OPERACIONES (ROTA SECUENCIALMENTE DÍA A DÍA)
        analistas_hoy = [
            c
            for c in disponibles_hoy
            if "ANALISTA" in programacion_matriz[c]["CARGO_ORIGINAL"]
        ]
        req_analistas = demandas_dia_actual.get(
            "ANALISTA DE OPERACIONES", []
        )

        for cod_an in analistas_hoy:
            d_an = programacion_matriz[cod_an]
            idx_pat = d_an["PATRON_BASE"]
            pos_actual = d_an["POSICION_SECUENCIA_ANALISTA"]

            turno_sugerido = patrones_analistas[idx_pat][pos_actual]

            programacion_matriz[cod_an][col_nombre] = turno_sugerido
            d_an["HISTORIAL_CARGOS_DIARIOS"][col_nombre] = (
                "ANALISTA DE OPERACIONES"
            )
            d_an["HISTORIAL_TURNOS_LIMPIOS"][col_nombre] = turno_sugerido

            d_an["POSICION_SECUENCIA_ANALISTA"] = (pos_actual + 1) % 14

            if turno_sugerido not in NO_WORKING_TERMS:
                d_an["ULTIMA_FRANJA"] = clasificar_franja(turno_sugerido)
                _, dt_salida = calcular_datetimes_turno(
                    fecha_col, turno_sugerido
                )
                d_an["SALIDA_PREVIA_DT"] = dt_salida
                d_an["VIENE_DE_DESCANSO"] = False
                d_an["DIAS_SEGUIDOS_TRABAJADOS"] += 1
            else:
                d_an["SALIDA_PREVIA_DT"] = None
                d_an["VIENE_DE_DESCANSO"] = True
                d_an["DIAS_SEGUIDOS_TRABAJADOS"] = 0
                d_an["LIBRES_POR_SEMANA"][semana_actual_idx] += 1

            if turno_sugerido in req_analistas:
                req_analistas.remove(turno_sugerido)

        # ETAPA 3: TÉCNICOS PLANTA
        tecnicos_planta_hoy = sorted(
            [c for c in disponibles_hoy if c in tecnicos_planta_cods],
            key=lambda c: prioridad_por_restriccion_horaria(
                c, programacion_matriz
            ),
        )
        turnos_tec_disp = demandas_dia_actual.get(
            "TÉCNICO DE OPERACIONES", []
        )

        hay_polivalentes_tec = any(
            programacion_matriz[c]["CARGO_SECUNDARIO"]
            == "TÉCNICO DE OPERACIONES"
            for c in disponibles_hoy
        )
        orden_turnos_tec = (
            ["06:00-15:00", "03:00-11:00", "11:00-19:00", "19:00-27:00"]
            if hay_polivalentes_tec
            else ["03:00-11:00", "11:00-19:00", "19:00-27:00", "06:00-15:00"]
        )

        for t_target in orden_turnos_tec:
            if t_target in turnos_tec_disp:
                for cod_tec in tecnicos_planta_hoy:
                    if (
                        programacion_matriz[cod_tec].get(col_nombre)
                        is not None
                    ):
                        continue
                    d_tec = programacion_matriz[cod_tec]
                    dt_ent, dt_salida = calcular_datetimes_turno(
                        fecha_col, t_target
                    )
                    min_desc = (
                        24.0 if d_tec.get("VIENE_DE_DESCANSO") else 12.0
                    )
                    if calcular_descanso_suficiente(
                        d_tec["SALIDA_PREVIA_DT"], dt_ent, min_horas=min_desc
                    ):
                        programacion_matriz[cod_tec][col_nombre] = t_target
                        d_tec.update(
                            {
                                "TURNO_FIJO_BLOQUE": t_target,
                                "ULTIMA_FRANJA": clasificar_franja(
                                    t_target
                                ),
                                "VIENE_DE_DESCANSO": False,
                                "SALIDA_PREVIA_DT": dt_salida,
                                "DIAS_SEGUIDOS_TRABAJADOS": d_tec[
                                    "DIAS_SEGUIDOS_TRABAJADOS"
                                ]
                                + 1,
                            }
                        )
                        d_tec["HISTORIAL_CARGOS_DIARIOS"][col_nombre] = (
                            "TÉCNICO DE OPERACIONES"
                        )
                        d_tec["HISTORIAL_TURNOS_LIMPIOS"][col_nombre] = (
                            t_target
                        )
                        turnos_tec_disp.remove(t_target)
                        break

        for cod_tec in tecnicos_planta_hoy:
            d_tec = programacion_matriz[cod_tec]
            if programacion_matriz[cod_tec].get(col_nombre) is None:
                if turnos_tec_disp:
                    cand_t = turnos_tec_disp.pop(0)
                    _, dt_salida = calcular_datetimes_turno(
                        fecha_col, cand_t
                    )
                    programacion_matriz[cod_tec][col_nombre] = cand_t
                    d_tec.update(
                        {
                            "TURNO_FIJO_BLOQUE": cand_t,
                            "ULTIMA_FRANJA": clasificar_franja(cand_t),
                            "VIENE_DE_DESCANSO": False,
                            "SALIDA_PREVIA_DT": dt_salida,
                            "DIAS_SEGUIDOS_TRABAJADOS": d_tec[
                                "DIAS_SEGUIDOS_TRABAJADOS"
                            ]
                            + 1,
                        }
                    )
                    d_tec["HISTORIAL_CARGOS_DIARIOS"][col_nombre] = (
                        "TÉCNICO DE OPERACIONES"
                    )
                    d_tec["HISTORIAL_TURNOS_LIMPIOS"][col_nombre] = cand_t
                else:
                    programacion_matriz[cod_tec][col_nombre] = "AO"
                    d_tec["HISTORIAL_CARGOS_DIARIOS"][col_nombre] = (
                        "TÉCNICO DE OPERACIONES"
                    )
                    d_tec["HISTORIAL_TURNOS_LIMPIOS"][col_nombre] = "AO"

        # ETAPA 4: COBERTURA INTER-CARGO (SOLO SI SU PUESTO ORIGINAL TIENE RESPALDO SUFICIENTE)
        resto_empleados = sorted(
            [
                c
                for c in disponibles_hoy
                if "ANALISTA"
                not in programacion_matriz[c]["CARGO_ORIGINAL"]
                and c not in tecnicos_planta_hoy
            ],
            key=lambda c: prioridad_por_restriccion_horaria(
                c, programacion_matriz
            ),
        )

        empleados_polivalentes = [
            c
            for c in resto_empleados
            if programacion_matriz[c]["CARGO_SECUNDARIO"] is not None
        ]

        for cod_poli in empleados_polivalentes:
            if programacion_matriz[cod_poli].get(col_nombre) is not None:
                continue

            d_poli = programacion_matriz[cod_poli]
            cargo_orig_poli = d_poli["CARGO_ORIGINAL"]
            cargo_sec = d_poli["CARGO_SECUNDARIO"]
            turnos_faltantes_sec = demandas_dia_actual.get(cargo_sec, [])

            # Verificar cuántos compañeros quedan disponibles para cubrir el cargo original
            disp_cargo_orig = sum(
                1
                for c_o in resto_empleados
                if programacion_matriz[c_o]["CARGO_ORIGINAL"]
                == cargo_orig_poli
                and programacion_matriz[c_o].get(col_nombre) is None
            )
            dem_cargo_orig = len(demandas_dia_actual.get(cargo_orig_poli, []))

            # Solo sale a cubrir otro cargo si su puesto original no queda desatendido
            if disp_cargo_orig <= dem_cargo_orig:
                continue

            if turnos_faltantes_sec:
                franja_pref = (
                    "TARDE"
                    if (
                        d_poli.get("VIENE_DE_DESCANSO")
                        and d_poli.get("FRANJA_PREVIA_DESCANSO") == "NOCHE"
                    )
                    else d_poli.get("ULTIMA_FRANJA")
                )
                turnos_ordenados_sec = sorted(
                    list(turnos_faltantes_sec),
                    key=lambda t: (
                        0 if clasificar_franja(t) == franja_pref else 1,
                        extraer_horas(t)[0] if extraer_horas(t) else 99,
                    ),
                )
                for cand_t in turnos_ordenados_sec:
                    dt_ent, dt_salida = calcular_datetimes_turno(
                        fecha_col, cand_t
                    )
                    min_desc = (
                        24.0 if d_poli.get("VIENE_DE_DESCANSO") else 12.0
                    )
                    if calcular_descanso_suficiente(
                        d_poli["SALIDA_PREVIA_DT"], dt_ent, min_horas=min_desc
                    ):
                        sufijo_cargo = (
                            " TO"
                            if "TÉCNICO" in cargo_sec
                            or "TECNICO" in cargo_sec
                            else ""
                        )
                        programacion_matriz[cod_poli][
                            col_nombre
                        ] = f"{cand_t}{sufijo_cargo}"
                        d_poli.update(
                            {
                                "TURNO_FIJO_BLOQUE": cand_t,
                                "ULTIMA_FRANJA": clasificar_franja(cand_t),
                                "VIENE_DE_DESCANSO": False,
                                "SALIDA_PREVIA_DT": dt_salida,
                                "DIAS_SEGUIDOS_TRABAJADOS": d_poli[
                                    "DIAS_SEGUIDOS_TRABAJADOS"
                                ]
                                + 1,
                            }
                        )
                        d_poli["HISTORIAL_CARGOS_DIARIOS"][
                            col_nombre
                        ] = cargo_sec
                        d_poli["HISTORIAL_TURNOS_LIMPIOS"][
                            col_nombre
                        ] = cand_t
                        turnos_faltantes_sec.remove(cand_t)
                        break

        # ETAPA 5: COBERTURA REGULAR PUESTO ORIGINAL (CON REGLA NOCHE -> L -> TARDE)
        for cod_emp in resto_empleados:
            if programacion_matriz[cod_emp].get(col_nombre) is not None:
                continue

            d_emp = programacion_matriz[cod_emp]
            cargo_orig = d_emp["CARGO_ORIGINAL"]
            turnos_disp_cargo = demandas_dia_actual.get(cargo_orig, [])

            if turnos_disp_cargo:
                viene_desc = d_emp.get("VIENE_DE_DESCANSO", False)
                franja_ref = (
                    d_emp["FRANJA_PREVIA_DESCANSO"]
                    if viene_desc
                    else d_emp["ULTIMA_FRANJA"]
                )
                franjas_permitidas = (
                    obtener_siguiente_franja_permitida(
                        franja_ref, viene_de_descanso=viene_desc
                    )
                    if franja_ref
                    else ["MAÑANA", "TARDE", "NOCHE"]
                )

                # Franja objetivo principal (ej. si viene de L tras NOCHE -> TARDE)
                franja_objetivo = franjas_permitidas[0]

                cand_list = [
                    t
                    for t in turnos_disp_cargo
                    if clasificar_franja(t) in franjas_permitidas
                ]

                cand_list.sort(
                    key=lambda t: (
                        0 if clasificar_franja(t) == franja_objetivo else 1,
                        extraer_horas(t)[0] if extraer_horas(t) else 99,
                    )
                )

                turno_a_asignar = None
                min_desc = 24.0 if viene_desc else 12.0

                for cand_t in cand_list:
                    dt_ent, dt_salida = calcular_datetimes_turno(
                        fecha_col, cand_t
                    )
                    if calcular_descanso_suficiente(
                        d_emp["SALIDA_PREVIA_DT"], dt_ent, min_horas=min_desc
                    ):
                        turno_a_asignar = cand_t
                        turnos_disp_cargo.remove(cand_t)
                        break

                if not turno_a_asignar and turnos_disp_cargo:
                    for cand_t in list(turnos_disp_cargo):
                        dt_ent, dt_salida = calcular_datetimes_turno(
                            fecha_col, cand_t
                        )
                        if calcular_descanso_suficiente(
                            d_emp["SALIDA_PREVIA_DT"],
                            dt_ent,
                            min_horas=min_desc,
                        ):
                            turno_a_asignar = cand_t
                            turnos_disp_cargo.remove(cand_t)
                            break

                if turno_a_asignar:
                    programacion_matriz[cod_emp][col_nombre] = turno_a_asignar
                    _, dt_salida = calcular_datetimes_turno(
                        fecha_col, turno_a_asignar
                    )
                    d_emp.update(
                        {
                            "TURNO_FIJO_BLOQUE": turno_a_asignar,
                            "ULTIMA_FRANJA": clasificar_franja(
                                turno_a_asignar
                            ),
                            "VIENE_DE_DESCANSO": False,
                            "SALIDA_PREVIA_DT": dt_salida,
                            "DIAS_SEGUIDOS_TRABAJADOS": d_emp[
                                "DIAS_SEGUIDOS_TRABAJADOS"
                            ]
                            + 1,
                        }
                    )
                    d_emp["HISTORIAL_CARGOS_DIARIOS"][col_nombre] = cargo_orig
                    d_emp["HISTORIAL_TURNOS_LIMPIOS"][
                        col_nombre
                    ] = turno_a_asignar
                else:
                    programacion_matriz[cod_emp][col_nombre] = "AO"
                    d_emp["HISTORIAL_CARGOS_DIARIOS"][col_nombre] = cargo_orig
                    d_emp["HISTORIAL_TURNOS_LIMPIOS"][col_nombre] = "AO"
            else:
                programacion_matriz[cod_emp][col_nombre] = "AO"
                d_emp["HISTORIAL_CARGOS_DIARIOS"][col_nombre] = cargo_orig
                d_emp["HISTORIAL_TURNOS_LIMPIOS"][col_nombre] = "AO"

        # ETAPA 5B: INTERCAMBIO INTELIGENTE PARA RESCATAR TURNOS FALTANTES CON PERSONAL EN AO
        for cargo_k, turnos_sin_cubrir in demandas_dia_actual.items():
            if not turnos_sin_cubrir or "ANALISTA" in cargo_k:
                continue

            en_ao_hoy = [
                c
                for c in disponibles_hoy
                if programacion_matriz[c]["CARGO_ORIGINAL"] == cargo_k
                and programacion_matriz[c].get(col_nombre) == "AO"
            ]
            asignados_hoy = [
                c
                for c in disponibles_hoy
                if programacion_matriz[c]["CARGO_ORIGINAL"] == cargo_k
                and programacion_matriz[c]["HISTORIAL_CARGOS_DIARIOS"].get(
                    col_nombre
                )
                == cargo_k
                and programacion_matriz[c]["HISTORIAL_TURNOS_LIMPIOS"].get(
                    col_nombre
                )
                not in NO_WORKING_TERMS
            ]

            for turno_pend in list(turnos_sin_cubrir):
                if not en_ao_hoy:
                    break
                cod_ao = en_ao_hoy[0]
                d_ao = programacion_matriz[cod_ao]
                dt_ent_pend, dt_sal_pend = calcular_datetimes_turno(
                    fecha_col, turno_pend
                )
                if not dt_ent_pend:
                    continue

                # Buscar un compañero asignado que pueda tomar el turno pendiente y ceder el suyo al de AO
                for cod_comp in asignados_hoy:
                    d_comp = programacion_matriz[cod_comp]
                    turno_comp = d_comp["HISTORIAL_TURNOS_LIMPIOS"].get(
                        col_nombre
                    )
                    dt_ent_comp, dt_sal_comp = calcular_datetimes_turno(
                        fecha_col, turno_comp
                    )
                    if not dt_ent_comp:
                        continue

                    # Para evaluar el descanso del compañero usamos su salida previa antes de hoy
                    # Aproximación segura: verificar si el turno pendiente inicia >= turno_comp o si cumple descanso
                    puede_comp_tomar_pend = (
                        dt_ent_pend >= dt_ent_comp
                        or d_comp.get("SALIDA_PREVIA_DT") is None
                    )
                    min_desc_ao = (
                        24.0 if d_ao.get("VIENE_DE_DESCANSO") else 12.0
                    )
                    puede_ao_tomar_comp = calcular_descanso_suficiente(
                        d_ao.get("SALIDA_PREVIA_DT"),
                        dt_ent_comp,
                        min_horas=min_desc_ao,
                    )

                    if puede_comp_tomar_pend and puede_ao_tomar_comp:
                        # Realizar intercambio
                        programacion_matriz[cod_comp][col_nombre] = turno_pend
                        d_comp["HISTORIAL_TURNOS_LIMPIOS"][
                            col_nombre
                        ] = turno_pend
                        d_comp["ULTIMA_FRANJA"] = clasificar_franja(turno_pend)
                        d_comp["SALIDA_PREVIA_DT"] = dt_sal_pend

                        programacion_matriz[cod_ao][col_nombre] = turno_comp
                        d_ao["HISTORIAL_TURNOS_LIMPIOS"][
                            col_nombre
                        ] = turno_comp
                        d_ao["ULTIMA_FRANJA"] = clasificar_franja(turno_comp)
                        d_ao["SALIDA_PREVIA_DT"] = dt_sal_comp
                        d_ao["VIENE_DE_DESCANSO"] = False
                        d_ao["DIAS_SEGUIDOS_TRABAJADOS"] += 1

                        turnos_sin_cubrir.remove(turno_pend)
                        en_ao_hoy.remove(cod_ao)
                        break

        # Marcar sobrantes finales
        for cod_e in disponibles_hoy:
            if programacion_matriz[cod_e].get(col_nombre) is None:
                programacion_matriz[cod_e][col_nombre] = "AO"

        # ETAPA 6: SOBRE TIEMPO (HORAS EXTRAS - FLEXIBLE Y COBERTURA PARCIAL)
        for cargo_k, v_turnos_pendientes in demandas_dia_actual.items():
            if not v_turnos_pendientes:
                continue

            candidatos_st = [
                c
                for c in disponibles_hoy
                if programacion_matriz[c]["MAX_HORAS_EXTRA_DIA"] > 0
                and col_nombre
                not in programacion_matriz[c]["HISTORIAL_SOBRETIEMPO"]
                and (
                    programacion_matriz[c]["CARGO_ORIGINAL"] == cargo_k
                    or programacion_matriz[c]["CARGO_SECUNDARIO"] == cargo_k
                )
                and programacion_matriz[c]["HISTORIAL_TURNOS_LIMPIOS"].get(
                    col_nombre
                )
                not in NO_WORKING_TERMS
            ]

            for turno_falta in list(v_turnos_pendientes):
                dt_ent_f, dt_sal_f = calcular_datetimes_turno(
                    fecha_col, turno_falta
                )
                if not dt_ent_f:
                    continue

                for cod_st in candidatos_st:
                    d_st = programacion_matriz[cod_st]
                    turno_actual_st = d_st["HISTORIAL_TURNOS_LIMPIOS"].get(
                        col_nombre
                    )
                    dt_ent_act, dt_sal_act = calcular_datetimes_turno(
                        fecha_col, turno_actual_st
                    )

                    if not dt_ent_act:
                        continue

                    horas_st_max = float(d_st["MAX_HORAS_EXTRA_DIA"])

                    gap_despues = (dt_ent_f - dt_sal_act).total_seconds() / 3600.0
                    gap_antes = (dt_ent_act - dt_sal_f).total_seconds() / 3600.0

                    es_contiguo_despues = 0.0 <= gap_despues <= 1.0
                    es_contiguo_antes = 0.0 <= gap_antes <= 1.0

                    if es_contiguo_despues or es_contiguo_antes:
                        duracion_turno_falta = (
                            dt_sal_f - dt_ent_f
                        ).total_seconds() / 3600.0
                        horas_st_asignadas = min(
                            duracion_turno_falta, horas_st_max
                        )

                        tipo_cobertura = (
                            "Completa"
                            if horas_st_asignadas >= duracion_turno_falta
                            else "Parcial"
                        )

                        d_st["HISTORIAL_SOBRETIEMPO"][col_nombre] = {
                            "TURNO_BASE": turno_actual_st,
                            "TURNO_CUBIERTO_ST": turno_falta,
                            "HORAS_EXTRA": horas_st_asignadas,
                            "COBERTURA": tipo_cobertura,
                        }

                        val_actual_vis = programacion_matriz[cod_st][col_nombre]
                        programacion_matriz[cod_st][
                            col_nombre
                        ] = f"{val_actual_vis} (+{int(horas_st_asignadas)}h ST)"

                        if es_contiguo_despues:
                            d_st["SALIDA_PREVIA_DT"] = dt_sal_act + timedelta(
                                hours=horas_st_asignadas
                            )

                        v_turnos_pendientes.remove(turno_falta)
                        d_st["HISTORIAL_TURNOS_LIMPIOS"][
                            col_nombre
                        ] = f"{turno_actual_st} / ST {turno_falta}"
                        candidatos_st.remove(cod_st)
                        break

    columnas_excluir = {
        "INCIDENCIA_TIPO",
        "INCIDENCIA_INI",
        "INCIDENCIA_FIN",
        "PATRON_BASE",
        "TURNO_FIJO_BLOQUE",
        "TURNO_PREVIO_DESCANSO",
        "FRANJA_PREVIA_DESCANSO",
        "VIENE_DE_DESCANSO",
        "ULTIMA_FRANJA",
        "CARGO_ORIGINAL",
        "CARGO_SECUNDARIO",
        "MAX_HORAS_EXTRA_DIA",
        "SALIDA_PREVIA_DT",
        "DIAS_SEGUIDOS_TRABAJADOS",
        "POSICION_SECUENCIA_ANALISTA",
        "HISTORIAL_CARGOS_DIARIOS",
        "HISTORIAL_TURNOS_LIMPIOS",
        "HISTORIAL_SOBRETIEMPO",
        "LIBRES_POR_SEMANA",
    }
    df_resultado = pd.DataFrame(
        [
            {k: v for k, v in datos.items() if k not in columnas_excluir}
            for datos in programacion_matriz.values()
        ]
    )

    return df_resultado, programacion_matriz


# ==========================================
# 4. INTERFAZ Y GENERACIÓN
# ==========================================
if df_empleados is not None:
    col_b1, col_b2 = st.columns([1, 2])

    with col_b1:
        if st.button(
            "🔍 1. Evaluar Capacidad y Previsión", use_container_width=True
        ):
            df_balance = calcular_balance_capacidad(
                df_empleados,
                semanas,
                fecha_inicio_date,
                matriz_demanda,
                libres_por_semana_base,
                fechas_festivas_sel,
            )
            st.session_state.df_balance = df_balance

if "df_balance" in st.session_state:
    st.markdown("---")
    st.subheader("📊 2. Vistazo de Prefactibilidad (Demandas vs Disponibilidad)")

    df_bal = st.session_state.df_balance

    pivot_balance = df_bal.pivot_table(
        index="CARGO",
        columns="DÍA",
        values="BALANCE (+ SOBRA / - FALTA)",
        aggfunc="first",
    )

    cols_dias_ordenadas = sorted(
        pivot_balance.columns.tolist(),
        key=lambda x: df_bal[df_bal["DÍA"] == x]["FECHA_ORDEN"].iloc[0],
    )
    pivot_balance = pivot_balance.reindex(columns=cols_dias_ordenadas)

    def colorear_balance(val):
        if val < 0:
            return "background-color: #ff4b4b; color: white; font-weight: bold;"
        elif val == 0:
            return "background-color: #e6ffed; color: #0d5a22;"
        return "background-color: #d0f0fd; color: #0c4a6e;"

    styler_bal = pivot_balance.style
    if hasattr(styler_bal, "map"):
        styler_bal = styler_bal.map(colorear_balance)
    else:
        styler_bal = styler_bal.applymap(colorear_balance)

    st.dataframe(styler_bal, use_container_width=True)

    hay_deficit = (pivot_balance < 0).any().any()

    if hay_deficit:
        st.warning(
            "⚠️ **Aviso de Déficit Detectado:** La cantidad de personal en algunos cargos es insuficiente para cubrir la totalidad de tareas solicitadas. Puedes ir a la barra lateral para ajustar descansos, polivalencia o sobretiempo, o bien pulsar abajo para continuar."
        )

    col_c1, col_c2 = st.columns([2, 1])
    with col_c1:
        if st.button(
            "🚀 2. Confirmar y Generar Malla Horaria Definitiva",
            type="primary",
            use_container_width=True,
        ):
            df_res, dict_matriz = generar_malla_matriz(
                df_empleados,
                semanas,
                fecha_inicio_date,
                matriz_demanda,
                libres_por_semana_base,
                fechas_festivas_sel,
                df_semana_anterior,
                reemplazos_config,
                sobretiempo_config,
            )
            st.session_state.df_resultado = df_res
            st.session_state.dict_matriz = dict_matriz

# ==========================================
# 5. VISUALIZACIÓN DE RESULTADOS GENERADOS
# ==========================================
if "df_resultado" in st.session_state:
    df_resultado = st.session_state.df_resultado
    dict_matriz = st.session_state.dict_matriz

    st.markdown("---")
    st.subheader("3. Malla Horaria Generada")
    st.dataframe(df_resultado, use_container_width=True)

    # AUDITORÍA DE HORAS Y DÍAS LIBRES POR COLABORADOR
    cols_fechas_malla = [
        c
        for c in df_resultado.columns
        if c not in ["CODIGO", "NOMBRE", "CARGO"]
    ]

    resumen_auditoria = []
    for cod_emp, d_emp in dict_matriz.items():
        dias_trab = 0
        dias_libres = 0
        dias_ao = 0
        horas_ord = 0.0
        horas_st = sum(
            v["HORAS_EXTRA"]
            for v in d_emp.get("HISTORIAL_SOBRETIEMPO", {}).values()
        )

        for col_f in cols_fechas_malla:
            val_limpio = d_emp.get("HISTORIAL_TURNOS_LIMPIOS", {}).get(
                col_f, ""
            )
            if val_limpio == "L":
                dias_libres += 1
            elif val_limpio == "AO":
                dias_ao += 1
            elif val_limpio not in NO_WORKING_TERMS and val_limpio != "":
                dias_trab += 1
                turno_base = val_limpio.split(" / ST ")[0]
                dt_i, dt_f = calcular_datetimes_turno(datetime.now(), turno_base)
                if dt_i and dt_f:
                    horas_ord += (dt_f - dt_i).total_seconds() / 3600.0

        resumen_auditoria.append(
            {
                "CÓDIGO": cod_emp,
                "NOMBRE": d_emp["NOMBRE"],
                "CARGO": d_emp["CARGO_ORIGINAL"],
                "DÍAS LABORADOS": dias_trab,
                "DÍAS LIBRES (L)": dias_libres,
                "DÍAS DISPONIBLE (AO)": dias_ao,
                "HORAS ORDINARIAS": round(horas_ord, 1),
                "HORAS EXTRA (ST)": round(horas_st, 1),
            }
        )
    df_auditoria = pd.DataFrame(resumen_auditoria)

    with st.expander(
        "📋 Auditoría Individual: Días Trabajados, Libres (L) y Horas",
        expanded=False,
    ):
        st.dataframe(df_auditoria, use_container_width=True)

    # REPORTE COBERTURA
    st.markdown("---")
    st.subheader("🔄 Reporte de Cobertura de Turnos de Diferente Cargo")

    reporte_coberturas = []
    for cod_emp, d_emp in dict_matriz.items():
        cargo_orig = d_emp["CARGO_ORIGINAL"]
        cargo_sec = d_emp["CARGO_SECUNDARIO"]
        historial_cargos = d_emp.get("HISTORIAL_CARGOS_DIARIOS", {})
        historial_turnos_limpios = d_emp.get("HISTORIAL_TURNOS_LIMPIOS", {})

        if cargo_sec:
            for s in range(semanas):
                cols_semana = cols_fechas_malla[s * 7 : (s + 1) * 7]
                dias_cubiertos = []
                for col in cols_semana:
                    if historial_cargos.get(col) == cargo_sec:
                        val_t = historial_turnos_limpios.get(col, "")
                        if val_t not in NO_WORKING_TERMS:
                            dias_cubiertos.append(f"{col.split()[0]}: {val_t}")

                reporte_coberturas.append(
                    {
                        "SEMANA": f"Semana {s + 1}",
                        "CÓDIGO": cod_emp,
                        "NOMBRE": d_emp["NOMBRE"],
                        "CARGO PERTENECIENTE": cargo_orig,
                        "CARGO CUBIERTO": cargo_sec,
                        "DETALLE DE TURNOS APORTADOS": ", ".join(dias_cubiertos)
                        if dias_cubiertos
                        else "Trabajó en su cargo original / Libre",
                    }
                )

    df_coberturas = pd.DataFrame(reporte_coberturas)
    if not df_coberturas.empty:
        st.info(
            "ℹ️ A continuación se detalla el personal que realizó coberturas en un cargo secundario:"
        )
        st.dataframe(df_coberturas, use_container_width=True)

    # REPORTE DETALLADO DE SOBRE TIEMPO ASIGNADO
    st.markdown("---")
    st.subheader("⏰ Reporte Detallado de Sobre Tiempo Asignado por Día y Tarea")

    reporte_st_detallado = []
    for cod_emp, d_emp in dict_matriz.items():
        historial_st = d_emp.get("HISTORIAL_SOBRETIEMPO", {})
        for col_f, datos_st in historial_st.items():
            reporte_st_detallado.append(
                {
                    "DÍA": col_f.replace("\n", " "),
                    "CÓDIGO": cod_emp,
                    "NOMBRE": d_emp["NOMBRE"],
                    "CARGO": d_emp["CARGO_ORIGINAL"],
                    "TURNO ORDINARIO": datos_st["TURNO_BASE"],
                    "TAREA / TURNO CUBIERTO EN ST": datos_st[
                        "TURNO_CUBIERTO_ST"
                    ],
                    "TIPO COBERTURA": datos_st.get("COBERTURA", "Completa"),
                    "HORAS EXTRA": f"{int(datos_st['HORAS_EXTRA'])} hrs",
                }
            )

    df_reporte_st = pd.DataFrame(reporte_st_detallado)
    if not df_reporte_st.empty:
        st.success(
            "✅ Se asignó sobre tiempo para cubrir los siguientes turnos desatendidos:"
        )
        st.dataframe(df_reporte_st, use_container_width=True)
    else:
        st.info(
            "ℹ️ No fue necesario asignar sobre tiempo adicional o no había personal disponible para cubrir las faltas restantes."
        )

    # RESUMEN DE FALTANTES
    st.markdown("---")
    st.subheader("🚨 Resumen Semanal de Tareas Desatendidas / Faltantes")

    for s in range(semanas):
        cols_semana = cols_fechas_malla[s * 7 : (s + 1) * 7]
        cols_semana_limpias = [c.replace("\n", " ") for c in cols_semana]

        datos_resumen_semana = []
        cargos_evaluar = df_empleados["CARGO"].unique().tolist()

        for cargo in cargos_evaluar:
            cargo_clean = str(cargo).strip().upper()
            sub_mat = (
                df_tareas_req[df_tareas_req["CARGO"] == cargo_clean]
                if "CARGO" in df_tareas_req.columns
                else pd.DataFrame()
            )

            for col_f in cols_semana:
                dia_nombre_ext = col_f.split("\n")[0].upper()
                tipo_col_mat = (
                    "SABADO"
                    if dia_nombre_ext in ["SÁBADO", "SABADO"]
                    else ("DOMINGO" if dia_nombre_ext == "DOMINGO" else "HABIL")
                )
                col_target_mat = next(
                    (c for c in sub_mat.columns if tipo_col_mat in c), None
                )

                turnos_req = (
                    [
                        str(x).strip()
                        for x in sub_mat[col_target_mat].dropna().tolist()
                        if str(x).strip() != ""
                    ]
                    if col_target_mat
                    else []
                )

                turnos_cubiertos = []
                for cod_e, d_e in dict_matriz.items():
                    if (
                        d_e.get("HISTORIAL_CARGOS_DIARIOS", {}).get(col_f)
                        == cargo_clean
                    ):
                        val_limpio = d_e.get(
                            "HISTORIAL_TURNOS_LIMPIOS", {}
                        ).get(col_f, "")
                        if val_limpio not in NO_WORKING_TERMS:
                            for sub_t in val_limpio.split(" / ST "):
                                turnos_cubiertos.append(sub_t)

                for tr in set(turnos_req):
                    cant_req_t = turnos_req.count(tr)
                    cant_cub_t = turnos_cubiertos.count(tr)
                    faltante_t = max(0, cant_req_t - cant_cub_t)

                    datos_resumen_semana.append(
                        {
                            "CARGO": cargo_clean,
                            "TURNO": tr,
                            "DÍA": col_f.replace("\n", " "),
                            "FALTAN POR ASIGNAR": faltante_t,
                        }
                    )

        df_sem_res = pd.DataFrame(datos_resumen_semana)

        if not df_sem_res.empty:
            tabla_pivot = df_sem_res.pivot_table(
                index=["CARGO", "TURNO"],
                columns="DÍA",
                values="FALTAN POR ASIGNAR",
                aggfunc="sum",
                fill_value=0,
            )

            cols_existentes_ordenadas = [
                c for c in cols_semana_limpias if c in tabla_pivot.columns
            ]
            tabla_pivot = tabla_pivot.reindex(columns=cols_existentes_ordenadas)

            def resaltar_faltantes_rojo(val):
                if isinstance(val, (int, float)) and val > 0:
                    return "background-color: #ff4b4b; color: white; font-weight: bold;"
                return "background-color: #e6ffed; color: #0d5a22;"

            st.markdown(f"##### 📌 Semana {s + 1}")

            styler = tabla_pivot.style
            if hasattr(styler, "map"):
                styler = styler.map(resaltar_faltantes_rojo)
            else:
                styler = styler.applymap(resaltar_faltantes_rojo)

            st.dataframe(styler, use_container_width=True)

    # EXPORTACIÓN
    st.markdown("---")
    st.subheader("📥 Exportación de Reportes Operativos")

    col_exp1, col_exp2 = st.columns(2)
    buffer_excel = io.BytesIO()

    with pd.ExcelWriter(buffer_excel) as writer:
        df_resultado.to_excel(
            writer, sheet_name="Malla Horaria", index=False
        )
        df_auditoria.to_excel(
            writer, sheet_name="Auditoría Horas y Libres", index=False
        )
        if not df_coberturas.empty:
            df_coberturas.to_excel(
                writer, sheet_name="Reemplazos Inter-Cargo", index=False
            )
        if not df_reporte_st.empty:
            df_reporte_st.to_excel(
                writer, sheet_name="Sobre Tiempo Asignado", index=False
            )

    buffer_excel.seek(0)

    with col_exp1:
        st.download_button(
            label="📥 Descargar Malla Horaria Completa (Excel)",
            data=buffer_excel,
            file_name=f"Malla_Horaria_{fecha_inicio_date.strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    with col_exp2:
        st.download_button(
            label="📄 Descargar Malla Simplificada (CSV)",
            data=df_resultado.to_csv(index=False).encode("utf-8"),
            file_name=f"Malla_Horaria_{fecha_inicio_date.strftime('%Y%m%d')}.csv",
            mime="text/csv",
        )
