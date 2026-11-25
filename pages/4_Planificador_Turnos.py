import copy
import datetime as dt
import io
import re
import unicodedata
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Planificador de Turnos Operativos", layout="wide")

DAYS_ES = ["LUNES", "MARTES", "MIÉRCOLES", "JUEVES", "VIERNES", "SÁBADO", "DOMINGO"]
SHORT_DAYS_ES = ["LUN", "MAR", "MIÉ", "JUE", "VIE", "SÁB", "DOM"]
MONTHS_ES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]

CARGOS_ORDEN = [
    "Analista de Operaciones",
    "Auxiliar de Operaciones",
    "Operador Líder",
    "Técnico de Operaciones",
    "Auxiliar de Alistamiento",
    "Conductor de Operaciones",
    "Personal de Mantenimiento",
]

SPECIAL_STATES = ("L", "VACACIONES", "SUSPEND", "INCAPACIDAD", "PERMISO")
NOVEDAD_OPTIONS = ["VACACIONES", "SUSPEND", "L", "INCAPACIDAD", "PERMISO"]

DEFAULT_ROLE_CONFIGS = {
    "Analista de Operaciones": {
        "rotation_pool": ["19:00-27:00", "11:00-19:00", "06:00-15:00", "03:00-11:00"],
        "daily_targets": {"19:00-27:00": 1, "11:00-19:00": 1, "06:00-15:00": 0, "03:00-11:00": 1},
        "saturday_targets": {"19:00-27:00": 1, "11:00-19:00": 1, "06:00-15:00": 0, "03:00-11:00": 1},
        "sunday_targets": {"19:00-27:00": 1, "11:00-19:00": 1, "06:00-15:00": 0, "03:00-11:00": 1},
    },
    "Auxiliar de Operaciones": {
        "rotation_pool": ["19:00-27:00", "23:00-31:00", "11:00-19:00", "03:00-11:00"],
        "daily_targets": {"19:00-27:00": 2, "23:00-31:00": 1, "11:00-19:00": 2, "03:00-11:00": 2},
        "saturday_targets": {"19:00-27:00": 2, "23:00-31:00": 1, "11:00-19:00": 2, "03:00-11:00": 2},
        "sunday_targets": {"19:00-27:00": 2, "23:00-31:00": 1, "11:00-19:00": 2, "03:00-11:00": 2},
    },
    "Operador Líder": {
        "rotation_pool": ["20:00-28:00"],
        "daily_targets": {"20:00-28:00": 1},
        "saturday_targets": {"20:00-28:00": 1},
        "sunday_targets": {"20:00-28:00": 1},
    },
    "Técnico de Operaciones": {
        "rotation_pool": ["19:00-27:00", "11:00-19:00", "03:00-11:00"],
        "daily_targets": {"19:00-27:00": 1, "11:00-19:00": 2, "03:00-11:00": 2},
        "saturday_targets": {"19:00-27:00": 1, "11:00-19:00": 2, "03:00-11:00": 2},
        "sunday_targets": {"19:00-27:00": 1, "11:00-19:00": 2, "03:00-11:00": 2},
    },
    "Auxiliar de Alistamiento": {
        "rotation_pool": ["22:00-30:00", "20:00-28:00", "08:00-17:00"],
        "daily_targets": {"22:00-30:00": 1, "20:00-28:00": 1, "08:00-17:00": 1},
        "saturday_targets": {"22:00-30:00": 1, "20:00-28:00": 1, "08:00-17:00": 1},
        "sunday_targets": {"22:00-30:00": 1, "20:00-28:00": 1, "08:00-17:00": 1},
    },
    "Conductor de Operaciones": {
        "rotation_pool": ["11:00-18:00", "03:00-10:00"],
        "daily_targets": {"11:00-18:00": 1, "03:00-10:00": 1},
        "saturday_targets": {"11:00-18:00": 1, "03:00-10:00": 1},
        "sunday_targets": {"11:00-18:00": 1, "03:00-10:00": 1},
    },
    "Personal de Mantenimiento": {
        "rotation_pool": ["08:00-17:00"],
        "daily_targets": {"08:00-17:00": 1},
        "saturday_targets": {"08:00-17:00": 1},
        "sunday_targets": {"08:00-17:00": 1},
    },
}

DEFAULT_RULES = {
    "min_rest_between": 12,
    "min_rest_off": 24,
    "min_work_between_off": 5,
    "max_consecutive": 7,
    "rotate_after_off": True,
    "auto_assign_at": True,
}

SEED_EMPLOYEES = [
    {"id": "12896", "codigo": "12896", "nombre": "Alexis Mojica", "cargo": "Analista de Operaciones", "puede_at": True, "activo": True},
    {"id": "13563", "codigo": "13563", "nombre": "Kathleen Batista", "cargo": "Analista de Operaciones", "puede_at": True, "activo": True},
    {"id": "12520", "codigo": "12520", "nombre": "Keily Agrazal", "cargo": "Analista de Operaciones", "puede_at": True, "activo": True},
    {"id": "3332", "codigo": "3332", "nombre": "Kevin Pitty", "cargo": "Analista de Operaciones", "puede_at": False, "activo": True},
    {"id": "669", "codigo": "669", "nombre": "Roberto Velarde", "cargo": "Analista de Operaciones", "puede_at": False, "activo": True},
    {"id": "4211", "codigo": "4211", "nombre": "Dariel Cardales", "cargo": "Auxiliar de Operaciones", "puede_at": True, "activo": True},
    {"id": "3393", "codigo": "3393", "nombre": "David Vega", "cargo": "Auxiliar de Operaciones", "puede_at": True, "activo": True},
    {"id": "13488", "codigo": "13488", "nombre": "Edwin Aranda", "cargo": "Auxiliar de Operaciones", "puede_at": True, "activo": True},
    {"id": "12717", "codigo": "12717", "nombre": "Francisco Delgado", "cargo": "Auxiliar de Operaciones", "puede_at": True, "activo": True},
    {"id": "12644", "codigo": "12644", "nombre": "Idelfonso Sánchez", "cargo": "Auxiliar de Operaciones", "puede_at": False, "activo": True},
    {"id": "13580", "codigo": "13580", "nombre": "Juan Patiño", "cargo": "Auxiliar de Operaciones", "puede_at": False, "activo": True},
    {"id": "81569", "codigo": "81569", "nombre": "Kathleen Amat", "cargo": "Auxiliar de Operaciones", "puede_at": False, "activo": True},
    {"id": "10142", "codigo": "10142", "nombre": "Javier Caballero", "cargo": "Auxiliar de Operaciones", "puede_at": False, "activo": True},
    {"id": "4607", "codigo": "4607", "nombre": "Nelson Mojica", "cargo": "Auxiliar de Operaciones", "puede_at": True, "activo": True},
    {"id": "759", "codigo": "759", "nombre": "Daniel Solis", "cargo": "Operador Líder", "puede_at": False, "activo": True},
    {"id": "10043", "codigo": "10043", "nombre": "Gilberto Torrero", "cargo": "Operador Líder", "puede_at": False, "activo": True},
    {"id": "13479", "codigo": "13479", "nombre": "Edier Barrios", "cargo": "Técnico de Operaciones", "puede_at": False, "activo": True},
    {"id": "12560", "codigo": "12560", "nombre": "Estefany Villamil", "cargo": "Técnico de Operaciones", "puede_at": False, "activo": True},
    {"id": "11383", "codigo": "11383", "nombre": "Kathia Gonzalez", "cargo": "Técnico de Operaciones", "puede_at": False, "activo": True},
    {"id": "671", "codigo": "671", "nombre": "Luis Guardao", "cargo": "Técnico de Operaciones", "puede_at": False, "activo": True},
    {"id": "11889", "codigo": "11889", "nombre": "Daniel Barría", "cargo": "Auxiliar de Alistamiento", "puede_at": False, "activo": True},
    {"id": "13031", "codigo": "13031", "nombre": "Ismael Caballero", "cargo": "Auxiliar de Alistamiento", "puede_at": False, "activo": True},
    {"id": "9939", "codigo": "9939", "nombre": "Kevin Miranda", "cargo": "Auxiliar de Alistamiento", "puede_at": False, "activo": True},
    {"id": "11408", "codigo": "11408", "nombre": "Luis Castillo", "cargo": "Auxiliar de Alistamiento", "puede_at": False, "activo": True},
    {"id": "974", "codigo": "974", "nombre": "José De León", "cargo": "Conductor de Operaciones", "puede_at": False, "activo": True},
    {"id": "11824", "codigo": "11824", "nombre": "Raúl Marín", "cargo": "Conductor de Operaciones", "puede_at": False, "activo": True},
    {"id": "11200", "codigo": "11200", "nombre": "Alex González", "cargo": "Personal de Mantenimiento", "puede_at": False, "activo": True},
]

SEED_WEEK_RAW = {
    "12896": ["03:00-11:00 AT", "03:00-11:00 AT", "03:00-11:00 AT", "03:00-11:00", "03:00-11:00 AT", "03:00-11:00", "L"],
    "13563": ["VACACIONES", "VACACIONES", "VACACIONES", "VACACIONES", "VACACIONES", "L", "03:00-11:00"],
    "12520": ["11:00-19:00", "11:00-19:00", "11:00-19:00", "11:00-19:00", "L", "19:00-27:00", "19:00-27:00"],
    "3332": ["19:00-27:00", "19:00-27:00", "19:00-27:00", "19:00-27:00", "19:00-27:00", "L", "11:00-19:00"],
    "669": ["L", "06:00-15:00", "06:00-15:00", "06:00-15:00 CAP", "08:00-17:00", "11:00-19:00", "08:00-17:00"],
    "4211": ["03:00-11:00", "03:00-11:00", "03:00-11:00", "03:00-11:00 AT", "SUSPEND", "03:00-11:00", "L"],
    "3393": ["L", "SUSPEND", "11:00-19:00", "11:00-19:00", "11:00-27:00", "11:00-19:00", "11:00-19:00 AT"],
    "13488": ["11:00-19:00", "11:00-19:00", "L", "03:00-11:00", "03:00-11:00", "19:00-27:00", "19:00-27:00"],
    "12717": ["VACACIONES", "VACACIONES", "VACACIONES", "VACACIONES", "VACACIONES", "03:00-11:00 AT", "03:00-11:00"],
    "12644": ["03:00-11:00", "03:00-11:00", "03:00-11:00", "03:00-11:00", "03:00-11:00", "L", "03:00-11:00"],
    "13580": ["19:00-27:00", "19:00-27:00", "19:00-27:00", "19:00-27:00", "L", "23:00-35:00", "23:00-31:00"],
    "81569": ["23:00-31:00", "23:00-31:00", "23:00-31:00", "23:00-31:00", "23:00-31:00", "L", "11:00-19:00"],
    "10142": ["23:00-31:00", "23:00-31:00", "23:00-31:00", "23:00-31:00", "23:00-31:00", "L", "23:00-31:00"],
    "4607": ["11:00-19:00", "11:00-19:00", "11:00-19:00", "11:00-19:00 AT/CAP", "11:00-19:00", "11:00-19:00 AT", "L"],
    "759": ["20:00-28:00", "20:00-28:00", "20:00-28:00", "20:00-28:00", "20:00-28:00", "L", "20:00-28:00"],
    "10043": ["20:00-28:00", "20:00-28:00", "20:00-28:00", "20:00-28:00", "L", "20:00-28:00", "20:00-28:00"],
    "13479": ["11:00-19:00", "11:00-19:00", "11:00-19:00", "L", "03:00-19:00", "03:00-11:00", "03:00-11:00"],
    "12560": ["19:00-27:00", "19:00-27:00", "19:00-27:00", "19:00-27:00", "19:00-27:00", "L", "11:00-19:00"],
    "11383": ["11:00-19:00", "11:00-19:00", "11:00-19:00", "11:00-19:00", "11:00-19:00", "11:00-19:00", "L"],
    "671": ["03:00-11:00", "03:00-11:00", "03:00-11:00 CAP", "03:00-11:00", "L", "19:00-27:00", "19:00-27:00"],
    "11889": ["20:00-28:00", "20:00-28:00", "20:00-28:00", "20:00-28:00", "L", "22:00-30:00", "22:00-30:00"],
    "13031": ["22:00-30:00", "22:00-30:00", "22:00-30:00", "22:00-30:00", "22:00-30:00", "L", "08:00-17:00"],
    "9939": ["08:00-17:00", "08:00-17:00", "08:00-17:00", "08:00-17:00", "08:00-17:00", "L", "08:00-17:00"],
    "11408": ["20:00-28:00", "20:00-28:00", "20:00-28:00", "20:00-28:00", "20:00-28:00", "20:00-28:00", "L"],
    "974": ["03:00-10:00", "03:00-10:00", "03:00-10:00", "03:00-10:00", "03:00-10:00", "L", "11:00-18:00"],
    "11824": ["11:00-18:00", "11:00-18:00", "11:00-18:00", "11:00-18:00", "L", "03:00-10:00", "03:00-10:00"],
    "11200": ["08:00-17:00", "08:00-17:00", "08:00-17:00", "08:00-17:00", "08:00-17:00", "08:00-17:00", "L"],
}


def strip_accents(text):
    return "".join(c for c in unicodedata.normalize("NFD", str(text)) if unicodedata.category(c) != "Mn")


def current_monday_iso():
    today = dt.date.today()
    mon = today - dt.timedelta(days=today.weekday())
    return mon.strftime("%Y-%m-%d")


def format_short_date(iso_str):
    d = dt.date.fromisoformat(iso_str)
    return f"{d.day}-{MONTHS_ES[d.month - 1]}"


def add_days_iso(iso_str, days):
    d = dt.date.fromisoformat(iso_str) + dt.timedelta(days=days)
    return d.strftime("%Y-%m-%d")


def week_label_from_monday(monday_iso):
    return f"Semana {format_short_date(monday_iso)} — {format_short_date(add_days_iso(monday_iso, 6))}"


def parse_raw_cell(raw_val, locked=False):
    clean = re.sub(r"\s+", " ", str(raw_val or "").strip())
    upper = clean.upper()
    if not clean or upper in ("L", "LIBRE"):
        return {
            "raw": "L" if upper in ("L", "LIBRE") else "",
            "base": "L" if upper in ("L", "LIBRE") else "",
            "is_at": False,
            "is_cap": False,
            "start_min": None,
            "end_min": None,
            "locked": locked,
        }
    if "VAC" in upper:
        return {"raw": "VACACIONES", "base": "VACACIONES", "is_at": False, "is_cap": False, "start_min": None, "end_min": None, "locked": True}
    if "SUSP" in upper:
        return {"raw": "SUSPEND", "base": "SUSPEND", "is_at": False, "is_cap": False, "start_min": None, "end_min": None, "locked": True}
    if "INCAP" in upper:
        return {"raw": "INCAPACIDAD", "base": "INCAPACIDAD", "is_at": False, "is_cap": False, "start_min": None, "end_min": None, "locked": True}
    if "PERM" in upper:
        return {"raw": "PERMISO", "base": "PERMISO", "is_at": False, "is_cap": False, "start_min": None, "end_min": None, "locked": True}

    is_at = bool(re.search(r"\bAT\b", upper))
    is_cap = bool(re.search(r"\bCAP\b", upper))
    m = re.search(r"(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})", upper)
    if m:
        sh, sm, eh, em = int(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(4))
        base = f"{sh:02d}:{sm:02d}-{eh:02d}:{em:02d}"
        start_min = sh * 60 + sm
        end_min = eh * 60 + em
        if end_min <= start_min:
            end_min += 24 * 60
        return {"raw": clean, "base": base, "is_at": is_at, "is_cap": is_cap, "start_min": start_min, "end_min": end_min, "locked": locked}

    return {"raw": clean, "base": clean, "is_at": is_at, "is_cap": is_cap, "start_min": None, "end_min": None, "locked": locked}


def format_cell_raw(base, is_at=False, is_cap=False):
    if not base:
        return ""
    if base in SPECIAL_STATES:
        return base
    tags = []
    if is_at:
        tags.append("AT")
    if is_cap:
        tags.append("CAP")
    return f"{base} {'/'.join(tags)}" if tags else base


def is_work_shift(cell):
    return bool(cell and cell.get("base") and cell["base"] not in SPECIAL_STATES and cell.get("start_min") is not None)


def is_absence_state(cell):
    return bool(cell and cell.get("base") in ("VACACIONES", "SUSPEND", "INCAPACIDAD", "PERMISO"))


def is_tech_no_at(shift_code):
    if not shift_code:
        return False
    m = re.match(r"^(\d{1,2}):(\d{2})", str(shift_code).strip())
    if not m:
        return False
    return int(m.group(1)) == 19 and int(m.group(2)) == 0


def covers_shift_window(cell, target_shift):
    if not is_work_shift(cell):
        return False
    if cell["base"] == target_shift:
        return True
    parsed = parse_raw_cell(target_shift)
    if parsed["start_min"] is None or parsed["end_min"] is None:
        return False
    return cell["start_min"] <= parsed["start_min"] and cell["end_min"] >= parsed["end_min"]


def get_all_cargo_shifts(cfg):
    if not cfg:
        return []
    return list(dict.fromkeys(
        cfg.get("rotation_pool", [])
        + list(cfg.get("daily_targets", {}).keys())
        + list(cfg.get("saturday_targets", {}).keys())
        + list(cfg.get("sunday_targets", {}).keys())
    ))


def get_shift_target_for_day(cfg, shift_code, day_idx):
    if not cfg:
        return 0
    if day_idx == 5:
        sat = cfg.get("saturday_targets")
        if sat is not None and shift_code in sat:
            return int(sat.get(shift_code, 0))
    if day_idx == 6:
        sun = cfg.get("sunday_targets")
        if sun is not None and shift_code in sun:
            return int(sun.get(shift_code, 0))
    return int(cfg.get("daily_targets", {}).get(shift_code, 0))


def get_total_cargo_target_for_day(cfg, day_idx):
    if not cfg:
        return 0
    return sum(get_shift_target_for_day(cfg, s, day_idx) for s in get_all_cargo_shifts(cfg))


def get_enabled_pool(cfg):
    pool = cfg.get("rotation_pool", []) if cfg else []
    enabled = [
        s for s in pool
        if any(get_shift_target_for_day(cfg, s, d) > 0 for d in range(7))
    ]
    return enabled if enabled else pool


def get_cell_colors(cell):
    base = (cell or {}).get("base", "")
    is_at = (cell or {}).get("is_at", False)
    if not base:
        return ("#FFFFFF", "#94A3B8", "#CBD5E1")
    if base == "L":
        return ("#C6B39A", "#800000", "#A89378")
    if base == "VACACIONES":
        return ("#FEF3C7", "#92400E", "#D97706")
    if base == "SUSPEND":
        return ("#DC2626", "#FFFFFF", "#991B1B")
    if base in ("INCAPACIDAD", "PERMISO"):
        return ("#E2E8F0", "#1E293B", "#94A3B8")
    if is_at:
        return ("#C850B0", "#000000", "#9E3889")
    if base.startswith("03:"):
        return ("#70AD47", "#800000", "#558733")
    if base.startswith(("06:", "07:")):
        return ("#B6C3B4", "#800000", "#92A390")
    if base.startswith("08:"):
        return ("#D9D9D9", "#000000", "#AEAEAE")
    if base.startswith("11:"):
        return ("#CC9900", "#800000", "#9E7600")
    if base.startswith("13:"):
        return ("#C850B0", "#800000", "#9E3889")
    if base.startswith("19:"):
        return ("#8E9FB4", "#800000", "#6C7F96")
    if base.startswith("20:"):
        return ("#25A9D8", "#800000", "#1885AD")
    if base.startswith("22:"):
        return ("#949494", "#7A0000", "#6E6E6E")
    if base.startswith("23:"):
        return ("#B6BE5F", "#800000", "#8E9640")

    start_min = (cell or {}).get("start_min")
    if start_min is not None:
        start_h = start_min // 60
        if start_h < 5:
            return ("#70AD47", "#800000", "#558733")
        if start_h < 8:
            return ("#B6C3B4", "#800000", "#92A390")
        if start_h < 10:
            return ("#D9D9D9", "#000000", "#AEAEAE")
        if start_h < 13:
            return ("#CC9900", "#800000", "#9E7600")
        if start_h < 18:
            return ("#C850B0", "#800000", "#9E3889")
        if start_h < 20:
            return ("#8E9FB4", "#800000", "#6C7F96")
        if start_h < 22:
            return ("#25A9D8", "#800000", "#1885AD")
        if start_h == 22:
            return ("#949494", "#7A0000", "#6E6E6E")
        return ("#B6BE5F", "#800000", "#8E9640")

    return ("#E2E8F0", "#0F172A", "#94A3B8")


def get_novedad_badge_colors(tipo):
    if tipo == "VACACIONES":
        return ("#FEF3C7", "#92400E", "#FCD34D")
    if tipo == "SUSPEND":
        return ("#FEE2E2", "#991B1B", "#FCA5A5")
    if tipo == "L":
        return ("#FCE4D6", "#9A3412", "#FDBA74")
    if tipo == "INCAPACIDAD":
        return ("#F3E8FF", "#6B21A8", "#D8B4FE")
    return ("#E0F2FE", "#075985", "#7DD3FC")


def seed_week():
    assignments = {eid: [parse_raw_cell(x) for x in row] for eid, row in SEED_WEEK_RAW.items()}
    return {
        "id": "week_2026-10-26",
        "start_date": "2026-10-26",
        "label": week_label_from_monday("2026-10-26"),
        "assignments": assignments,
    }


def init_state():
    if "employees" not in st.session_state:
        st.session_state["employees"] = copy.deepcopy(SEED_EMPLOYEES)
    if "weeks" not in st.session_state:
        st.session_state["weeks"] = []
    if "active_week_idx" not in st.session_state:
        st.session_state["active_week_idx"] = 0
    if "role_configs" not in st.session_state:
        st.session_state["role_configs"] = copy.deepcopy(DEFAULT_ROLE_CONFIGS)
    if "rules" not in st.session_state:
        st.session_state["rules"] = copy.deepcopy(DEFAULT_RULES)
    if "scheduled_novedades" not in st.session_state:
        st.session_state["scheduled_novedades"] = []
    if "rerotate_offset" not in st.session_state:
        st.session_state["rerotate_offset"] = 0


def sync_widget_state_to_configs():
    rules = st.session_state["rules"]
    for rk, wk in [
        ("min_rest_between", "t4_min_rest_between"),
        ("min_rest_off", "t4_min_rest_off"),
        ("min_work_between_off", "t4_min_work"),
        ("max_consecutive", "t4_max_cons"),
        ("rotate_after_off", "t4_rot_off"),
        ("auto_assign_at", "t4_auto_at"),
    ]:
        if wk in st.session_state:
            rules[rk] = st.session_state[wk]


def apply_novedad_to_week(week, nov):
    start_d = dt.date.fromisoformat(week["start_date"])
    nov_start = dt.date.fromisoformat(nov["start_date"])
    nov_end = dt.date.fromisoformat(nov["end_date"])
    new_week = copy.deepcopy(week)
    applied = 0
    for d in range(7):
        cur = start_d + dt.timedelta(days=d)
        if nov_start <= cur <= nov_end:
            row = new_week["assignments"].setdefault(nov["emp_id"], [parse_raw_cell("") for _ in range(7)])
            row[d] = parse_raw_cell(nov["tipo"], locked=True)
            applied += 1
    return new_week, applied


def get_rest_hours_gap(prev_cell, next_cell):
    if not is_work_shift(prev_cell) or not is_work_shift(next_cell):
        return None
    return round((next_cell["start_min"] + 1440 - prev_cell["end_min"]) / 60.0, 1)


def get_rest_hours_across_day_off(prev_cell, next_cell):
    if not is_work_shift(prev_cell) or not is_work_shift(next_cell):
        return None
    return round((next_cell["start_min"] + 2880 - prev_cell["end_min"]) / 60.0, 1)


def is_transition_valid(prev_shift_str, prev_offset, cand_str, cand_day, crossed_off, rules):
    if not prev_shift_str:
        return True
    wp = parse_raw_cell(prev_shift_str)
    wn = parse_raw_cell(cand_str)
    if wp["start_min"] is None or wn["start_min"] is None:
        return True
    rest_h = (cand_day * 1440 + wn["start_min"] - (prev_offset * 1440 + wp["end_min"])) / 60.0
    if crossed_off:
        return rest_h >= rules["min_rest_off"]
    return rest_h >= rules["min_rest_between"]


def check_range_fit(row, prev_row, start_d, end_d, cand_shift, rules):
    prev_s = None
    prev_off = start_d - 1
    prev_crossed = False
    for k in range(start_d - 1, -1, -1):
        c = row[k]
        if c.get("base") == "L":
            prev_crossed = True
            continue
        if is_work_shift(c):
            prev_s = c["base"]
            prev_off = k
            break
        break
    if not prev_s and prev_row and len(prev_row) == 7:
        for k in range(6, -1, -1):
            c = prev_row[k]
            if c.get("base") == "L":
                prev_crossed = True
                continue
            if is_work_shift(c):
                prev_s = c["base"]
                prev_off = k - 7
                break
            break
    if not is_transition_valid(prev_s, prev_off, cand_shift, start_d, prev_crossed, rules):
        return False

    next_s = None
    next_off = end_d + 1
    next_crossed = False
    for k in range(end_d + 1, 7):
        c = row[k]
        if c.get("base") == "L":
            next_crossed = True
            continue
        if is_work_shift(c):
            next_s = c["base"]
            next_off = k
            break
        break
    if next_s and not is_transition_valid(cand_shift, end_d, next_s, next_off, next_crossed, rules):
        return False
    return True


def preserves_rotation_across_off(row, prev_row, start_d, end_d, cand_shift, rules):
    if not rules.get("rotate_after_off", True):
        return True

    saw_off_before = False
    shift_before_off = None
    for k in range(start_d - 1, -1, -1):
        c = row[k]
        if c.get("base") == "L":
            saw_off_before = True
            continue
        if is_work_shift(c):
            if saw_off_before:
                shift_before_off = c["base"]
            break
        break
    if not shift_before_off and saw_off_before and prev_row and len(prev_row) == 7:
        for k in range(6, -1, -1):
            c = prev_row[k]
            if c.get("base") == "L":
                saw_off_before = True
                continue
            if is_work_shift(c):
                shift_before_off = c["base"]
                break
            break
    if start_d == 0 and prev_row and len(prev_row) == 7 and prev_row[6].get("base") == "L":
        for k in range(5, -1, -1):
            if is_work_shift(prev_row[k]):
                shift_before_off = prev_row[k]["base"]
                break
    if shift_before_off and shift_before_off == cand_shift:
        return False

    saw_off_after = False
    shift_after_off = None
    for k in range(end_d + 1, 7):
        c = row[k]
        if c.get("base") == "L":
            saw_off_after = True
            continue
        if is_work_shift(c):
            if saw_off_after:
                shift_after_off = c["base"]
            break
        break
    if saw_off_after and shift_after_off and shift_after_off == cand_shift:
        return False

    return True


def compute_uncovered_tasks(week, employees, role_configs):
    uncovered = []
    active_emps = [e for e in employees if e["activo"]]
    for cargo in CARGOS_ORDEN:
        cfg = role_configs.get(cargo)
        if not cfg:
            continue
        shifts = get_all_cargo_shifts(cfg)
        cargo_emps = [e for e in active_emps if e["cargo"] == cargo]
        for d in range(7):
            for s in shifts:
                target = get_shift_target_for_day(cfg, s, d)
                if target <= 0:
                    continue
                if cargo == "Técnico de Operaciones":
                    if is_tech_no_at(s):
                        assigned = sum(1 for e in cargo_emps if covers_shift_window((week["assignments"].get(e["id"], [{}] * 7))[d], s))
                    else:
                        assigned = sum(
                            1 for e in active_emps
                            if covers_shift_window((week["assignments"].get(e["id"], [{}] * 7))[d], s)
                            and (e["cargo"] == "Técnico de Operaciones" or (week["assignments"].get(e["id"], [{}] * 7))[d].get("is_at"))
                        )
                else:
                    assigned = sum(1 for e in cargo_emps if covers_shift_window((week["assignments"].get(e["id"], [{}] * 7))[d], s))
                if assigned < target:
                    uncovered.append({
                        "cargo": cargo,
                        "day_idx": d,
                        "day_name": DAYS_ES[d],
                        "shift_code": s,
                        "target": target,
                        "assigned": assigned,
                        "missing": target - assigned,
                    })
    return uncovered


def get_daily_tech_coverage(week, employees, role_configs):
    tech_cfg = role_configs.get("Técnico de Operaciones", {})
    tech_shifts = get_all_cargo_shifts(tech_cfg)
    active_emps = [e for e in employees if e["activo"]]
    results = []
    for d in range(7):
        for s in tech_shifts:
            target = get_shift_target_for_day(tech_cfg, s, d)
            no_at = is_tech_no_at(s)
            direct_members = []
            at_members = []
            for e in active_emps:
                cell = (week["assignments"].get(e["id"], [{}] * 7))[d]
                if not covers_shift_window(cell, s):
                    continue
                if e["cargo"] == "Técnico de Operaciones":
                    direct_members.append({
                        "id": e["id"],
                        "codigo": e["codigo"],
                        "nombre": e["nombre"],
                        "cargo": e["cargo"],
                        "raw": cell.get("raw", ""),
                        "is_at": False,
                    })
                elif not no_at and cell.get("is_at"):
                    at_members.append({
                        "id": e["id"],
                        "codigo": e["codigo"],
                        "nombre": e["nombre"],
                        "cargo": e["cargo"],
                        "raw": cell.get("raw", ""),
                        "is_at": True,
                    })
            total_assigned = len(direct_members) + len(at_members)
            results.append({
                "day_idx": d,
                "day_name": DAYS_ES[d],
                "short_day": SHORT_DAYS_ES[d],
                "shift_code": s,
                "no_at": no_at,
                "target": target,
                "direct_members": direct_members,
                "at_members": at_members,
                "total_assigned": total_assigned,
                "missing": max(0, target - total_assigned),
            })
    return results


def validate_schedule(week, employees, role_configs, rules, prev_week=None):
    issues = []
    active_emps = [e for e in employees if e["activo"]]
    for emp in active_emps:
        row = week["assignments"].get(emp["id"], [])
        if len(row) < 7:
            continue
        prev_sun = prev_week["assignments"].get(emp["id"], [None] * 7)[6] if prev_week else None
        for d in range(7):
            cur = row[d]
            prev = prev_sun if d == 0 else row[d - 1]
            if cur.get("base") == "L" and d < 6 and is_work_shift(row[d + 1]) and is_work_shift(prev):
                g = get_rest_hours_across_day_off(prev, row[d + 1])
                if g is not None and g < rules["min_rest_off"]:
                    issues.append({
                        "type": "REST_DAY_OFF",
                        "emp_id": emp["id"],
                        "day_idx": d,
                        "msg": f"{DAYS_ES[d]}: {emp['nombre']} ({emp['codigo']}): {g}h en Día Libre (mín. {rules['min_rest_off']}h).",
                    })
            if is_work_shift(cur) and is_work_shift(prev):
                g = get_rest_hours_gap(prev, cur)
                if g is not None and g < rules["min_rest_between"]:
                    issues.append({
                        "type": "REST_12H",
                        "emp_id": emp["id"],
                        "day_idx": d,
                        "msg": f"{DAYS_ES[d]}: {emp['nombre']} ({emp['codigo']}): {g}h entre {prev['base']} y {cur['base']} (mín. {rules['min_rest_between']}h).",
                    })

    tech_cfg = role_configs.get("Técnico de Operaciones", {})
    tech_shifts = get_all_cargo_shifts(tech_cfg)
    for d in range(7):
        for s in tech_shifts:
            target = get_shift_target_for_day(tech_cfg, s, d)
            if target <= 0:
                continue
            no_at = is_tech_no_at(s)
            covered = 0
            for e in active_emps:
                c = week["assignments"].get(e["id"], [{}] * 7)[d]
                if not covers_shift_window(c, s):
                    continue
                if e["cargo"] == "Técnico de Operaciones" or (not no_at and c.get("is_at")):
                    covered += 1
            if covered < target:
                issues.append({
                    "type": "UNCOVERED_TECH",
                    "emp_id": None,
                    "day_idx": d,
                    "msg": f"{DAYS_ES[d]}: Faltan {target - covered} Técnico(s) en {s}" + (" (solo Técnico directo, sin AT)." if no_at else "."),
                })
    return issues


def auto_assign_at(week, employees, role_configs, rules, prev_week=None):
    new_week = copy.deepcopy(week)
    active_emps = [e for e in employees if e["activo"]]
    for e in active_emps:
        if e["cargo"] == "Técnico de Operaciones":
            continue
        row = new_week["assignments"].get(e["id"], [])
        for d in range(min(7, len(row))):
            if row[d].get("is_at") and not row[d].get("locked"):
                row[d] = parse_raw_cell(format_cell_raw(row[d]["base"], is_at=False, is_cap=row[d].get("is_cap", False)))

    tech_cfg = role_configs.get("Técnico de Operaciones", {})
    tech_shifts = get_all_cargo_shifts(tech_cfg)
    candidates = [e for e in active_emps if e["puede_at"] and e["cargo"] in ("Analista de Operaciones", "Auxiliar de Operaciones")]
    assigned_cnt = 0

    for d in range(7):
        for s in tech_shifts:
            if is_tech_no_at(s):
                continue
            target = get_shift_target_for_day(tech_cfg, s, d)
            if target <= 0:
                continue
            covered = 0
            for e in active_emps:
                c = new_week["assignments"].get(e["id"], [{}] * 7)[d]
                if covers_shift_window(c, s) and (e["cargo"] == "Técnico de Operaciones" or c.get("is_at")):
                    covered += 1
            while covered < target:
                same_cands = []
                for e in candidates:
                    row = new_week["assignments"].get(e["id"], [])
                    if len(row) < 7:
                        continue
                    c = row[d]
                    if c.get("locked") or is_absence_state(c) or c.get("base") == "L" or c.get("is_at"):
                        continue
                    if c.get("base") == s:
                        at_load = sum(1 for x in row if x.get("is_at"))
                        same_cands.append((at_load, 0 if e["cargo"] == "Auxiliar de Operaciones" else 1, e))

                if same_cands:
                    same_cands.sort(key=lambda t: (t[0], t[1]))
                    chosen = same_cands[0][2]
                    row = new_week["assignments"][chosen["id"]]
                    row[d] = parse_raw_cell(format_cell_raw(s, is_at=True, is_cap=row[d].get("is_cap", False)), locked=False)
                    covered += 1
                    assigned_cnt += 1
                    continue

                flex_cands = []
                for e in candidates:
                    row = new_week["assignments"].get(e["id"], [])
                    if len(row) < 7:
                        continue
                    c = row[d]
                    if c.get("locked") or is_absence_state(c) or c.get("base") == "L" or c.get("is_at") or not is_work_shift(c):
                        continue
                    own_shift = c["base"]
                    own_target = get_shift_target_for_day(role_configs.get(e["cargo"], {}), own_shift, d)
                    own_count = sum(
                        1 for emp2 in active_emps
                        if emp2["cargo"] == e["cargo"]
                        and new_week["assignments"].get(emp2["id"], [{}] * 7)[d].get("base") == own_shift
                    )
                    if own_count <= own_target:
                        continue
                    prev_row = prev_week["assignments"].get(e["id"]) if prev_week else None
                    if not check_range_fit(row, prev_row, d, d, s, rules):
                        continue
                    if not preserves_rotation_across_off(row, prev_row, d, d, s, rules):
                        continue
                    at_load = sum(1 for x in row if x.get("is_at"))
                    flex_cands.append((-(own_count - own_target), at_load, 0 if e["cargo"] == "Auxiliar de Operaciones" else 1, e))

                if flex_cands:
                    flex_cands.sort(key=lambda t: (t[0], t[1], t[2]))
                    chosen = flex_cands[0][3]
                    row = new_week["assignments"][chosen["id"]]
                    row[d] = parse_raw_cell(format_cell_raw(s, is_at=True, is_cap=row[d].get("is_cap", False)), locked=False)
                    covered += 1
                    assigned_cnt += 1
                    continue
                break
    return new_week, assigned_cnt


def generate_next_week(
    base_week,
    employees,
    role_configs,
    rules,
    override_start_iso=None,
    existing_target=None,
    scheduled_novedades=None,
    is_same_week_baseline=False,
    seed_offset=0,
):
    next_start = override_start_iso or add_days_iso(base_week["start_date"], 7)
    next_end_dt = dt.date.fromisoformat(next_start) + dt.timedelta(days=6)
    next_start_dt = dt.date.fromisoformat(next_start)
    seed_offset = max(0, int(seed_offset or 0))

    new_week = {
        "id": f"week_{next_start}",
        "start_date": next_start,
        "label": week_label_from_monday(next_start),
        "assignments": {},
    }

    for emp in employees:
        eid = emp["id"]
        if existing_target and eid in existing_target["assignments"]:
            row = [
                copy.deepcopy(c) if (c.get("locked") or is_absence_state(c)) else parse_raw_cell("")
                for c in existing_target["assignments"][eid]
            ]
        else:
            row = [parse_raw_cell("") for _ in range(7)]

        if scheduled_novedades:
            for nov in scheduled_novedades:
                if nov["emp_id"] != eid:
                    continue
                ns = dt.date.fromisoformat(nov["start_date"])
                ne = dt.date.fromisoformat(nov["end_date"])
                if ne < next_start_dt or ns > next_end_dt:
                    continue
                for d in range(7):
                    cur_d = next_start_dt + dt.timedelta(days=d)
                    if ns <= cur_d <= ne:
                        row[d] = parse_raw_cell(nov["tipo"], locked=True)
        new_week["assignments"][eid] = row

    def get_rotated_shift(cargo, current_shift, step=1):
        pool = get_enabled_pool(role_configs.get(cargo, {}))
        if not pool:
            return current_shift
        if current_shift in pool:
            return pool[(pool.index(current_shift) + step) % len(pool)]
        for p in pool:
            if p != current_shift:
                return p
        return pool[0]

    def count_shift(cargo, day_idx, shift_code):
        cnt = 0
        for emp in employees:
            if not emp["activo"] or emp["cargo"] != cargo:
                continue
            r = new_week["assignments"].get(emp["id"], [])
            if len(r) == 7 and r[day_idx].get("base") == shift_code:
                cnt += 1
        return cnt

    def count_active_cargo(cargo, day_idx):
        cnt = 0
        for emp in employees:
            if not emp["activo"] or emp["cargo"] != cargo:
                continue
            r = new_week["assignments"].get(emp["id"], [])
            if len(r) == 7 and is_work_shift(r[day_idx]):
                cnt += 1
        return cnt

    def pick_best_shift(cargo, day_indices, prev_row, row, preferred, avoid=None, strict_avoid=False, tie_breaker_offset=0):
        cfg = role_configs.get(cargo, {})
        pool = get_enabled_pool(cfg)
        if not pool or not day_indices:
            return preferred

        def fits_all(s):
            sim = copy.deepcopy(row)
            for d in day_indices:
                if sim[d].get("locked") or is_absence_state(sim[d]) or sim[d].get("base") == "L":
                    continue
                if not check_range_fit(sim, prev_row, d, d, s, rules):
                    return False
                sim[d] = parse_raw_cell(s)
            return True

        if strict_avoid and avoid and any(s != avoid for s in pool):
            cand_pool = [s for s in pool if s != avoid]
        else:
            cand_pool = pool

        valid_strict = [s for s in cand_pool if fits_all(s)]
        valid_any = valid_strict if valid_strict else [s for s in pool if fits_all(s)]
        base_valid = valid_any if valid_any else cand_pool
        rot_shift = tie_breaker_offset % len(base_valid) if tie_breaker_offset > 0 else 0
        valid = base_valid[rot_shift:] + base_valid[:rot_shift] if rot_shift else base_valid

        best_s = valid[0]
        best_score = -10**9
        for s in valid:
            score = 0
            for d in day_indices:
                t = get_shift_target_for_day(cfg, s, d)
                cur = count_shift(cargo, d, s)
                if t <= 0:
                    score -= 160
                elif cur < t:
                    score += 260 + (t - cur) * 50
                    if cargo == "Técnico de Operaciones" and is_tech_no_at(s):
                        score += 450
                else:
                    score -= 90 * (cur - t + 1)
            if s == preferred:
                score += 120
            if avoid and s == avoid:
                score -= 2000 if strict_avoid else 180
            if cargo == "Técnico de Operaciones" and is_tech_no_at(s):
                score += 25
            if score > best_score:
                best_score = score
                best_s = s
        return best_s

    for cargo in CARGOS_ORDEN:
        cfg = role_configs.get(cargo, {})
        pool = get_enabled_pool(cfg)
        if not pool:
            continue
        base_cargo_emps = [e for e in employees if e["activo"] and e["cargo"] == cargo]
        if not base_cargo_emps:
            continue
        emp_rot = seed_offset % len(base_cargo_emps)
        cargo_emps = base_cargo_emps[emp_rot:] + base_cargo_emps[:emp_rot] if emp_rot else base_cargo_emps

        min_w = rules["min_work_between_off"]
        max_c = rules["max_consecutive"]

        for emp_idx, emp in enumerate(cargo_emps):
            eid = emp["id"]
            base_row = base_week["assignments"].get(eid, [parse_raw_cell("") for _ in range(7)])
            prev_row = None if is_same_week_baseline else base_row
            row = new_week["assignments"][eid]
            if sum(1 for c in row if is_absence_state(c)) >= 5:
                continue

            prev_sun = prev_row[6] if (prev_row and len(prev_row) == 7) else parse_raw_cell("")
            if prev_row:
                last_work = next((c["base"] for c in reversed(prev_row) if is_work_shift(c)), pool[(emp_idx + seed_offset) % len(pool)])
            else:
                last_work = next((c["base"] for c in base_row if is_work_shift(c)), pool[(emp_idx + seed_offset) % len(pool)])

            streak = 0
            if prev_row:
                for c in reversed(prev_row):
                    if is_work_shift(c):
                        streak += 1
                    else:
                        break

            locked_off_idx = next((i for i, c in enumerate(row) if c.get("locked") and c.get("base") == "L"), None)
            if locked_off_idx is not None:
                off_day = locked_off_idx
            else:
                has_abs_l = any(is_absence_state(c) or c.get("base") == "L" for c in row)
                if has_abs_l:
                    off_day = None
                else:
                    existing_off = next((i for i, c in enumerate(base_row) if c.get("base") == "L"), -1) if is_same_week_baseline else -1
                    if is_same_week_baseline and existing_off != -1 and seed_offset == 0:
                        off_day = existing_off
                    else:
                        if prev_row:
                            ideal = max(0, min(6, max_c - 1 - streak))
                            min_allowed = max(0, min_w - streak)
                            max_allowed = min(6, max_c - streak)
                        else:
                            ideal = (existing_off + seed_offset) % 7 if existing_off != -1 else (emp_idx + seed_offset) % 7
                            min_allowed = 0
                            max_allowed = 6

                        candidates = [d for d in range(min_allowed, max_allowed + 1) if not row[d].get("locked") and not is_absence_state(row[d])]
                        if not candidates:
                            candidates = [d for d in range(7) if not row[d].get("locked") and not is_absence_state(row[d])]
                        if candidates:
                            total_cargo_cap = len(cargo_emps)
                            off_day = candidates[0]
                            best_off_score = -10**9
                            for cd in candidates:
                                sc = 0
                                day_target_sum = get_total_cargo_target_for_day(cfg, cd)
                                active_already = count_active_cargo(cargo, cd)
                                if total_cargo_cap - 1 < day_target_sum or active_already <= day_target_sum:
                                    sc -= 700
                                sc -= abs(cd - ideal) * 20
                                sc -= sum(1 for o in cargo_emps if new_week["assignments"][o["id"]][cd].get("base") == "L") * 120
                                if cd == 5 and day_target_sum >= total_cargo_cap:
                                    sc -= 40
                                if is_same_week_baseline and seed_offset > 0 and cd == existing_off:
                                    sc -= 15
                                if sc > best_off_score:
                                    best_off_score = sc
                                    off_day = cd
                        else:
                            off_day = None

            if off_day is not None and not row[off_day].get("locked"):
                row[off_day] = parse_raw_cell("L")

            if rules["rotate_after_off"] and off_day is not None:
                pre_days = [d for d in range(0, off_day) if not row[d].get("locked") and not is_absence_state(row[d]) and row[d].get("base") != "L"]
                pre_shift = last_work
                if pre_days:
                    if is_same_week_baseline:
                        pref_start = get_rotated_shift(cargo, last_work, seed_offset) if seed_offset > 0 else (last_work if last_work in pool else pool[0])
                        pre_shift = pick_best_shift(cargo, pre_days, None, row, pref_start, tie_breaker_offset=seed_offset)
                    elif prev_sun.get("base") == "L":
                        rot_mon = get_rotated_shift(cargo, last_work, 1 + (seed_offset % max(1, len(pool) - 1)))
                        pre_shift = pick_best_shift(cargo, pre_days, prev_row, row, rot_mon, avoid=last_work, strict_avoid=True, tie_breaker_offset=seed_offset)
                    else:
                        sunday_work_shift = prev_sun["base"] if is_work_shift(prev_sun) else last_work
                        sunday_net_need = sum(
                            (1 if count_shift(cargo, d, sunday_work_shift) < get_shift_target_for_day(cfg, sunday_work_shift, d) else -1)
                            for d in pre_days
                        )
                        if (
                            sunday_work_shift in pool
                            and sunday_net_need > 0
                            and all(check_range_fit(row, prev_row, d, d, sunday_work_shift, rules) for d in pre_days)
                        ):
                            pre_shift = sunday_work_shift
                        else:
                            pre_shift = pick_best_shift(cargo, pre_days, prev_row, row, sunday_work_shift, tie_breaker_offset=seed_offset)
                    for d in pre_days:
                        row[d] = parse_raw_cell(pre_shift)

                post_days = [d for d in range(off_day + 1, 7) if not row[d].get("locked") and not is_absence_state(row[d]) and row[d].get("base") != "L"]
                if post_days:
                    base_before_l = pre_shift if pre_days else last_work
                    rot_step = 1 + (seed_offset % max(1, len(pool) - 1))
                    rot_candidate = get_rotated_shift(cargo, base_before_l, rot_step)
                    post_shift = pick_best_shift(cargo, post_days, prev_row, row, rot_candidate, avoid=base_before_l, strict_avoid=True, tie_breaker_offset=seed_offset)
                    for d in post_days:
                        row[d] = parse_raw_cell(post_shift)
            else:
                work_days = [d for d in range(7) if not row[d].get("locked") and not is_absence_state(row[d]) and row[d].get("base") != "L"]
                rot_candidate = get_rotated_shift(cargo, last_work, 1 + (seed_offset % max(1, len(pool))))
                chosen = pick_best_shift(cargo, work_days, prev_row, row, rot_candidate, avoid=(last_work if rules["rotate_after_off"] else None), strict_avoid=rules["rotate_after_off"], tie_breaker_offset=seed_offset)
                for d in work_days:
                    row[d] = parse_raw_cell(chosen)

        shifts_all = get_all_cargo_shifts(cfg)
        for d in range(7):
            under = [s for s in shifts_all if get_shift_target_for_day(cfg, s, d) > 0 and count_shift(cargo, d, s) < get_shift_target_for_day(cfg, s, d)]
            for u_shift in under:
                while count_shift(cargo, d, u_shift) < get_shift_target_for_day(cfg, u_shift, d):
                    moved = False
                    over = sorted(
                        [s for s in shifts_all if s != u_shift and count_shift(cargo, d, s) > get_shift_target_for_day(cfg, s, d)],
                        key=lambda s: -(count_shift(cargo, d, s) - get_shift_target_for_day(cfg, s, d)),
                    )
                    for o_shift in over:
                        for emp in cargo_emps:
                            r = new_week["assignments"][emp["id"]]
                            if r[d].get("locked") or r[d].get("base") != o_shift:
                                continue
                            end_d = d
                            while (
                                end_d + 1 < 7
                                and not r[end_d + 1].get("locked")
                                and r[end_d + 1].get("base") == o_shift
                                and count_shift(cargo, end_d + 1, u_shift) < get_shift_target_for_day(cfg, u_shift, end_d + 1)
                                and count_shift(cargo, end_d + 1, o_shift) > get_shift_target_for_day(cfg, o_shift, end_d + 1)
                            ):
                                end_d += 1
                            emp_prev_row = None if is_same_week_baseline else base_week["assignments"].get(emp["id"])
                            if end_d > d and check_range_fit(r, emp_prev_row, d, end_d, u_shift, rules) and preserves_rotation_across_off(r, emp_prev_row, d, end_d, u_shift, rules):
                                for dd in range(d, end_d + 1):
                                    r[dd] = parse_raw_cell(u_shift)
                                moved = True
                                break
                        if moved:
                            break
                    if not moved:
                        for o_shift in over:
                            for emp in cargo_emps:
                                r = new_week["assignments"][emp["id"]]
                                if r[d].get("locked") or r[d].get("base") != o_shift:
                                    continue
                                emp_prev_row = None if is_same_week_baseline else base_week["assignments"].get(emp["id"])
                                if check_range_fit(r, emp_prev_row, d, d, u_shift, rules) and preserves_rotation_across_off(r, emp_prev_row, d, d, u_shift, rules):
                                    r[d] = parse_raw_cell(u_shift)
                                    moved = True
                                    break
                            if moved:
                                break
                    if not moved:
                        break

    at_assigned = 0
    if rules.get("auto_assign_at", True):
        new_week, at_assigned = auto_assign_at(new_week, employees, role_configs, rules, None if is_same_week_baseline else base_week)
    return new_week, at_assigned


def trigger_rerotate_active_week():
    weeks = st.session_state["weeks"]
    if not weeks:
        return None, 0
    idx_a = min(st.session_state["active_week_idx"], len(weeks) - 1)
    cur_w = weeks[idx_a]
    prev_w = weeks[idx_a - 1] if idx_a > 0 else None
    next_offset = st.session_state.get("rerotate_offset", 0) + 1
    st.session_state["rerotate_offset"] = next_offset
    reg, at_c = generate_next_week(
        prev_w or cur_w,
        st.session_state["employees"],
        st.session_state["role_configs"],
        st.session_state["rules"],
        override_start_iso=cur_w["start_date"],
        existing_target=cur_w,
        scheduled_novedades=st.session_state["scheduled_novedades"],
        is_same_week_baseline=(prev_w is None),
        seed_offset=next_offset,
    )
    st.session_state["weeks"][idx_a] = reg
    return cur_w["label"], at_c


def analyze_rotation(curr_week, prev_week, employees, role_configs, rules):
    rows = []
    for emp in employees:
        if not emp["activo"]:
            continue
        curr_cells = curr_week["assignments"].get(emp["id"], [parse_raw_cell("") for _ in range(7)])
        prev_cells = prev_week["assignments"].get(emp["id"], [parse_raw_cell("") for _ in range(7)]) if prev_week else []
        prev_off = next((i for i in range(6, -1, -1) if prev_cells and prev_cells[i].get("base") == "L"), None)
        curr_off = next((i for i, c in enumerate(curr_cells) if c.get("base") == "L"), None)

        combined = (prev_cells + curr_cells) if prev_cells else curr_cells
        days_before_off = None
        if prev_cells and curr_off is not None and prev_off is not None:
            days_before_off = (6 - prev_off) + curr_off
        elif curr_off is not None:
            cnt = 0
            for i in range(curr_off - 1, -1, -1):
                if is_work_shift(curr_cells[i]):
                    cnt += 1
                else:
                    break
            days_before_off = cnt

        rest_hours = None
        if curr_off is not None:
            global_off = (len(prev_cells) + curr_off) if prev_cells else curr_off
            if 0 < global_off < len(combined) - 1:
                rest_hours = get_rest_hours_across_day_off(combined[global_off - 1], combined[global_off + 1])

        max_seq, run = 0, 0
        for c in combined:
            if is_work_shift(c):
                run += 1
                max_seq = max(max_seq, run)
            else:
                run = 0

        off_ok = True if days_before_off is None else (rules["min_work_between_off"] <= days_before_off <= rules["max_consecutive"])
        rest_ok = True if rest_hours is None else (rest_hours >= rules["min_rest_off"])

        rot_status, rot_detail = "NO_PREV", "Sin semana previa"
        if sum(1 for c in curr_cells if is_absence_state(c)) >= 5:
            rot_status, rot_detail = "ABSENCE", "Novedad (Vacaciones / Suspensión)"
        elif prev_cells:
            prev_last = next((c["base"] for c in reversed(prev_cells) if is_work_shift(c)), None)
            if curr_off is not None:
                before_l = next((curr_cells[i]["base"] for i in range(curr_off - 1, -1, -1) if is_work_shift(curr_cells[i])), prev_last)
                after_l = next((curr_cells[i]["base"] for i in range(curr_off + 1, 7) if is_work_shift(curr_cells[i])), None)
                if before_l and after_l and before_l != after_l:
                    rot_status, rot_detail = "ROTATED_AFTER_L", f"Rotó tras L ({DAYS_ES[curr_off]}): {before_l} → {after_l}"
                elif before_l and after_l and before_l == after_l:
                    rot_status, rot_detail = "SAME_SHIFT", f"Mantuvo {before_l} tras L ({DAYS_ES[curr_off]})"
                elif prev_last and before_l and prev_last != before_l:
                    rot_status, rot_detail = "ROTATED_ON_MONDAY", f"Rotó desde inicio: {prev_last} → {before_l}"
                elif before_l:
                    rot_status, rot_detail = "SAME_SHIFT", f"Mantuvo {before_l} (L el domingo)"
            else:
                first_c = next((c["base"] for c in curr_cells if is_work_shift(c)), None)
                if prev_last and first_c and prev_last != first_c:
                    rot_status, rot_detail = "ROTATED_ON_MONDAY", f"Rotó desde inicio: {prev_last} → {first_c}"
                elif first_c:
                    rot_status, rot_detail = "SAME_SHIFT", f"Mantuvo {first_c}"

        rows.append({
            "emp": emp,
            "prev_cells": prev_cells,
            "curr_cells": curr_cells,
            "prev_off": prev_off,
            "curr_off": curr_off,
            "days_before_off": days_before_off,
            "rest_hours": rest_hours,
            "max_seq": max_seq,
            "off_ok": off_ok,
            "rest_ok": rest_ok,
            "rot_status": rot_status,
            "rot_detail": rot_detail,
        })
    return rows


def get_prev_ref_summary(emp_id, curr_week, prev_week, active_row, rules):
    if not prev_week:
        return ("—", "Primera semana")
    prev_row = prev_week["assignments"].get(emp_id, [parse_raw_cell("") for _ in range(7)])
    prev_sun = prev_row[6] if len(prev_row) == 7 else parse_raw_cell("")
    prev_off = next((i for i in range(6, -1, -1) if prev_row[i].get("base") == "L"), None)
    curr_off = next((i for i, c in enumerate(active_row) if c.get("base") == "L"), None)
    prev_last = next((c["base"] for c in reversed(prev_row) if is_work_shift(c)), None)

    sun_txt = prev_sun.get("raw") or "—"
    prev_off_txt = SHORT_DAYS_ES[prev_off] if prev_off is not None else "—"
    sem_ant = f"Dom: {sun_txt} | L: {prev_off_txt}"

    if prev_off is not None and curr_off is not None:
        days_cnt = (6 - prev_off) + curr_off
        badge = f"{days_cnt}d→L"
    elif curr_off is not None:
        badge = f"L {SHORT_DAYS_ES[curr_off]}"
    else:
        badge = "Sin L"

    rot_txt = "Novedad" if sum(1 for c in active_row if is_absence_state(c)) >= 5 else "Sin cambios"
    if rot_txt != "Novedad":
        if curr_off is not None:
            b_l = next((active_row[i]["base"] for i in range(curr_off - 1, -1, -1) if is_work_shift(active_row[i])), prev_last)
            a_l = next((active_row[i]["base"] for i in range(curr_off + 1, 7) if is_work_shift(active_row[i])), None)
            if b_l and a_l and b_l != a_l:
                rot_txt = f"Rotó tras L ({SHORT_DAYS_ES[curr_off]}): {b_l} → {a_l}"
            elif prev_last and b_l and prev_last != b_l:
                rot_txt = f"Rotó inicio: {prev_last} → {b_l}"
            elif b_l:
                rot_txt = f"Mantuvo {b_l}"
        else:
            first_c = next((c["base"] for c in active_row if is_work_shift(c)), None)
            if prev_last and first_c and prev_last != first_c:
                rot_txt = f"Rotó inicio: {prev_last} → {first_c}"
            elif first_c:
                rot_txt = f"Mantuvo {first_c}"
    return (sem_ant, f"{badge} · {rot_txt}")


def export_week_to_excel_bytes(week, employees):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Horario Semanal"

    start_dt = dt.date.fromisoformat(week["start_date"])
    headers = ["CODIGO", "NOMBRE", "CARGO"] + [
        f"{DAYS_ES[i]} {(start_dt + dt.timedelta(days=i)).day}-{(start_dt + dt.timedelta(days=i)).month:02d}"
        for i in range(7)
    ]
    ws.append(headers)
    thin = Side(border_style="thin", color="475569")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for col_idx in range(1, 11):
        c = ws.cell(row=1, column=col_idx)
        c.font = Font(bold=True, color="0F172A", size=11)
        c.fill = PatternFill("solid", fgColor="E2E8F0")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border
    ws.row_dimensions[1].height = 26

    sorted_emps = sorted(
        [e for e in employees if e["activo"]],
        key=lambda e: (CARGOS_ORDEN.index(e["cargo"]) if e["cargo"] in CARGOS_ORDEN else 99, e["nombre"]),
    )

    for r_idx, emp in enumerate(sorted_emps, start=2):
        row_cells = week["assignments"].get(emp["id"], [parse_raw_cell("") for _ in range(7)])
        vals = [emp["codigo"], emp["nombre"], emp["cargo"]] + [c.get("raw", "") for c in row_cells]
        for c_idx, val in enumerate(vals, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.border = border
            if c_idx <= 3:
                cell.alignment = Alignment(horizontal="center" if c_idx == 1 else "left", vertical="center")
                cell.font = Font(size=10, color="0F172A")
            else:
                sc = row_cells[c_idx - 4]
                bg, fg, _ = get_cell_colors(sc)
                cell.fill = PatternFill("solid", fgColor=bg.lstrip("#"))
                cell.font = Font(size=11, bold=True, color=fg.lstrip("#"))
                cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[r_idx].height = 24

    widths = [10, 24, 26] + [18] * 7
    for idx, w in enumerate(widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(idx)].width = w

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.getvalue()


def parse_dataframe_to_roster(df, start_iso):
    cols = {strip_accents(str(c)).upper().strip(): c for c in df.columns}
    col_cod = next((v for k, v in cols.items() if "COD" in k), None)
    col_nom = next((v for k, v in cols.items() if "NOMB" in k), None)
    col_car = next((v for k, v in cols.items() if "CARGO" in k or "PUESTO" in k), None)
    day_cols = []
    for i, d_es in enumerate(DAYS_ES):
        d_norm = strip_accents(d_es)
        match = next((v for k, v in cols.items() if d_norm in k or d_norm[:3] in k), None)
        if not match and len(df.columns) >= 10:
            match = df.columns[3 + i]
        day_cols.append(match)

    if not col_nom or len([c for c in day_cols if c is not None]) < 7:
        raise ValueError("El archivo o texto debe incluir columnas CODIGO, NOMBRE, CARGO y los 7 días (LUNES a DOMINGO).")

    new_emps = []
    assignments = {}
    for idx, row in df.iterrows():
        nombre = str(row.get(col_nom, "") or "").strip()
        if not nombre or nombre.lower() == "nan":
            continue
        codigo = str(row.get(col_cod, "") or "").strip() if col_cod else str(idx + 1)
        if codigo.endswith(".0"):
            codigo = codigo[:-2]
        cargo_raw = str(row.get(col_car, "") or "Auxiliar de Operaciones").strip() if col_car else "Auxiliar de Operaciones"
        cargo = next((c for c in CARGOS_ORDEN if strip_accents(c).lower() == strip_accents(cargo_raw).lower()), "Auxiliar de Operaciones")
        cells = [parse_raw_cell("" if pd.isna(row.get(dc)) else str(row.get(dc))) for dc in day_cols]
        eid = codigo or f"emp_{idx}"
        new_emps.append({
            "id": eid,
            "codigo": codigo,
            "nombre": nombre,
            "cargo": cargo,
            "puede_at": any(c.get("is_at") for c in cells) or cargo in ("Analista de Operaciones", "Auxiliar de Operaciones"),
            "activo": True,
        })
        assignments[eid] = cells

    if not new_emps:
        raise ValueError("No se encontraron filas válidas de colaboradores.")

    new_week = {
        "id": f"week_{start_iso}",
        "start_date": start_iso,
        "label": week_label_from_monday(start_iso),
        "assignments": assignments,
    }
    return new_emps, new_week


def parse_pasted_text_to_roster(text, start_iso):
    lines = [ln for ln in text.splitlines() if ln.strip()]
    if len(lines) < 2:
        raise ValueError("Pega la fila de encabezados (CODIGO, NOMBRE, CARGO, LUNES...DOMINGO) y al menos un colaborador.")
    sep = "\t" if "\t" in lines[0] else (";" if lines[0].count(";") >= 5 else ",")
    rows = [[cell.strip() for cell in ln.split(sep)] for ln in lines]
    max_len = max(len(r) for r in rows)
    rows = [r + [""] * (max_len - len(r)) for r in rows]
    df = pd.DataFrame(rows[1:], columns=rows[0])
    return parse_dataframe_to_roster(df, start_iso)


def render_novedad_form(key_prefix, employees, default_start_str, allow_rerotate=True, active_week_exists=True):
    emps_sorted = sorted(
        [e for e in employees if e["activo"]],
        key=lambda e: (CARGOS_ORDEN.index(e["cargo"]) if e["cargo"] in CARGOS_ORDEN else 99, e["nombre"]),
    )
    emp_labels = [f"{e['codigo']} — {e['nombre']} ({e['cargo']})" for e in emps_sorted]
    default_monday = dt.date.fromisoformat(default_start_str)

    dt_start_key = f"{key_prefix}_d1"
    dt_end_key = f"{key_prefix}_d2"
    if dt_start_key not in st.session_state:
        st.session_state[dt_start_key] = default_monday
    if dt_end_key not in st.session_state:
        st.session_state[dt_end_key] = default_monday + dt.timedelta(days=4)

    p1, p2, p3, p4 = st.columns(4)
    if p1.button("Solo 1 día", key=f"{key_prefix}_p1", use_container_width=True):
        st.session_state[dt_start_key] = default_monday
        st.session_state[dt_end_key] = default_monday
        st.rerun()
    if p2.button("Lun–Vie (5 días)", key=f"{key_prefix}_p2", use_container_width=True):
        st.session_state[dt_start_key] = default_monday
        st.session_state[dt_end_key] = default_monday + dt.timedelta(days=4)
        st.rerun()
    if p3.button("Sáb–Dom", key=f"{key_prefix}_p3", use_container_width=True):
        st.session_state[dt_start_key] = default_monday + dt.timedelta(days=5)
        st.session_state[dt_end_key] = default_monday + dt.timedelta(days=6)
        st.rerun()
    if p4.button("Toda la semana", key=f"{key_prefix}_p4", use_container_width=True):
        st.session_state[dt_start_key] = default_monday
        st.session_state[dt_end_key] = default_monday + dt.timedelta(days=6)
        st.rerun()

    c1, c2, c3, c4 = st.columns([2.2, 1.2, 1.1, 1.1])
    sel_emp_lbl = c1.selectbox("Colaborador", emp_labels, key=f"{key_prefix}_emp")
    sel_tipo = c2.selectbox("Tipo de novedad", NOVEDAD_OPTIONS, key=f"{key_prefix}_tipo")
    d_start = c3.date_input("Desde", key=dt_start_key)
    d_end = c4.date_input("Hasta", key=dt_end_key)

    o1, o2, o3 = st.columns([1.4, 1.4, 1.4])
    save_future = o1.checkbox("Guardar para futuras semanas", value=True, key=f"{key_prefix}_save")
    apply_now = o2.checkbox(
        "Aplicar en semanas ya creadas",
        value=active_week_exists,
        disabled=not active_week_exists,
        key=f"{key_prefix}_apply",
    )
    rerotate = (
        o3.checkbox("Re-rotar semana actual", value=False, disabled=not active_week_exists, key=f"{key_prefix}_rerot")
        if allow_rerotate
        else False
    )

    if st.button("Guardar Novedad", type="primary", key=f"{key_prefix}_submit"):
        emp_idx = emp_labels.index(sel_emp_lbl)
        emp = emps_sorted[emp_idx]
        s_iso = min(d_start, d_end).strftime("%Y-%m-%d")
        e_iso = max(d_start, d_end).strftime("%Y-%m-%d")
        nov = {
            "id": f"nov_{emp['id']}_{dt.datetime.now().timestamp()}",
            "emp_id": emp["id"],
            "tipo": sel_tipo,
            "start_date": s_iso,
            "end_date": e_iso,
        }
        if save_future:
            st.session_state["scheduled_novedades"].append(nov)

        applied_total = 0
        if apply_now and st.session_state["weeks"]:
            new_weeks = []
            for w in st.session_state["weeks"]:
                nw, cnt = apply_novedad_to_week(w, nov)
                applied_total += cnt
                new_weeks.append(nw)
            st.session_state["weeks"] = new_weeks

        if rerotate and st.session_state["weeks"]:
            trigger_rerotate_active_week()

        st.success(f"Novedad {sel_tipo} registrada para {emp['nombre']} ({s_iso} al {e_iso}). Días aplicados: {applied_total}.")
        st.rerun()


def render_no_active_week_reminder(section_title):
    card_html = (
        "<div style='background:#FFFFFF;border:1px solid #CBD5E1;border-left:5px solid #0369A1;"
        "border-radius:10px;padding:24px 26px;margin:12px 0;max-width:860px;'>"
        "<div style='display:inline-block;background:#E0F2FE;color:#075985;border:1px solid #BAE6FD;"
        "border-radius:6px;padding:3px 9px;font-size:11px;font-weight:700;text-transform:uppercase;"
        "letter-spacing:0.04em;margin-bottom:10px;'>Paso inicial requerido</div>"
        "<h3 style='margin:0 0 8px 0;color:#0F172A;font-size:20px;'>No hay ninguna semana activa cargada todavía</h3>"
        f"<p style='margin:0 0 14px 0;color:#475569;font-size:14px;line-height:1.5;'>Para usar <strong>{section_title}</strong>, primero debes cargar tu programación base y establecerla como <strong>Semana Activa</strong>.</p>"
        "<div style='background:#F8FAFC;border:1px solid #E2E8F0;border-radius:8px;padding:14px 18px;color:#334155;font-size:13px;line-height:1.65;'>"
        "<div style='font-weight:700;color:#0F172A;margin-bottom:6px;'>¿Dónde y cómo cargar tu programación?</div>"
        "<ol style='margin:0;padding-left:18px;'>"
        "<li>Ve a la pestaña superior <strong>Importar / Exportar Excel</strong>.</li>"
        "<li>Selecciona el <strong>Lunes de inicio</strong> correspondiente a tu semana.</li>"
        "<li>Elige <strong>Subir archivo Excel (.xlsx)</strong> o <strong>Pegar tabla desde Excel (Ctrl + C / Ctrl + V)</strong> con las columnas <code>CODIGO, NOMBRE, CARGO, LUNES ... DOMINGO</code>.</li>"
        "<li>Al confirmar la carga, esa semana quedará colocada automáticamente como tu <strong>Semana Activa</strong>.</li>"
        "</ol>"
        "</div>"
        "</div>"
    )
    st.markdown(card_html, unsafe_allow_html=True)


# ----------------- INICIO DE APP -----------------
init_state()
sync_widget_state_to_configs()

st.title("Planificador de Turnos Operativos")
st.caption("Rotación semanal tras descansos · Descansos 12h / 24h · Meta diaria (Lun–Vie, Sáb, Dom) · Códigos de cobertura Técnico AT · Gestión de turnos y personal")

weeks = st.session_state["weeks"]
employees = st.session_state["employees"]
role_configs = st.session_state["role_configs"]
rules = st.session_state["rules"]
has_weeks = len(weeks) > 0

if has_weeks:
    if st.session_state["active_week_idx"] >= len(weeks):
        st.session_state["active_week_idx"] = len(weeks) - 1
    active_idx = st.session_state["active_week_idx"]
    active_week = weeks[active_idx]
    prev_week = weeks[active_idx - 1] if active_idx > 0 else None
else:
    active_idx = 0
    active_week = None
    prev_week = None

top1, top2, top3, top4, top5 = st.columns([2.8, 1.0, 1.5, 1.8, 1.6])
with top1:
    if has_weeks:
        week_labels = [f"{i + 1}. {w['label']}" for i, w in enumerate(weeks)]
        chosen_lbl = st.selectbox("Semana activa", week_labels, index=active_idx)
        chosen_idx = week_labels.index(chosen_lbl)
        if chosen_idx != active_idx:
            st.session_state["active_week_idx"] = chosen_idx
            st.rerun()
    else:
        st.selectbox("Semana activa", ["— Sin semana activa cargada —"], index=0, disabled=True)

with top2:
    st.write("")
    st.write("")
    if st.button("Eliminar", disabled=not has_weeks, use_container_width=True):
        st.session_state["weeks"].pop(active_idx)
        st.session_state["active_week_idx"] = max(0, active_idx - 1)
        st.rerun()

with top3:
    st.write("")
    st.write("")
    if st.button("Re-rotar semana actual", key="top_rerotate_btn", disabled=not has_weeks, use_container_width=True):
        lbl_r, at_r = trigger_rerotate_active_week()
        st.success(f"Semana '{lbl_r}' re-rotada correctamente ({at_r} cobertura(s) AT).")
        st.rerun()

with top4:
    st.write("")
    st.write("")
    if st.button("Generar Siguiente Semana", type="primary", disabled=not has_weeks, use_container_width=True):
        nw, at_c = generate_next_week(
            weeks[-1],
            employees,
            role_configs,
            rules,
            scheduled_novedades=st.session_state["scheduled_novedades"],
        )
        st.session_state["weeks"].append(nw)
        st.session_state["active_week_idx"] = len(st.session_state["weeks"]) - 1
        st.success(f"Semana '{nw['label']}' generada ({at_c} turno(s) AT asignados).")
        st.rerun()

with top5:
    st.write("")
    st.write("")
    if has_weeks:
        excel_bytes = export_week_to_excel_bytes(active_week, employees)
        st.download_button(
            "Descargar Excel (.xlsx)",
            data=excel_bytes,
            file_name=f"Horario_{active_week['start_date']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    else:
        st.button("Descargar Excel (.xlsx)", disabled=True, use_container_width=True)

if has_weeks:
    issues = validate_schedule(active_week, employees, role_configs, rules, prev_week)
    uncovered = compute_uncovered_tasks(active_week, employees, role_configs)
    active_emps_cnt = sum(1 for e in employees if e["activo"])
    novedades_cnt = sum(
        1 for e in employees if e["activo"]
        and any(c.get("base") in ("VACACIONES", "SUSPEND", "INCAPACIDAD", "PERMISO") for c in active_week["assignments"].get(e["id"], []))
    )
    at_cnt = sum(sum(1 for c in active_week["assignments"].get(e["id"], []) if c.get("is_at")) for e in employees if e["activo"])
    rest_alerts_cnt = sum(1 for i in issues if i["type"] != "UNCOVERED_TECH")
    tech_gaps_cnt = sum(1 for i in issues if i["type"] == "UNCOVERED_TECH")
    uncovered_total = sum(u["missing"] for u in uncovered)

    m1, m2, m3, m4, m5, m6 = st.columns(6)
    m1.metric("Activos", f"{active_emps_cnt} / {len(employees)}")
    m2.metric("Novedades", novedades_cnt)
    m3.metric("Turnos AT", at_cnt)
    m4.metric("Descansos 12h/24h", "OK" if rest_alerts_cnt == 0 else f"{rest_alerts_cnt} alertas")
    m5.metric("Técnicos", "100% cubierto" if tech_gaps_cnt == 0 else f"{tech_gaps_cnt} faltante(s)")
    m6.metric("Tareas sin cubrir", uncovered_total)
else:
    issues = []
    uncovered = []
    st.info("Sin semana activa: abre la pestaña **Importar / Exportar Excel** para subir tu programación o pegarla desde Excel.")

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "Horario Semanal",
    "Rotación y Descansos",
    "Cobertura y Apoyo AT",
    "Turnos, Reglas y Personal",
    "Importar / Exportar Excel",
])

with tab1:
    if not has_weeks:
        render_no_active_week_reminder("Horario Semanal")
    else:
        f1, f2, f3, f4, f5, f6 = st.columns([1.9, 1.4, 1.1, 1.2, 1.3, 1.3])
        search_q = f1.text_input("Buscar código o nombre", placeholder="Ej. 671 o Guardao")
        cargo_sel = f2.selectbox("Filtrar por cargo", ["TODOS"] + CARGOS_ORDEN, key="t1_cargo")
        only_alerts = f3.checkbox(f"Solo alertas ({len(issues)})", key="t1_alerts")
        show_prev_ref = f4.checkbox("Ver ref. semana anterior", value=True, key="t1_prev")
        if f5.button("Auto-cubrir AT", use_container_width=True, key="t1_auto_at_top"):
            nw, cnt = auto_assign_at(active_week, employees, role_configs, rules, prev_week)
            st.session_state["weeks"][active_idx] = nw
            st.success(f"Se asignaron {cnt} turno(s) AT automáticamente.")
            st.rerun()
        if f6.button("Re-rotar semana actual", type="primary", use_container_width=True, key="t1_rerotate_top"):
            lbl_r, at_r = trigger_rerotate_active_week()
            st.success(f"Semana '{lbl_r}' re-rotada ({at_r} cobertura(s) AT).")
            st.rerun()

        legend_items = [
            ("#70AD47", "#558733", "03:00"),
            ("#B6C3B4", "#92A390", "06:00 / 07:00"),
            ("#D9D9D9", "#AEAEAE", "08:00"),
            ("#CC9900", "#9E7600", "11:00"),
            ("#C850B0", "#9E3889", "13:00 / Apoyo AT"),
            ("#8E9FB4", "#6C7F96", "19:00"),
            ("#25A9D8", "#1885AD", "20:00"),
            ("#949494", "#6E6E6E", "22:00"),
            ("#B6BE5F", "#8E9640", "23:00"),
            ("#C6B39A", "#A89378", "L (Libre)"),
        ]
        legend_html = "<div style='display:flex;flex-wrap:wrap;gap:14px;align-items:center;margin:4px 0 10px 0;font-size:12px;color:#334155;'>"
        for l_bg, l_brd, l_txt in legend_items:
            legend_html += (
                f"<span style='display:inline-flex;align-items:center;gap:6px;'>"
                f"<span style='display:inline-block;width:14px;height:14px;border-radius:3px;background:{l_bg};border:1px solid {l_brd};'></span>"
                f"<span>{l_txt}</span></span>"
            )
        legend_html += "</div>"
        st.markdown(legend_html, unsafe_allow_html=True)

        start_dt = dt.date.fromisoformat(active_week["start_date"])
        day_headers = [f"{SHORT_DAYS_ES[i]} {(start_dt + dt.timedelta(days=i)).day}-{MONTHS_ES[(start_dt + dt.timedelta(days=i)).month - 1]}" for i in range(7)]

        issue_emp_ids = {i["emp_id"] for i in issues if i["emp_id"]}
        filtered_emps = sorted(
            [e for e in employees if e["activo"]],
            key=lambda e: (CARGOS_ORDEN.index(e["cargo"]) if e["cargo"] in CARGOS_ORDEN else 99, e["nombre"]),
        )
        if cargo_sel != "TODOS":
            filtered_emps = [e for e in filtered_emps if e["cargo"] == cargo_sel]
        if search_q.strip():
            q = search_q.strip().lower()
            filtered_emps = [e for e in filtered_emps if q in e["nombre"].lower() or q in e["codigo"].lower()]
        if only_alerts:
            filtered_emps = [e for e in filtered_emps if e["id"] in issue_emp_ids]

        active_emps_all = [e for e in employees if e["activo"]]
        th_style = "padding:10px 11px;border:1px solid #94A3B8;background:#F1F5F9;color:#0F172A;font-size:12px;font-weight:700;text-align:left;white-space:nowrap;"
        th_cells = f"<th style='{th_style}'>CODIGO</th><th style='{th_style}'>NOMBRE</th><th style='{th_style}'>CARGO</th>"
        if show_prev_ref:
            th_cells += f"<th style='{th_style}'>SEM ANT (Dom / L)</th><th style='{th_style}'>DÍAS P/ L Y ROT.</th>"
        for d_idx, dh in enumerate(day_headers):
            at_day_cnt = sum(1 for e in active_emps_all if (active_week["assignments"].get(e["id"], [{}] * 7))[d_idx].get("is_at"))
            th_cells += (
                f"<th style='{th_style}text-align:center;min-width:146px;'>"
                f"<div>{dh}</div>"
                f"<div style='font-size:11px;color:#86198F;font-weight:700;'>{at_day_cnt} AT</div>"
                f"</th>"
            )

        rows_html_list = []
        for e in filtered_emps:
            r = active_week["assignments"].get(e["id"], [parse_raw_cell("") for _ in range(7)])
            at_badge = " <span style='color:#9D174D;font-size:10px;font-weight:700;'>AT</span>" if e["puede_at"] else ""
            tr_cells = (
                f"<td style='padding:8px 10px;border:1px solid #CBD5E1;font-family:monospace;font-size:13px;font-weight:700;white-space:nowrap;'>{e['codigo']}</td>"
                f"<td style='padding:8px 10px;border:1px solid #CBD5E1;font-size:13px;font-weight:600;white-space:nowrap;'>{e['nombre']}{at_badge}</td>"
                f"<td style='padding:8px 10px;border:1px solid #CBD5E1;font-size:12px;white-space:nowrap;color:#334155;'>{e['cargo']}</td>"
            )
            if show_prev_ref:
                sem_ant, rot_txt = get_prev_ref_summary(e["id"], active_week, prev_week, r, rules)
                tr_cells += (
                    f"<td style='padding:8px 10px;border:1px solid #CBD5E1;font-size:11px;color:#475569;white-space:nowrap;'>{sem_ant}</td>"
                    f"<td style='padding:8px 10px;border:1px solid #CBD5E1;font-size:11px;color:#0F172A;white-space:nowrap;'>{rot_txt}</td>"
                )
            for d_idx in range(7):
                c = r[d_idx]
                bg, fg, _ = get_cell_colors(c)
                lock_icon = " 🔒" if c.get("locked") else ""
                has_err = any(i["emp_id"] == e["id"] and i["day_idx"] == d_idx for i in issues)
                cell_border = "2px solid #DC2626" if has_err else "1px solid #64748B"
                txt = c.get("raw") or "—"
                tr_cells += (
                    f"<td style='min-width:146px;height:42px;padding:9px 12px;border:{cell_border};"
                    f"background:{bg};color:{fg};font-family:ui-monospace,SFMono-Regular,Menlo,monospace;"
                    f"font-weight:800;font-size:14px;text-align:center;letter-spacing:-0.02em;white-space:nowrap;'>"
                    f"{txt}{lock_icon}</td>"
                )
            rows_html_list.append(f"<tr>{tr_cells}</tr>")

        table_html = (
            "<div style='overflow-x:auto;border:1px solid #94A3B8;border-radius:8px;background:#FFFFFF;'>"
            "<table style='border-collapse:collapse;width:100%;font-family:system-ui,-apple-system,sans-serif;font-size:13px;'>"
            f"<thead><tr>{th_cells}</tr></thead>"
            f"<tbody>{''.join(rows_html_list)}</tbody>"
            "</table></div>"
        )
        st.markdown(table_html, unsafe_allow_html=True)

        b1, b2 = st.columns([3, 1.4])
        b1.caption(f"Mostrando {len(filtered_emps)} de {active_emps_cnt} colaboradores activos ({len(employees) - active_emps_cnt} inactivos en directorio).")
        if b2.button("Re-rotar semana actual", key="t1_rerotate_bottom", use_container_width=True):
            lbl_r, at_r = trigger_rerotate_active_week()
            st.success(f"Semana '{lbl_r}' re-rotada aplicando tus reglas y Meta diaria ({at_r} turno(s) AT asignados).")
            st.rerun()

        with st.expander("➕ Agregar novedad por rango (Vacaciones, Suspensión, Libre, Incapacidad, Permiso)"):
            render_novedad_form("t1_nov", employees, active_week["start_date"], allow_rerotate=True, active_week_exists=True)

        with st.expander("✏️ Editar turno o novedad puntual de un colaborador (uno o varios días)"):
            emps_sorted = sorted(
                [e for e in employees if e["activo"]],
                key=lambda e: (CARGOS_ORDEN.index(e["cargo"]) if e["cargo"] in CARGOS_ORDEN else 99, e["nombre"]),
            )
            e_labels = [f"{e['codigo']} — {e['nombre']} ({e['cargo']})" for e in emps_sorted]
            ed1, ed2, ed3 = st.columns([2.2, 1.1, 1.1])
            sel_e = ed1.selectbox("Colaborador", e_labels, key="ed_emp")
            d_from = ed2.selectbox("Desde día", DAYS_ES, index=0, key="ed_d1")
            d_to = ed3.selectbox("Hasta día", DAYS_ES, index=0, key="ed_d2")

            emp_obj = emps_sorted[e_labels.index(sel_e)]
            d_idx_1 = DAYS_ES.index(d_from)
            d_idx_2 = max(d_idx_1, DAYS_ES.index(d_to))
            cur_cell = active_week["assignments"].get(emp_obj["id"], [parse_raw_cell("") for _ in range(7)])[d_idx_1]

            role_pool = get_all_cargo_shifts(role_configs.get(emp_obj["cargo"], {}))
            preset_opts = ["(Escribir manual)"] + role_pool + ["L", "VACACIONES", "SUSPEND", "INCAPACIDAD", "PERMISO"]
            ed4, ed5, ed6, ed7 = st.columns([1.8, 1.8, 1.0, 1.0])
            sel_preset = ed4.selectbox("Turno o novedad rápida", preset_opts, key="ed_preset")
            manual_val = ed5.text_input("Valor de celda", value=cur_cell.get("raw", "") if sel_preset == "(Escribir manual)" else sel_preset, key="ed_val")
            chk_at = ed6.checkbox("AT", value=cur_cell.get("is_at", False), key="ed_at")
            chk_lock = ed7.checkbox("Fijar al rotar", value=cur_cell.get("locked", False), key="ed_lock")
            save_future_ed = st.checkbox("Guardar también en Novedades programadas si es novedad", value=True, key="ed_save_nov")

            if st.button("Guardar cambio en rango seleccionado", type="primary", key="ed_save_btn"):
                base_clean = manual_val.strip()
                if base_clean.upper() in SPECIAL_STATES:
                    final_raw = base_clean.upper()
                    final_lock = True
                else:
                    parsed_m = parse_raw_cell(base_clean)
                    final_raw = format_cell_raw(parsed_m["base"], is_at=chk_at, is_cap=parsed_m["is_cap"])
                    final_lock = chk_lock

                row = active_week["assignments"].setdefault(emp_obj["id"], [parse_raw_cell("") for _ in range(7)])
                for dd in range(d_idx_1, d_idx_2 + 1):
                    row[dd] = parse_raw_cell(final_raw, locked=final_lock)

                if save_future_ed and base_clean.upper() in SPECIAL_STATES:
                    st.session_state["scheduled_novedades"].append({
                        "id": f"nov_{emp_obj['id']}_{dt.datetime.now().timestamp()}",
                        "emp_id": emp_obj["id"],
                        "tipo": base_clean.upper(),
                        "start_date": add_days_iso(active_week["start_date"], d_idx_1),
                        "end_date": add_days_iso(active_week["start_date"], d_idx_2),
                    })
                st.success(f"Actualizado {emp_obj['nombre']} de {DAYS_ES[d_idx_1]} a {DAYS_ES[d_idx_2]}.")
                st.rerun()

        if issues:
            st.subheader("Alertas activas de esta semana")
            for iss in issues:
                st.warning(iss["msg"])

with tab2:
    if not has_weeks:
        render_no_active_week_reminder("Rotación y Descansos")
    else:
        st.subheader("Análisis de Días para Descanso (L) y Rotación Tras Días Libres")
        st.caption("Compara la semana anterior contra la semana nueva, registra novedades antes de generarla y verifica días continuos antes de L.")

        default_mode = 0 if prev_week else 1
        mode_choice = st.radio(
            "Modo de visualización",
            [f"Semana Activa ({active_week['label']})", "Previsualizar Siguiente Semana"],
            index=default_mode,
            horizontal=True,
        )
        is_preview = mode_choice == "Previsualizar Siguiente Semana"

        p1, p2, p3, p4 = st.columns([1.2, 1.2, 1.4, 1.6])
        rules["min_work_between_off"] = int(p1.number_input("Mín. días trabajados entre L", 3, 6, value=int(rules["min_work_between_off"]), key="t2_min_work"))
        rules["max_consecutive"] = int(p2.number_input("Máx. días seguidos sin descanso", 5, 10, value=int(rules["max_consecutive"]), key="t2_max_cons"))
        with p3:
            st.write("")
            st.write("")
            rules["rotate_after_off"] = st.checkbox("Rotar turno después del Día Libre (L)", value=bool(rules["rotate_after_off"]), key="t2_rot")

        preview_week, preview_at = generate_next_week(
            weeks[-1],
            employees,
            role_configs,
            rules,
            scheduled_novedades=st.session_state["scheduled_novedades"],
        )
        with p4:
            st.write("")
            st.write("")
            if is_preview:
                if st.button(f"Aplicar y Crear {preview_week['label']}", type="primary", use_container_width=True):
                    st.session_state["weeks"].append(preview_week)
                    st.session_state["active_week_idx"] = len(st.session_state["weeks"]) - 1
                    st.success(f"Semana '{preview_week['label']}' creada ({preview_at} cobertura(s) AT).")
                    st.rerun()

        sched_novs = st.session_state["scheduled_novedades"]
        with st.expander(f"📅 Novedades programadas ({len(sched_novs)}) — Vacaciones, Suspensión, Libre antes de generar"):
            if sched_novs:
                for idx_n, nov in enumerate(sched_novs):
                    emp_n = next((e for e in employees if e["id"] == nov["emp_id"]), None)
                    emp_str = f"{emp_n['codigo']} {emp_n['nombre']}" if emp_n else nov["emp_id"]
                    nc1, nc2 = st.columns([5, 1])
                    bg_n, fg_n, bd_n = get_novedad_badge_colors(nov["tipo"])
                    nc1.markdown(
                        f"<div style='padding:6px 10px;border-radius:6px;background:{bg_n};color:{fg_n};border:1px solid {bd_n};font-size:12px;'>"
                        f"<strong>{emp_str}</strong> · <strong>{nov['tipo']}</strong> · {nov['start_date']} → {nov['end_date']}"
                        "</div>",
                        unsafe_allow_html=True,
                    )
                    if nc2.button("Quitar", key=f"t2_del_nov_{idx_n}"):
                        st.session_state["scheduled_novedades"].pop(idx_n)
                        st.rerun()
            else:
                st.info("No hay novedades programadas. Regístralas abajo para que queden bloqueadas en la siguiente semana.")
            render_novedad_form("t2_nov", employees, preview_week["start_date"], allow_rerotate=False, active_week_exists=True)

        eval_curr = preview_week if is_preview else active_week
        eval_prev = weeks[-1] if is_preview else prev_week
        analyses = analyze_rotation(eval_curr, eval_prev, employees, role_configs, rules)

        active_an = [a for a in analyses if a["rot_status"] != "ABSENCE"]
        valid_days = [a["days_before_off"] for a in active_an if a["days_before_off"] is not None]
        avg_days = round(sum(valid_days) / len(valid_days), 1) if valid_days else 0
        rotated_cnt = sum(1 for a in active_an if a["rot_status"] in ("ROTATED_AFTER_L", "ROTATED_ON_MONDAY"))
        interval_bad = sum(1 for a in active_an if not a["off_ok"])

        sm1, sm2, sm3, sm4 = st.columns(4)
        sm1.metric("Ventana comparada", eval_curr["label"])
        sm2.metric("Promedio días antes de L", f"{avg_days} días")
        sm3.metric("Rotaciones aplicadas", f"{rotated_cnt} / {len(active_an)}")
        sm4.metric("Intervalo fuera de rango", interval_bad)

        fc1, fc2 = st.columns([2, 2])
        cargo_an = fc1.selectbox("Filtrar cargo", ["TODOS"] + CARGOS_ORDEN, key="t2_cargo")
        only_bad = fc2.checkbox("Mostrar solo sin rotación o alertas de descanso", key="t2_only_bad")

        filtered_an = sorted(
            analyses,
            key=lambda a: (CARGOS_ORDEN.index(a["emp"]["cargo"]) if a["emp"]["cargo"] in CARGOS_ORDEN else 99, a["emp"]["nombre"]),
        )
        if cargo_an != "TODOS":
            filtered_an = [a for a in filtered_an if a["emp"]["cargo"] == cargo_an]
        if only_bad:
            filtered_an = [a for a in filtered_an if not a["off_ok"] or not a["rest_ok"] or a["rot_status"] == "SAME_SHIFT"]

        def strip_html(cells):
            if not cells:
                return "<span style='color:#94A3B8;'>Sin semana previa</span>"
            boxes = []
            for d, c in enumerate(cells):
                bg, fg, brd = get_cell_colors(c)
                b = c.get("base", "")
                if b == "L":
                    line1, line2 = "LIBRE", ""
                elif b == "VACACIONES":
                    line1, line2 = "VAC", ""
                elif b == "SUSPEND":
                    line1, line2 = "SUSP", ""
                elif b == "INCAPACIDAD":
                    line1, line2 = "INCAP", ""
                elif b == "PERMISO":
                    line1, line2 = "PERM", ""
                elif "-" in b:
                    p1_h, p2_h = b.split("-", 1)
                    line1 = p1_h
                    line2 = f"-{p2_h}" + (" AT" if c.get("is_at") else "")
                else:
                    line1 = b[:6] if b else "—"
                    line2 = "AT" if c.get("is_at") else ""
                sub_html = f"<div style='font-size:11px;font-weight:700;margin-top:1px;'>{line2}</div>" if line2 else ""
                boxes.append(
                    f"<div title='{DAYS_ES[d]}: {c.get('raw', '')}' style='min-width:68px;padding:6px 5px;border-radius:5px;"
                    f"background:{bg};color:{fg};border:1px solid {brd};text-align:center;font-family:monospace;line-height:1.18;'>"
                    f"<div style='font-size:10px;opacity:0.8;font-weight:700;margin-bottom:2px;'>{SHORT_DAYS_ES[d]}</div>"
                    f"<div style='font-size:12px;font-weight:800;'>{line1}</div>"
                    f"{sub_html}"
                    "</div>"
                )
            return "<div style='display:flex;gap:6px;flex-wrap:nowrap;'>" + "".join(boxes) + "</div>"

        th2 = "padding:8px 10px;border:1px solid #CBD5E1;background:#F1F5F9;color:#0F172A;font-size:12px;text-align:left;white-space:nowrap;"
        comp_rows = []
        for a in filtered_an:
            e = a["emp"]
            days_txt = "—" if a["rot_status"] == "ABSENCE" else (f"{a['days_before_off']} días seg." if a["days_before_off"] is not None else f"Sin L ({a['max_seq']}d)")
            days_color = "#047857" if a["off_ok"] else "#B45309"
            hrs_txt = f"{a['rest_hours']}h" if a["rest_hours"] is not None else "—"
            hrs_color = "#047857" if a["rest_ok"] else "#DC2626"
            rot_color = "#B45309" if a["rot_status"] == "SAME_SHIFT" else "#0F172A"
            comp_rows.append(
                "<tr>"
                f"<td style='padding:8px 10px;border:1px solid #E2E8F0;min-width:190px;'><div style='font-weight:600;'>{e['codigo']} {e['nombre']}</div><div style='font-size:11px;color:#64748B;'>{e['cargo']}</div></td>"
                f"<td style='padding:8px 10px;border:1px solid #E2E8F0;min-width:520px;'>{strip_html(a['prev_cells'])}</td>"
                f"<td style='padding:8px 10px;border:1px solid #E2E8F0;min-width:520px;'>{strip_html(a['curr_cells'])}</td>"
                f"<td style='padding:8px 10px;border:1px solid #E2E8F0;text-align:center;font-weight:700;color:{days_color};white-space:nowrap;'>{days_txt}</td>"
                f"<td style='padding:8px 10px;border:1px solid #E2E8F0;text-align:center;font-weight:700;color:{hrs_color};white-space:nowrap;'>{hrs_txt}</td>"
                f"<td style='padding:8px 10px;border:1px solid #E2E8F0;color:{rot_color};min-width:220px;'>{a['rot_detail']}</td>"
                "</tr>"
            )

        comp_table_html = (
            "<div style='overflow-x:auto;border:1px solid #CBD5E1;border-radius:8px;background:#FFFFFF;'>"
            "<table style='border-collapse:collapse;width:100%;font-family:system-ui,-apple-system,sans-serif;font-size:12px;'>"
            f"<thead><tr><th style='{th2}'>Colaborador y Cargo</th><th style='{th2}'>Semana Anterior (Lun → Dom)</th>"
            f"<th style='{th2}'>{'Nueva Semana Propuesta' if is_preview else 'Semana Actual'} (Lun → Dom)</th>"
            f"<th style='{th2}'>Días p/ Descanso</th><th style='{th2}'>Horas en L</th><th style='{th2}'>Rotación tras Día Libre</th></tr></thead>"
            f"<tbody>{''.join(comp_rows)}</tbody>"
            "</table></div>"
        )
        st.markdown(comp_table_html, unsafe_allow_html=True)

with tab3:
    if not has_weeks:
        render_no_active_week_reminder("Cobertura y Apoyo AT")
    else:
        st.subheader("Cobertura de Técnico de Operaciones y Códigos AT por Día")
        st.caption("Visualiza los códigos de quién cubre directamente como Técnico y quién cubre las tareas como Apoyo Técnico (AT) cada día. El turno 19:00-27:00 requiere Técnico directo.")

        daily_cov = get_daily_tech_coverage(active_week, employees, role_configs)
        tech_cfg = role_configs.get("Técnico de Operaciones", {})
        tech_shifts = get_all_cargo_shifts(tech_cfg)

        c3_m1, c3_m2, c3_m3 = st.columns(3)
        c3_m1.metric("Total coberturas AT en la semana", sum(len(x["at_members"]) for x in daily_cov))
        gaps_t = sum(x["missing"] for x in daily_cov)
        c3_m2.metric("Estado cobertura Técnicos", "100% cubierto" if gaps_t == 0 else f"{gaps_t} faltante(s)")
        c3_m3.metric("Personal habilitado AT activo", sum(1 for e in employees if e["activo"] and e["puede_at"]))

        th3 = "padding:8px 10px;border:1px solid #CBD5E1;background:#F1F5F9;color:#0F172A;font-size:12px;text-align:left;white-space:nowrap;"
        start_dt3 = dt.date.fromisoformat(active_week["start_date"])
        d_hdrs3 = [f"{SHORT_DAYS_ES[i]} {(start_dt3 + dt.timedelta(days=i)).day}-{MONTHS_ES[(start_dt3 + dt.timedelta(days=i)).month - 1]}" for i in range(7)]
        hdr3_html = f"<th style='{th3}'>Turno Técnico</th><th style='{th3}'>Meta L-V/Sáb/Dom</th>" + "".join(
            f"<th style='{th3}min-width:155px;'>{dh}</th>" for dh in d_hdrs3
        )

        body3_rows = []
        for s in tech_shifts:
            no_at = is_tech_no_at(s)
            m_wk = get_shift_target_for_day(tech_cfg, s, 0)
            m_sat = get_shift_target_for_day(tech_cfg, s, 5)
            m_sun = get_shift_target_for_day(tech_cfg, s, 6)
            tds = [
                f"<td style='padding:8px 10px;border:1px solid #E2E8F0;font-weight:700;'>{s}<br><span style='font-size:10px;font-weight:500;color:#64748B;'>{'Solo Técnico directo' if no_at else 'Acepta AT'}</span></td>",
                f"<td style='padding:8px 10px;border:1px solid #E2E8F0;text-align:center;font-family:monospace;font-weight:700;'>{m_wk}/{m_sat}/{m_sun}</td>",
            ]
            for d_idx in range(7):
                item = next((x for x in daily_cov if x["day_idx"] == d_idx and x["shift_code"] == s), None)
                if not item:
                    tds.append("<td style='padding:8px;border:1px solid #E2E8F0;'>—</td>")
                    continue
                dir_codes = ", ".join(f"{m['codigo']} ({m['nombre'].split()[0]})" for m in item["direct_members"])
                at_codes = ", ".join(f"{m['codigo']} ({m['nombre'].split()[0]}) [AT]" for m in item["at_members"])
                dir_html = f"<div style='color:#1E3A8A;font-size:11px;font-weight:700;'>Téc: {dir_codes}</div>" if dir_codes else ""
                at_html = f"<div style='color:#9D174D;font-size:11px;font-weight:700;'>{at_codes}</div>" if at_codes else ""
                empty_html = "<div style='color:#94A3B8;font-size:11px;'>Sin asignación</div>" if (not dir_codes and not at_codes) else ""
                bg_c = "#FFFBEB" if (item["missing"] > 0 and item["target"] > 0) else "#FFFFFF"
                badge_color = "#047857" if item["missing"] == 0 else "#B45309"
                tds.append(
                    f"<td style='padding:7px 9px;border:1px solid #E2E8F0;background:{bg_c};vertical-align:top;'>"
                    f"<div style='font-family:monospace;font-weight:800;font-size:12px;color:{badge_color};margin-bottom:4px;'>{item['total_assigned']}/{item['target']}</div>"
                    f"{dir_html}{at_html}{empty_html}</td>"
                )
            body3_rows.append(f"<tr>{''.join(tds)}</tr>")

        st.markdown(
            "<div style='overflow-x:auto;border:1px solid #CBD5E1;border-radius:8px;background:#FFFFFF;'>"
            "<table style='border-collapse:collapse;width:100%;font-family:system-ui,-apple-system,sans-serif;font-size:12px;'>"
            f"<thead><tr>{hdr3_html}</tr></thead>"
            f"<tbody>{''.join(body3_rows)}</tbody>"
            "</table></div>",
            unsafe_allow_html=True,
        )

        sel_day_det = st.selectbox("Inspeccionar detalle diario exacto", DAYS_ES, key="t3_day_det")
        sel_day_idx = DAYS_ES.index(sel_day_det)
        det_items = [x for x in daily_cov if x["day_idx"] == sel_day_idx]
        d_cols = st.columns(max(1, len(det_items)))
        for idx_c, item in enumerate(det_items):
            with d_cols[idx_c]:
                st.markdown(f"**{item['shift_code']}** ({item['total_assigned']}/{item['target']})")
                for m in item["direct_members"]:
                    st.caption(f"Téc · **{m['codigo']}** — {m['nombre']}")
                for m in item["at_members"]:
                    st.caption(f"AT · **{m['codigo']}** — {m['nombre']} ({m['cargo']})")
                if item["missing"] > 0:
                    st.warning(f"Faltan {item['missing']}")

        st.divider()
        st.subheader("Tareas Sin Cubrir según Meta Diaria (Lun–Vie / Sáb / Dom)")
        if not uncovered:
            st.success("¡100% cubierto! Todas las tareas configuradas en la Meta diaria tienen su personal asignado en los 7 días.")
        else:
            st.warning(f"Hay {sum(u['missing'] for u in uncovered)} vacante(s) por cubrir en {len(uncovered)} turno(s).")
            u_cargo_f = st.selectbox("Filtrar tareas faltantes por cargo", ["TODOS"] + CARGOS_ORDEN, key="t3_unc_cargo")
            unc_rows = [u for u in uncovered if u_cargo_f == "TODOS" or u["cargo"] == u_cargo_f]
            st.dataframe(
                pd.DataFrame([
                    {
                        "Día": u["day_name"],
                        "Cargo": u["cargo"],
                        "Turno": u["shift_code"] + (" (Solo Técnico directo)" if u["cargo"] == "Técnico de Operaciones" and is_tech_no_at(u["shift_code"]) else ""),
                        "Meta": u["target"],
                        "Asignados": u["assigned"],
                        "Faltan": u["missing"],
                    }
                    for u in unc_rows
                ]),
                use_container_width=True,
                hide_index=True,
            )

with tab4:
    st.subheader("Turnos de Trabajo y Meta Diaria por Cargo (Crear, Modificar, Reordenar)")
    sel_cargo_t4 = st.selectbox("Selecciona el cargo a configurar", CARGOS_ORDEN, key="t4_sel_cargo")
    cfg_t4 = role_configs[sel_cargo_t4]
    shifts_t4 = get_all_cargo_shifts(cfg_t4)

    st.caption(f"Orden de rotación actual: {' → '.join(cfg_t4['rotation_pool']) if cfg_t4['rotation_pool'] else '—'}")
    for s_idx, s_code in enumerate(shifts_t4):
        in_pool = s_code in cfg_t4["rotation_pool"]
        pool_pos = cfg_t4["rotation_pool"].index(s_code) if in_pool else -1
        c_ord, c_name, c_wk, c_sat, c_sun, c_rot, c_del = st.columns([0.9, 1.5, 0.8, 0.8, 0.8, 0.8, 1.2])

        with c_ord:
            up_b, dn_b = st.columns(2)
            if up_b.button("↑", key=f"t4_up_{sel_cargo_t4}_{s_code}", disabled=(not in_pool or pool_pos == 0)):
                pool = cfg_t4["rotation_pool"]
                pool[pool_pos], pool[pool_pos - 1] = pool[pool_pos - 1], pool[pool_pos]
                st.rerun()
            if dn_b.button("↓", key=f"t4_dn_{sel_cargo_t4}_{s_code}", disabled=(not in_pool or pool_pos == len(cfg_t4["rotation_pool"]) - 1)):
                pool = cfg_t4["rotation_pool"]
                pool[pool_pos], pool[pool_pos + 1] = pool[pool_pos + 1], pool[pool_pos]
                st.rerun()

        new_s_txt = c_name.text_input("Horario", value=s_code, key=f"t4_ren_{sel_cargo_t4}_{s_code}", label_visibility="collapsed")
        cfg_t4["daily_targets"][s_code] = int(c_wk.number_input("Lun-Vie", 0, 30, value=int(cfg_t4["daily_targets"].get(s_code, 0)), key=f"t4_wk_{sel_cargo_t4}_{s_code}"))
        cfg_t4.setdefault("saturday_targets", {})[s_code] = int(c_sat.number_input("Sáb", 0, 30, value=int(cfg_t4.get("saturday_targets", {}).get(s_code, cfg_t4["daily_targets"].get(s_code, 0))), key=f"t4_sat_{sel_cargo_t4}_{s_code}"))
        cfg_t4.setdefault("sunday_targets", {})[s_code] = int(c_sun.number_input("Dom", 0, 30, value=int(cfg_t4.get("sunday_targets", {}).get(s_code, cfg_t4["daily_targets"].get(s_code, 0))), key=f"t4_sun_{sel_cargo_t4}_{s_code}"))

        want_pool = c_rot.checkbox("Rotar", value=in_pool, key=f"t4_inpool_{sel_cargo_t4}_{s_code}")
        if want_pool and not in_pool:
            cfg_t4["rotation_pool"].append(s_code)
            st.rerun()
        elif not want_pool and in_pool and len(cfg_t4["rotation_pool"]) > 1:
            cfg_t4["rotation_pool"].remove(s_code)
            st.rerun()

        with c_del:
            b_ren, b_rm = st.columns(2)
            if b_ren.button("Guardar", key=f"t4_save_sh_{sel_cargo_t4}_{s_code}"):
                parsed_ren = parse_raw_cell(new_s_txt)
                if parsed_ren["base"] and parsed_ren["start_min"] is not None:
                    nc = parsed_ren["base"]
                    if nc != s_code and nc not in shifts_t4:
                        cfg_t4["rotation_pool"] = [nc if x == s_code else x for x in cfg_t4["rotation_pool"]]
                        for bkt in ("daily_targets", "saturday_targets", "sunday_targets"):
                            if s_code in cfg_t4.get(bkt, {}):
                                cfg_t4[bkt][nc] = cfg_t4[bkt].pop(s_code)
                        st.success(f"Turno {s_code} actualizado a {nc}.")
                        st.rerun()
                else:
                    st.error("Formato inválido. Usa HH:MM-HH:MM.")
            if b_rm.button("Quitar", key=f"t4_rm_sh_{sel_cargo_t4}_{s_code}", disabled=(len(shifts_t4) <= 1)):
                cfg_t4["rotation_pool"] = [x for x in cfg_t4["rotation_pool"] if x != s_code]
                for bkt in ("daily_targets", "saturday_targets", "sunday_targets"):
                    cfg_t4.get(bkt, {}).pop(s_code, None)
                st.rerun()

    with st.expander(f"➕ Crear nuevo turno de trabajo para {sel_cargo_t4}"):
        ns1, ns2, ns3, ns4, ns5 = st.columns([1.8, 0.9, 0.9, 0.9, 1.0])
        new_sh_code = ns1.text_input("Nuevo horario (HH:MM-HH:MM)", placeholder="Ej. 14:00-22:00", key="t4_new_sh_code")
        new_sh_wk = int(ns2.number_input("Cant. Lun-Vie", 0, 30, value=1, key="t4_new_sh_wk"))
        new_sh_sat = int(ns3.number_input("Cant. Sábado", 0, 30, value=1, key="t4_new_sh_sat"))
        new_sh_sun = int(ns4.number_input("Cant. Domingo", 0, 30, value=1, key="t4_new_sh_sun"))
        with ns5:
            st.write("")
            st.write("")
            if st.button("Crear turno", type="primary", use_container_width=True, key="t4_create_sh_btn"):
                p_new = parse_raw_cell(new_sh_code)
                if not p_new["base"] or p_new["start_min"] is None:
                    st.error("Ingresa un horario válido HH:MM-HH:MM.")
                elif p_new["base"] in get_all_cargo_shifts(cfg_t4):
                    st.warning("Ese turno ya existe en este cargo.")
                else:
                    cb = p_new["base"]
                    cfg_t4["rotation_pool"].append(cb)
                    cfg_t4["daily_targets"][cb] = new_sh_wk
                    cfg_t4.setdefault("saturday_targets", {})[cb] = new_sh_sat
                    cfg_t4.setdefault("sunday_targets", {})[cb] = new_sh_sun
                    st.success(f"Turno {cb} agregado a {sel_cargo_t4}.")
                    st.rerun()

    if st.button("Aplicar reglas, turnos y Re-rotar semana activa", type="primary", disabled=not has_weeks, key="t4_apply_rerotate"):
        lbl_r, at_r = trigger_rerotate_active_week()
        st.success(f"Reglas y metas aplicadas a '{lbl_r}' ({at_r} turno(s) AT).")
        st.rerun()

    st.divider()
    st.subheader("Directorio de Personal (Activar / Desactivar Recursos y Habilitación AT)")
    d_f1, d_f2, d_f3 = st.columns([2, 1.5, 1.2])
    dir_q = d_f1.text_input("Buscar en directorio", placeholder="Código o nombre...", key="t4_dir_q")
    dir_cargo = d_f2.selectbox("Cargo", ["TODOS"] + CARGOS_ORDEN, key="t4_dir_cargo")
    dir_estado = d_f3.selectbox("Estado", ["Todos", "Activos", "Inactivos"], key="t4_dir_est")

    with st.expander("➕ Agregar nuevo colaborador al directorio"):
        ne1, ne2, ne3, ne4, ne5 = st.columns([1.0, 2.0, 1.8, 0.9, 1.0])
        ne_cod = ne1.text_input("Código", key="t4_ne_cod")
        ne_nom = ne2.text_input("Nombre completo", key="t4_ne_nom")
        ne_car = ne3.selectbox("Cargo", CARGOS_ORDEN, index=1, key="t4_ne_car")
        ne_at = ne4.checkbox("AT", value=False, key="t4_ne_at", disabled=(ne_car not in ("Analista de Operaciones", "Auxiliar de Operaciones")))
        with ne5:
            st.write("")
            st.write("")
            if st.button("Agregar", type="primary", key="t4_ne_btn", use_container_width=True):
                c_clean, n_clean = ne_cod.strip(), ne_nom.strip()
                if not c_clean or not n_clean:
                    st.error("Completa código y nombre.")
                elif any(x["codigo"] == c_clean or x["id"] == c_clean for x in employees):
                    st.error(f"Ya existe el código {c_clean}.")
                else:
                    employees.append({
                        "id": c_clean,
                        "codigo": c_clean,
                        "nombre": n_clean,
                        "cargo": ne_car,
                        "puede_at": bool(ne_at and ne_car in ("Analista de Operaciones", "Auxiliar de Operaciones")),
                        "activo": True,
                    })
                    st.success(f"{n_clean} agregado al directorio.")
                    st.rerun()

    dir_list = sorted(
        employees,
        key=lambda x: (0 if x["activo"] else 1, CARGOS_ORDEN.index(x["cargo"]) if x["cargo"] in CARGOS_ORDEN else 99, x["nombre"]),
    )
    if dir_cargo != "TODOS":
        dir_list = [x for x in dir_list if x["cargo"] == dir_cargo]
    if dir_estado == "Activos":
        dir_list = [x for x in dir_list if x["activo"]]
    elif dir_estado == "Inactivos":
        dir_list = [x for x in dir_list if not x["activo"]]
    if dir_q.strip():
        dq = dir_q.strip().lower()
        dir_list = [x for x in dir_list if dq in x["nombre"].lower() or dq in x["codigo"].lower()]

    for e in dir_list:
        ec1, ec2, ec3, ec4, ec5, ec6 = st.columns([0.8, 1.9, 1.8, 1.0, 0.9, 1.2])
        ec1.write(e["codigo"])
        ec2.write(e["nombre"] if e["activo"] else f"~~{e['nombre']}~~")
        ec3.write(e["cargo"])
        if e["cargo"] in ("Analista de Operaciones", "Auxiliar de Operaciones"):
            e["puede_at"] = ec4.checkbox("Habilitado AT", value=e["puede_at"], disabled=not e["activo"], key=f"t4_at_{e['id']}")
        else:
            ec4.caption("—")
        ec5.write("Activo" if e["activo"] else "Inactivo")
        btn_lbl = "Desactivar" if e["activo"] else "Reactivar"
        if ec6.button(btn_lbl, key=f"t4_toggle_{e['id']}", use_container_width=True):
            e["activo"] = not e["activo"]
            st.rerun()

    st.divider()
    st.subheader("Novedades Programadas y Reglas Automáticas")
    nov_default_start = active_week["start_date"] if active_week else current_monday_iso()
    render_novedad_form("t4_nov", employees, nov_default_start, allow_rerotate=True, active_week_exists=has_weeks)

    if st.session_state["scheduled_novedades"]:
        for idx_n, nov in enumerate(st.session_state["scheduled_novedades"]):
            emp_n = next((e for e in employees if e["id"] == nov["emp_id"]), None)
            emp_str = f"{emp_n['codigo']} {emp_n['nombre']}" if emp_n else nov["emp_id"]
            c_n1, c_n2 = st.columns([5, 1])
            bg_n, fg_n, bd_n = get_novedad_badge_colors(nov["tipo"])
            c_n1.markdown(
                f"<div style='padding:6px 10px;border-radius:6px;background:{bg_n};color:{fg_n};border:1px solid {bd_n};font-size:12px;'>"
                f"<strong>{emp_str}</strong> · <strong>{nov['tipo']}</strong> · {nov['start_date']} → {nov['end_date']}"
                "</div>",
                unsafe_allow_html=True,
            )
            if c_n2.button("Eliminar", key=f"t4_del_nov_{idx_n}"):
                st.session_state["scheduled_novedades"].pop(idx_n)
                st.rerun()

    r1, r2, r3 = st.columns(3)
    rules["min_rest_between"] = int(r1.number_input("Descanso mínimo entre turnos (h)", 8, 16, value=int(rules["min_rest_between"]), key="t4_min_rest_between"))
    rules["min_rest_off"] = int(r2.number_input("Descanso en Día Libre L (h)", 12, 48, value=int(rules["min_rest_off"]), key="t4_min_rest_off"))
    rules["min_work_between_off"] = int(r3.number_input("Mín. días trabajados entre descansos", 3, 6, value=int(rules["min_work_between_off"]), key="t4_min_work"))

    r4, r5, r6 = st.columns(3)
    rules["max_consecutive"] = int(r4.number_input("Máx. días seguidos sin descanso", 5, 10, value=int(rules["max_consecutive"]), key="t4_max_cons"))
    with r5:
        st.write("")
        st.write("")
        rules["rotate_after_off"] = st.checkbox("Rotar turno después del Día Libre (L)", value=bool(rules["rotate_after_off"]), key="t4_rot_off")
    with r6:
        st.write("")
        st.write("")
        rules["auto_assign_at"] = st.checkbox("Auto-asignar cobertura AT al generar semana", value=bool(rules["auto_assign_at"]), key="t4_auto_at")

with tab5:
    st.subheader("Cargar Programación como Semana Activa (Excel o Pegado)")
    st.caption("Cuando subas un archivo o pegues la tabla de Excel aquí, quedará establecida automáticamente como tu Semana Activa.")
    imp_date = st.date_input("Lunes de inicio de la semana importada", value=dt.date.fromisoformat(current_monday_iso()))
    imp_iso = imp_date.strftime("%Y-%m-%d")

    col_up, col_paste = st.columns(2)
    with col_up:
        st.markdown("**Opción 1: Subir archivo Excel (.xlsx / .xlsm / .xls)**")
        up_file = st.file_uploader("Sube un archivo Excel con columnas CODIGO, NOMBRE, CARGO y los 7 días", type=["xlsx", "xlsm", "xls"])
        if up_file is not None and st.button("Cargar archivo Excel como Semana Activa", type="primary"):
            try:
                df_in = pd.read_excel(up_file)
                new_emps, new_week = parse_dataframe_to_roster(df_in, imp_iso)
                st.session_state["employees"] = new_emps
                st.session_state["weeks"] = [new_week]
                st.session_state["active_week_idx"] = 0
                st.session_state["rerotate_offset"] = 0
                st.success(f"Cargados {len(new_emps)} colaboradores. '{new_week['label']}' es ahora tu Semana Activa.")
                st.rerun()
            except Exception as exc:
                st.error(f"Error al importar Excel: {exc}")

    with col_paste:
        st.markdown("**Opción 2: Pegar celdas desde Excel (Ctrl + C / Ctrl + V)**")
        pasted = st.text_area("Pega aquí desde CODIGO hasta DOMINGO", height=180, placeholder="CODIGO\tNOMBRE\tCARGO\tLUNES\tMARTES\tMIÉRCOLES\tJUEVES\tVIERNES\tSÁBADO\tDOMINGO\n12896\tAlexis Mojica\tAnalista de Operaciones\t...")
        if st.button("Cargar tabla pegada como Semana Activa", type="primary"):
            try:
                new_emps, new_week = parse_pasted_text_to_roster(pasted, imp_iso)
                st.session_state["employees"] = new_emps
                st.session_state["weeks"] = [new_week]
                st.session_state["active_week_idx"] = 0
                st.session_state["rerotate_offset"] = 0
                st.success(f"Cargados {len(new_emps)} colaboradores desde texto pegado. '{new_week['label']}' es ahora tu Semana Activa.")
                st.rerun()
            except Exception as exc:
                st.error(f"Error al interpretar texto pegado: {exc}")

    st.divider()
    if st.button("Cargar plantilla de ejemplo (opcional)", key="t5_load_seed"):
        st.session_state["employees"] = copy.deepcopy(SEED_EMPLOYEES)
        st.session_state["weeks"] = [seed_week()]
        st.session_state["active_week_idx"] = 0
        st.session_state["rerotate_offset"] = 0
        st.success("Se cargó la semana de ejemplo como Semana Activa.")
        st.rerun()
