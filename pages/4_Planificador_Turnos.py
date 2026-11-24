import datetime
import html
import io
import re
from typing import Dict, List, Optional, Tuple

import openpyxl
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

st.set_page_config(
    page_title="Planificador de Turnos Operativos",
    page_icon="📅",
    layout="wide",
)

# ==============================================================================
# CONSTANTES Y CONFIGURACIÓN BASE
# ==============================================================================

CARGOS_ORDEN = [
    "Analista de Operaciones",
    "Auxiliar de Operaciones",
    "Operador Líder",
    "Técnico de Operaciones",
    "Auxiliar de Alistamiento",
    "Conductor de Operaciones",
]

DIAS_SEMANA = [
    "LUNES",
    "MARTES",
    "MIÉRCOLES",
    "JUEVES",
    "VIERNES",
    "SÁBADO",
    "DOMINGO",
]

TURNOS_CATALOGO = [
    "03:00-11:00",
    "06:00-15:00",
    "08:00-17:00",
    "11:00-19:00",
    "19:00-27:00",
    "23:00-31:00",
    "20:00-28:00",
    "22:00-30:00",
    "03:00-10:00",
    "11:00-18:00",
]

ESTADOS_ESPECIALES = [
    "L",
    "VACACIONES",
    "SUSPEND",
    "INCAPACIDAD",
    "PERMISO",
]

MESES_ES = [
    "ene", "feb", "mar", "abr", "may", "jun",
    "jul", "ago", "sep", "oct", "nov", "dic",
]

# Secuencia estándar: Noche -> Tarde -> Mañana
DEFAULT_ROLE_CONFIGS = {
    "Analista de Operaciones": {
        "rotation_pool": ["19:00-27:00", "11:00-19:00", "06:00-15:00", "03:00-11:00"],
        "daily_targets": {
            "19:00-27:00": 1,
            "11:00-19:00": 1,
            "06:00-15:00": 1,
            "03:00-11:00": 1,
        },
        "saturday_targets": {},
        "sunday_targets": {},
    },
    "Auxiliar de Operaciones": {
        "rotation_pool": ["19:00-27:00", "23:00-31:00", "11:00-19:00", "03:00-11:00"],
        "daily_targets": {
            "19:00-27:00": 1,
            "23:00-31:00": 1,
            "11:00-19:00": 2,
            "03:00-11:00": 2,
        },
        "saturday_targets": {},
        "sunday_targets": {},
    },
    "Operador Líder": {
        "rotation_pool": ["20:00-28:00"],
        "daily_targets": {"20:00-28:00": 1},
        "saturday_targets": {},
        "sunday_targets": {},
    },
    "Técnico de Operaciones": {
        "rotation_pool": ["19:00-27:00", "11:00-19:00", "03:00-11:00"],
        "daily_targets": {
            "19:00-27:00": 1,
            "11:00-19:00": 1,
            "03:00-11:00": 1,
        },
        "saturday_targets": {},
        "sunday_targets": {},
    },
    "Auxiliar de Alistamiento": {
        "rotation_pool": ["22:00-30:00", "20:00-28:00", "08:00-17:00"],
        "daily_targets": {
            "22:00-30:00": 1,
            "20:00-28:00": 2,
            "08:00-17:00": 1,
        },
        "saturday_targets": {},
        "sunday_targets": {},
    },
    "Conductor de Operaciones": {
        "rotation_pool": ["11:00-18:00", "03:00-10:00"],
        "daily_targets": {
            "11:00-18:00": 1,
            "03:00-10:00": 1,
        },
        "saturday_targets": {},
        "sunday_targets": {},
    },
}

DEFAULT_RULES = {
    "min_rest_between_shifts": 12,
    "min_rest_day_off": 24,
    "days_off_per_week": 1,
    "min_work_days_between_off": 5,
    "max_consecutive_work_days": 7,
    "rotate_after_day_off": True,
    "auto_assign_at": True,
}

SEED_ROWS = [
    ("669", "Roberto Velarde", "Analista de Operaciones", False, True,
     ["L", "06:00-15:00", "06:00-15:00", "06:00-15:00 CAP", "08:00-17:00", "11:00-19:00", "08:00-17:00"]),
    ("12520", "Keily Agrazal", "Analista de Operaciones", True, True,
     ["11:00-19:00", "11:00-19:00", "11:00-19:00", "11:00-19:00", "L", "19:00-27:00", "19:00-27:00"]),
    ("13563", "Kathleen Batista", "Analista de Operaciones", True, True,
     ["VACACIONES", "VACACIONES", "VACACIONES", "VACACIONES", "VACACIONES", "L", "03:00-11:00"]),
    ("3332", "Kevin Pitty", "Analista de Operaciones", False, True,
     ["19:00-27:00", "19:00-27:00", "19:00-27:00", "19:00-27:00", "19:00-27:00", "L", "11:00-19:00"]),
    ("12896", "Alexis Mojica", "Analista de Operaciones", True, True,
     ["03:00-11:00 AT", "03:00-11:00 AT", "03:00-11:00 AT", "03:00-11:00", "03:00-11:00 AT", "03:00-11:00", "L"]),
    ("3393", "David Vega", "Auxiliar de Operaciones", True, True,
     ["L", "SUSPEND", "11:00-19:00", "11:00-19:00", "11:00-27:00", "11:00-19:00", "11:00-19:00 AT"]),
    ("4211", "Dariel Cardales", "Auxiliar de Operaciones", True, True,
     ["03:00-11:00", "03:00-11:00", "03:00-11:00", "03:00-11:00 AT", "SUSPEND", "03:00-11:00", "L"]),
    ("4607", "Nelson Mojica", "Auxiliar de Operaciones", True, True,
     ["11:00-19:00", "11:00-19:00", "11:00-19:00", "11:00-19:00 AT/CAP", "11:00-19:00", "11:00-19:00 AT", "L"]),
    ("12717", "Francisco Delgado", "Auxiliar de Operaciones", True, True,
     ["VACACIONES", "VACACIONES", "VACACIONES", "VACACIONES", "VACACIONES", "03:00-11:00 AT", "03:00-11:00"]),
    ("13580", "Juan Patiño", "Auxiliar de Operaciones", False, True,
     ["19:00-27:00", "19:00-27:00", "19:00-27:00", "19:00-27:00", "L", "23:00-35:00", "23:00-31:00"]),
    ("13488", "Edwin Aranda", "Auxiliar de Operaciones", True, True,
     ["11:00-19:00", "11:00-19:00", "L", "03:00-11:00", "03:00-11:00", "19:00-27:00", "19:00-27:00"]),
    ("12644", "Idelfonso Sánchez", "Auxiliar de Operaciones", False, True,
     ["03:00-11:00", "03:00-11:00", "03:00-11:00", "03:00-11:00", "03:00-11:00", "L", "03:00-11:00"]),
    ("81569", "Kathleen Amat", "Auxiliar de Operaciones", False, True,
     ["23:00-31:00", "23:00-31:00", "23:00-31:00", "23:00-31:00", "23:00-31:00", "L", "11:00-19:00"]),
    ("759", "Daniel Solis", "Operador Líder", False, True,
     ["20:00-28:00", "20:00-28:00", "20:00-28:00", "20:00-28:00", "20:00-28:00", "L", "20:00-28:00"]),
    ("10043", "Gilberto Torrero", "Operador Líder", False, True,
     ["20:00-28:00", "20:00-28:00", "20:00-28:00", "20:00-28:00", "L", "20:00-28:00", "20:00-28:00"]),
    ("671", "Luis Guardao", "Técnico de Operaciones", False, True,
     ["03:00-11:00", "03:00-11:00", "03:00-11:00 CAP", "03:00-11:00", "L", "19:00-27:00", "19:00-27:00"]),
    ("12560", "Estefany Villamil", "Técnico de Operaciones", False, True,
     ["19:00-27:00", "19:00-27:00", "19:00-27:00", "19:00-27:00", "19:00-27:00", "L", "11:00-19:00"]),
    ("11383", "Kathia Gonzalez", "Técnico de Operaciones", False, True,
     ["11:00-19:00", "11:00-19:00", "11:00-19:00", "11:00-19:00", "11:00-19:00", "11:00-19:00", "L"]),
    ("13479", "Edier Barrios", "Técnico de Operaciones", False, True,
     ["11:00-19:00", "11:00-19:00", "11:00-19:00", "L", "03:00-19:00", "03:00-11:00", "03:00-11:00"]),
    ("590", "Julio Aguilar", "Auxiliar de Alistamiento", False, True,
     ["20:00-28:00", "20:00-28:00", "20:00-28:00", "20:00-28:00", "20:00-28:00", "L", "20:00-28:00"]),
    ("7144", "Darjil Ortega", "Auxiliar de Alistamiento", False, True,
     ["22:00-30:00", "22:00-30:00", "22:00-30:00", "L", "20:00-28:00", "20:00-28:00", "20:00-28:00"]),
    ("11584", "Alberto Puchie", "Auxiliar de Alistamiento", False, True,
     ["08:00-17:00", "08:00-17:00", "L", "22:00-30:00", "22:00-30:00", "20:00-28:00", "20:00-28:00"]),
    ("11589", "Nelson Carrasquilla", "Auxiliar de Alistamiento", False, True,
     ["VACACIONES", "VACACIONES", "VACACIONES", "VACACIONES", "VACACIONES", "08:00-17:00", "08:00-17:00"]),
    ("12933", "Jose Canto", "Auxiliar de Alistamiento", False, True,
     ["20:00-28:00", "20:00-28:00", "20:00-28:00", "20:00-28:00", "L", "22:00-30:00", "22:00-30:00"]),
    ("12936", "Eric Batista", "Auxiliar de Alistamiento", False, False,
     ["", "", "", "", "", "", ""]),
    ("13618", "Isidro Concepcion", "Conductor de Operaciones", False, True,
     ["11:00-18:00", "11:00-18:00", "11:00-18:00", "11:00-18:00", "11:00-18:00", "11:00-18:00", "L"]),
    ("13635", "Ricardo Palacio", "Conductor de Operaciones", False, True,
     ["03:00-10:00", "03:00-10:00", "03:00-10:00", "03:00-10:00", "03:00-10:00", "L", "11:00-18:00"]),
]


# ==============================================================================
# UTILIDADES DE FECHAS, CELDAS Y COLORES
# ==============================================================================

def add_days_iso(iso_date: str, days: int) -> str:
    dt = datetime.date.fromisoformat(iso_date) + datetime.timedelta(days=days)
    return dt.isoformat()


def format_short_date_es(iso_date: str) -> str:
    dt = datetime.date.fromisoformat(iso_date)
    return f"{dt.day}-{MESES_ES[dt.month - 1]}"


def build_week_label(start_iso: str) -> str:
    end_iso = add_days_iso(start_iso, 6)
    return f"Semana {format_short_date_es(start_iso)} — {format_short_date_es(end_iso)}"


def parse_raw_cell(raw_text: str, locked: bool = False) -> dict:
    clean = re.sub(r"\s+", " ", str(raw_text or "").strip().upper())
    if not clean or clean in ("NAN", "NONE"):
        return {"raw": "", "baseShift": "", "isAT": False, "isCAP": False, "locked": locked}

    tokens = clean.split(" ")
    first = tokens[0]
    rest = " ".join(tokens[1:])

    is_at = bool(re.search(r"\bAT\b|AT/CAP|CAP/AT", clean))
    is_cap = bool(re.search(r"\bCAP\b|AT/CAP|CAP/AT", clean))

    if first in ("L", "LIBRE", "VACACIONES", "VAC", "SUSPEND", "SUSPENDIDO", "INCAPACIDAD", "PERMISO"):
        norm = (
            "L" if first == "LIBRE"
            else "VACACIONES" if first == "VAC"
            else "SUSPEND" if first == "SUSPENDIDO"
            else first
        )
        return {
            "raw": norm,
            "baseShift": norm,
            "isAT": False,
            "isCAP": False,
            "locked": bool(locked or norm in ("VACACIONES", "SUSPEND", "INCAPACIDAD", "PERMISO")),
        }

    base_shift = first
    suffix = []
    if is_at and is_cap:
        suffix.append("AT/CAP")
    else:
        if is_at:
            suffix.append("AT")
        if is_cap:
            suffix.append("CAP")

    if suffix:
        final_raw = f"{base_shift} {' '.join(suffix)}"
    elif rest:
        final_raw = clean
    else:
        final_raw = base_shift

    return {
        "raw": final_raw,
        "baseShift": base_shift,
        "isAT": is_at,
        "isCAP": is_cap,
        "locked": bool(locked),
    }


def format_cell_raw(base_shift: str, is_at: bool = False, is_cap: bool = False) -> str:
    base = str(base_shift or "").strip().upper()
    if not base:
        return ""
    if base in ESTADOS_ESPECIALES:
        return base
    if is_at and is_cap:
        return f"{base} AT/CAP"
    if is_at:
        return f"{base} AT"
    if is_cap:
        return f"{base} CAP"
    return base


def parse_shift_window(base_shift: str) -> Optional[Tuple[int, int, int]]:
    s = str(base_shift or "").strip().upper()
    m = re.match(r"^(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})$", s)
    if not m:
        return None
    sh, sm, eh, em = int(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(4))
    if sm < 0 or sm > 59 or em < 0 or em > 59:
        return None
    start_min = sh * 60 + sm
    end_min = eh * 60 + em
    if end_min <= start_min:
        end_min += 1440
    return start_min, end_min, end_min - start_min


def is_non_working_state(base_shift: str) -> bool:
    return str(base_shift or "").strip().upper() in ("", "L", "VACACIONES", "SUSPEND", "INCAPACIDAD", "PERMISO")


def is_tech_shift_no_at(shift_code: str) -> bool:
    return str(shift_code or "").strip().startswith("19:00")


def get_shift_target_for_day(cfg: dict, shift_code: str, day_idx: int) -> int:
    if not cfg:
        return 0
    if day_idx == 5:
        sat_map = cfg.get("saturday_targets") or {}
        if shift_code in sat_map:
            return int(sat_map[shift_code])
    if day_idx == 6:
        sun_map = cfg.get("sunday_targets") or {}
        if shift_code in sun_map:
            return int(sun_map[shift_code])
    return int((cfg.get("daily_targets") or {}).get(shift_code, 0))


def covers_shift_window(assigned_shift: str, required_shift: str) -> bool:
    if assigned_shift == required_shift:
        return True
    w_a = parse_shift_window(assigned_shift)
    w_r = parse_shift_window(required_shift)
    if not w_a or not w_r:
        return False
    return w_a[0] <= w_r[0] and w_a[1] >= w_r[1]


def get_cell_colors(cell: dict) -> Tuple[str, str, str]:
    base = str(cell.get("baseShift", "")).strip().upper()
    if base == "SUSPEND":
        return "#DC2626", "#FFFFFF", "DC2626"
    if base == "VACACIONES":
        return "#FFFFFF", "#DC2626", "FFFFFF"
    if base in ("INCAPACIDAD", "PERMISO"):
        return "#FFE4E6", "#9F1239", "FFE4E6"
    if base == "L":
        return "#FCE4D6", "#C00000", "FCE4D6"
    if cell.get("isAT") or base in ("03:00-19:00", "11:00-27:00", "23:00-35:00"):
        return "#F9A8D4", "#0F172A", "F4B8E4"
    if base.startswith("03:00-"):
        return "#92D050", "#0F172A", "92D050"
    if base.startswith("11:00-"):
        return "#FFC000", "#0F172A", "FFC000"
    if base.startswith("19:00-"):
        return "#B4C6E7", "#0F172A", "B4C6E7"
    if base.startswith("23:00-"):
        return "#FFF2CC", "#0F172A", "FFF2CC"
    if base.startswith("20:00-"):
        return "#00B0F0", "#0F172A", "00B0F0"
    if base.startswith("22:00-"):
        return "#A6A6A6", "#0F172A", "A6A6A6"
    if base in ("06:00-15:00", "08:00-17:00"):
        return "#FFFFFF", "#0F172A", "FFFFFF"
    return "#F8FAFC", "#334155", "F8FAFC"


def compute_transition_rest_hours(
    prev_cell: Optional[dict],
    prev_day_idx: int,
    next_cell: Optional[dict],
    next_day_idx: int,
) -> Optional[float]:
    if not prev_cell or not next_cell:
        return None
    w_prev = parse_shift_window(prev_cell.get("baseShift", ""))
    w_next = parse_shift_window(next_cell.get("baseShift", ""))
    if not w_prev or not w_next:
        return None
    end_abs = prev_day_idx * 1440 + w_prev[1]
    start_abs = next_day_idx * 1440 + w_next[0]
    return round((start_abs - end_abs) / 60.0, 1)


def is_transition_valid(
    last_shift: Optional[str],
    last_day_offset: int,
    cand_shift: str,
    cand_day: int,
    crossed_off: bool,
    rules: dict,
) -> bool:
    if not last_shift:
        return True
    w_prev = parse_shift_window(last_shift)
    w_next = parse_shift_window(cand_shift)
    if not w_prev or not w_next:
        return True
    rest_h = (cand_day * 1440 + w_next[0] - (last_day_offset * 1440 + w_prev[1])) / 60.0
    req = rules["min_rest_day_off"] if crossed_off else rules["min_rest_between_shifts"]
    return rest_h >= req


def get_primary_shift(cells: List[dict], fallback: str) -> str:
    counts: Dict[str, int] = {}
    for c in cells:
        b = c.get("baseShift", "")
        if parse_shift_window(b):
            counts[b] = counts.get(b, 0) + 1
    if not counts:
        return fallback
    return max(counts.items(), key=lambda kv: kv[1])[0]


def apply_novedad_to_week(week: dict, novedad: dict) -> Tuple[dict, int]:
    emp_id = novedad["emp_id"]
    start_iso = novedad["start_iso"]
    end_iso = novedad["end_iso"]
    tipo = novedad["tipo"]

    new_assign = {k: [dict(c) for c in v] for k, v in week["assignments"].items()}
    row = new_assign.get(emp_id) or [parse_raw_cell("") for _ in range(7)]
    applied_days = 0
    for d in range(7):
        d_iso = add_days_iso(week["start_date"], d)
        if start_iso <= d_iso <= end_iso:
            row[d] = parse_raw_cell(tipo, locked=True)
            applied_days += 1

    if applied_days == 0:
        return week, 0

    new_assign[emp_id] = row
    return {
        "id": week["id"],
        "label": week["label"],
        "start_date": week["start_date"],
        "assignments": new_assign,
    }, applied_days


# ==============================================================================
# MOTOR DE VALIDACIÓN, COBERTURA AT, FALTANTES Y GENERACIÓN / RE-ROTACIÓN
# ==============================================================================

def compute_uncovered_tasks(week: dict, employees: List[dict], role_configs: dict) -> List[dict]:
    uncovered = []
    for cargo in CARGOS_ORDEN:
        cfg = role_configs.get(cargo)
        if not cfg:
            continue
        shifts = list(dict.fromkeys(
            list(cfg.get("rotation_pool", [])) + list((cfg.get("daily_targets") or {}).keys())
        ))
        for d in range(7):
            for s_code in shifts:
                target = get_shift_target_for_day(cfg, s_code, d)
                if target <= 0:
                    continue
                no_at = cargo == "Técnico de Operaciones" and is_tech_shift_no_at(s_code)
                assigned = 0
                for emp in employees:
                    if not emp["activo"]:
                        continue
                    cell = (week["assignments"].get(emp["id"]) or [])[d] if emp["id"] in week["assignments"] else None
                    if not cell or not cell.get("baseShift"):
                        continue
                    if emp["cargo"] == cargo and covers_shift_window(cell["baseShift"], s_code):
                        assigned += 1
                    elif (
                        cargo == "Técnico de Operaciones"
                        and not no_at
                        and cell.get("isAT")
                        and covers_shift_window(cell["baseShift"], s_code)
                    ):
                        assigned += 1
                if assigned < target:
                    uncovered.append({
                        "day_idx": d,
                        "day_name": DIAS_SEMANA[d],
                        "cargo": cargo,
                        "shift": s_code,
                        "assigned": assigned,
                        "target": target,
                        "missing": target - assigned,
                        "reason": (
                            "Requiere Técnico directo (19:00 no admite AT)"
                            if no_at
                            else "Personal insuficiente, novedades o descanso 12h/24h"
                        ),
                    })
    return uncovered


def validate_schedule(
    week: dict,
    employees: List[dict],
    role_configs: dict,
    rules: dict,
    prev_week: Optional[dict] = None,
) -> List[dict]:
    issues = []
    for emp in employees:
        if not emp["activo"]:
            continue
        cells = week["assignments"].get(emp["id"], [])
        prev_cells = prev_week["assignments"].get(emp["id"], []) if prev_week else []

        last_work_cell = None
        last_work_offset = -1
        saw_off_since_last = False

        if len(prev_cells) == 7:
            for pd_idx in range(6, -1, -1):
                c = prev_cells[pd_idx]
                if not c:
                    continue
                if c["baseShift"] == "L":
                    saw_off_since_last = True
                    continue
                if parse_shift_window(c["baseShift"]):
                    last_work_cell = c
                    last_work_offset = pd_idx - 7
                    break
                if is_non_working_state(c["baseShift"]):
                    break

        dias_libres = 0
        dias_trabajados = 0
        streak = 0
        if len(prev_cells) == 7:
            for pd_idx in range(6, -1, -1):
                c = prev_cells[pd_idx]
                if not c or not parse_shift_window(c["baseShift"]):
                    break
                streak += 1

        streak_emitted = False

        for d in range(7):
            cell = cells[d] if d < len(cells) else parse_raw_cell("")
            b = cell.get("baseShift", "")
            if b == "L":
                dias_libres += 1
                saw_off_since_last = True
                streak = 0
                continue
            if b in ("VACACIONES", "SUSPEND", "INCAPACIDAD", "PERMISO", ""):
                last_work_cell = None
                saw_off_since_last = False
                streak = 0
                continue

            win = parse_shift_window(b)
            if win:
                dias_trabajados += 1
                streak += 1
                if not streak_emitted and rules["max_consecutive_work_days"] > 0 and streak > rules["max_consecutive_work_days"]:
                    streak_emitted = True
                    issues.append({
                        "emp_id": emp["id"],
                        "cargo": emp["cargo"],
                        "day_idx": d,
                        "type": "CONSECUTIVE_DAYS",
                        "message": f"{emp['nombre']} ({emp['codigo']}): {streak} días seguidos sin L (máx. {rules['max_consecutive_work_days']}).",
                    })

                if last_work_cell:
                    rh = compute_transition_rest_hours(last_work_cell, last_work_offset, cell, d)
                    if rh is not None:
                        if saw_off_since_last and rh < rules["min_rest_day_off"]:
                            issues.append({
                                "emp_id": emp["id"],
                                "cargo": emp["cargo"],
                                "day_idx": d,
                                "type": "REST_24H_LIBRE",
                                "message": f"{emp['nombre']} ({emp['codigo']}): solo {rh}h tras Día Libre (mín. {rules['min_rest_day_off']}h).",
                            })
                        elif not saw_off_since_last and rh < rules["min_rest_between_shifts"]:
                            issues.append({
                                "emp_id": emp["id"],
                                "cargo": emp["cargo"],
                                "day_idx": d,
                                "type": "REST_12H",
                                "message": f"{emp['nombre']} ({emp['codigo']}): {rh}h entre {last_work_cell['baseShift']} y {b} (mín. {rules['min_rest_between_shifts']}h).",
                            })
                last_work_cell = cell
                last_work_offset = d
                saw_off_since_last = False

        if dias_trabajados >= 6 and dias_libres < rules["days_off_per_week"]:
            issues.append({
                "emp_id": emp["id"],
                "cargo": emp["cargo"],
                "day_idx": 6,
                "type": "NO_DAY_OFF",
                "message": f"{emp['nombre']} ({emp['codigo']}) tiene {dias_libres} día(s) libre(s) esta semana.",
            })

    tech_cfg = role_configs.get("Técnico de Operaciones", {})
    tech_shifts = list(dict.fromkeys(
        list(tech_cfg.get("rotation_pool", [])) + list((tech_cfg.get("daily_targets") or {}).keys())
    ))
    for d in range(7):
        for s_code in tech_shifts:
            tgt = get_shift_target_for_day(tech_cfg, s_code, d)
            if tgt <= 0:
                continue
            no_at = is_tech_shift_no_at(s_code)
            covered = 0
            for emp in employees:
                if not emp["activo"]:
                    continue
                c = (week["assignments"].get(emp["id"]) or [])[d] if emp["id"] in week["assignments"] else None
                if not c or not c.get("baseShift"):
                    continue
                if emp["cargo"] == "Técnico de Operaciones" and covers_shift_window(c["baseShift"], s_code):
                    covered += 1
                elif c.get("isAT") and not no_at and covers_shift_window(c["baseShift"], s_code):
                    covered += 1
            if covered < tgt:
                issues.append({
                    "emp_id": None,
                    "cargo": "Técnico de Operaciones",
                    "day_idx": d,
                    "type": "UNCOVERED_TECH",
                    "shift": s_code,
                    "message": (
                        f"Turno {s_code} Técnico sin Técnico directo ({covered}/{tgt}; no admite AT)."
                        if no_at
                        else f"Turno {s_code} Técnico sin cobertura suficiente ({covered}/{tgt})."
                    ),
                })
    return issues


def auto_assign_tecnico_at(
    week: dict,
    employees: List[dict],
    role_configs: dict,
    rules: dict,
    prev_week: Optional[dict] = None,
) -> Tuple[dict, int]:
    next_assign = {k: [dict(c) for c in v] for k, v in week["assignments"].items()}

    for emp in employees:
        if not emp["activo"] or emp["cargo"] == "Técnico de Operaciones":
            continue
        row = next_assign.get(emp["id"])
        if not row:
            continue
        for d in range(7):
            if row[d].get("isAT") and not row[d].get("locked"):
                row[d] = parse_raw_cell(format_cell_raw(row[d]["baseShift"], False, row[d]["isCAP"]))

    tech_cfg = role_configs.get("Técnico de Operaciones", {})
    tech_shifts = list(dict.fromkeys(
        list(tech_cfg.get("rotation_pool", [])) + list((tech_cfg.get("daily_targets") or {}).keys())
    ))

    at_totals = {
        emp["id"]: sum(1 for c in next_assign.get(emp["id"], []) if c.get("isAT"))
        for emp in employees
    }
    assigned_count = 0

    for d in range(7):
        for s_code in tech_shifts:
            tgt = get_shift_target_for_day(tech_cfg, s_code, d)
            if tgt <= 0 or is_tech_shift_no_at(s_code):
                continue

            def count_covered() -> int:
                cnt = 0
                for emp in employees:
                    if not emp["activo"]:
                        continue
                    c = (next_assign.get(emp["id"]) or [])[d]
                    if emp["cargo"] == "Técnico de Operaciones" and covers_shift_window(c["baseShift"], s_code):
                        cnt += 1
                    elif c.get("isAT") and covers_shift_window(c["baseShift"], s_code):
                        cnt += 1
                return cnt

            covered = count_covered()
            while covered < tgt:
                same_shift = [
                    e for e in employees
                    if e["activo"]
                    and e.get("puede_at")
                    and e["cargo"] in ("Auxiliar de Operaciones", "Analista de Operaciones")
                    and not next_assign[e["id"]][d].get("isAT")
                    and not next_assign[e["id"]][d].get("locked")
                    and next_assign[e["id"]][d]["baseShift"] == s_code
                ]
                same_shift.sort(key=lambda e: (at_totals.get(e["id"], 0), 0 if e["cargo"] == "Auxiliar de Operaciones" else 1))

                if same_shift:
                    chosen = same_shift[0]
                    old = next_assign[chosen["id"]][d]
                    next_assign[chosen["id"]][d] = parse_raw_cell(
                        format_cell_raw(old["baseShift"], True, old["isCAP"])
                    )
                    at_totals[chosen["id"]] = at_totals.get(chosen["id"], 0) + 1
                    assigned_count += 1
                    covered += 1
                    continue

                flex_candidates = []
                for e in employees:
                    if not e["activo"] or not e.get("puede_at"):
                        continue
                    if e["cargo"] not in ("Auxiliar de Operaciones", "Analista de Operaciones"):
                        continue
                    row = next_assign.get(e["id"], [])
                    cell = row[d] if d < len(row) else None
                    if not cell or cell.get("isAT") or cell.get("locked") or is_non_working_state(cell["baseShift"]):
                        continue

                    prev_s, prev_off, prev_crossed = None, d - 1, False
                    for k in range(d - 1, -1, -1):
                        if row[k]["baseShift"] == "L":
                            prev_crossed = True
                            continue
                        if parse_shift_window(row[k]["baseShift"]):
                            prev_s = row[k]["baseShift"]
                            prev_off = k
                            break
                        break
                    if not prev_s and d == 0 and prev_week and e["id"] in prev_week["assignments"]:
                        p_row = prev_week["assignments"][e["id"]]
                        for k in range(6, -1, -1):
                            if p_row[k]["baseShift"] == "L":
                                prev_crossed = True
                                continue
                            if parse_shift_window(p_row[k]["baseShift"]):
                                prev_s = p_row[k]["baseShift"]
                                prev_off = k - 7
                                break
                            break
                    if not is_transition_valid(prev_s, prev_off, s_code, d, prev_crossed, rules):
                        continue

                    next_s, next_off, next_crossed = None, d + 1, False
                    for k in range(d + 1, 7):
                        if row[k]["baseShift"] == "L":
                            next_crossed = True
                            continue
                        if parse_shift_window(row[k]["baseShift"]):
                            next_s = row[k]["baseShift"]
                            next_off = k
                            break
                        break
                    if next_s and not is_transition_valid(s_code, d, next_s, next_off, next_crossed, rules):
                        continue
                    flex_candidates.append(e)

                def flex_sort_key(e: dict):
                    s_cur = next_assign[e["id"]][d]["baseShift"]
                    cur_cnt = sum(
                        1 for emp in employees
                        if emp["activo"] and emp["cargo"] == e["cargo"]
                        and next_assign[emp["id"]][d]["baseShift"] == s_cur
                    )
                    surplus = cur_cnt - get_shift_target_for_day(role_configs.get(e["cargo"], {}), s_cur, d)
                    return (
                        0 if surplus > 0 else 1,
                        at_totals.get(e["id"], 0),
                        0 if e["cargo"] == "Auxiliar de Operaciones" else 1,
                    )

                flex_candidates.sort(key=flex_sort_key)

                if flex_candidates:
                    chosen = flex_candidates[0]
                    old = next_assign[chosen["id"]][d]
                    next_assign[chosen["id"]][d] = parse_raw_cell(
                        format_cell_raw(s_code, True, old["isCAP"])
                    )
                    at_totals[chosen["id"]] = at_totals.get(chosen["id"], 0) + 1
                    assigned_count += 1
                    covered += 1
                    continue
                break

    return {
        "id": week["id"],
        "label": week["label"],
        "start_date": week["start_date"],
        "assignments": next_assign,
    }, assigned_count


def generate_next_week(
    prev_week: dict,
    employees: List[dict],
    role_configs: dict,
    rules: dict,
    target_start_iso: Optional[str] = None,
    preserve_locked_from: Optional[dict] = None,
    scheduled_novedades: Optional[List[dict]] = None,
) -> Tuple[dict, int]:
    new_start = target_start_iso or add_days_iso(prev_week["start_date"], 7)
    target_dates = [add_days_iso(new_start, d) for d in range(7)]
    novedades_list = scheduled_novedades or []
    new_assign: Dict[str, List[dict]] = {}

    off_count_by_cargo: Dict[str, List[int]] = {}
    shift_count_by_cargo: Dict[str, List[Dict[str, int]]] = {}
    for emp in employees:
        if emp["cargo"] not in off_count_by_cargo:
            off_count_by_cargo[emp["cargo"]] = [0] * 7
            shift_count_by_cargo[emp["cargo"]] = [{} for _ in range(7)]

    def get_active_pool(cargo: str) -> List[str]:
        cfg = role_configs.get(cargo, {})
        pool = list(cfg.get("rotation_pool") or ["08:00-17:00"])
        filtered = [
            s for s in pool
            if get_shift_target_for_day(cfg, s, 0) > 0
            or get_shift_target_for_day(cfg, s, 5) > 0
            or get_shift_target_for_day(cfg, s, 6) > 0
        ]
        return filtered if filtered else pool

    def pick_best_shift_for_days(
        cargo: str,
        active_pool: List[str],
        day_indices: List[int],
        prev_shift_ref: str,
        is_valid_fn,
    ) -> str:
        cfg = role_configs.get(cargo, {})
        counts_by_day = shift_count_by_cargo[cargo]
        idx_ref = active_pool.index(prev_shift_ref) if prev_shift_ref in active_pool else -1
        ideal_next = active_pool[(idx_ref + 1) % len(active_pool)] if idx_ref >= 0 else active_pool[0]

        best_shift = ideal_next
        best_score = float("-inf")
        for cand in active_pool:
            if not is_valid_fn(cand):
                continue
            deficit_score = 0
            surplus_penalty = 0
            for d in day_indices:
                tgt = get_shift_target_for_day(cfg, cand, d)
                cur = counts_by_day[d].get(cand, 0)
                if cur < tgt:
                    deficit_score += 450 if (cargo == "Técnico de Operaciones" and is_tech_shift_no_at(cand)) else 200
                elif cur >= tgt:
                    surplus_penalty += (cur - tgt + 1) * 70
            rot_bonus = 35 if cand == ideal_next else 0
            if len(active_pool) > 1 and cand == prev_shift_ref:
                rot_bonus -= 45
            total = deficit_score - surplus_penalty + rot_bonus
            if total > best_score:
                best_score = total
                best_shift = cand
        return best_shift if best_score != float("-inf") else ideal_next

    for emp_idx, emp in enumerate(employees):
        prev_cells = prev_week["assignments"].get(emp["id"], [])
        preserved_row = preserve_locked_from["assignments"].get(emp["id"], []) if preserve_locked_from else []
        emp_novedades = [n for n in novedades_list if n["emp_id"] == emp["id"]]

        existing_cells = []
        for d, d_iso in enumerate(target_dates):
            preserved = preserved_row[d] if d < len(preserved_row) else None
            if preserved and preserved.get("locked"):
                existing_cells.append(dict(preserved))
                continue
            matched_nov = None
            for nov in reversed(emp_novedades):
                if nov["start_iso"] <= d_iso <= nov["end_iso"]:
                    matched_nov = nov
                    break
            if matched_nov:
                existing_cells.append(parse_raw_cell(matched_nov["tipo"], locked=True))
            elif preserved:
                existing_cells.append(dict(preserved))
            else:
                existing_cells.append(parse_raw_cell(""))

        if not emp["activo"]:
            new_assign[emp["id"]] = [
                dict(existing_cells[d]) if existing_cells[d].get("locked") else parse_raw_cell("")
                for d in range(7)
            ]
            continue

        active_pool = get_active_pool(emp["cargo"])
        primary_prev = get_primary_shift(prev_cells, active_pool[0])
        pool_idx = active_pool.index(primary_prev) if primary_prev in active_pool else -1
        next_target_shift = (
            active_pool[(pool_idx + 1) % len(active_pool)]
            if pool_idx >= 0
            else active_pool[emp_idx % len(active_pool)]
        )

        prev_off_idx = next((i for i, c in enumerate(prev_cells) if c.get("baseShift") == "L"), -1)
        worked_tail_prev = 0
        for pd_idx in range(6, -1, -1):
            c = prev_cells[pd_idx] if pd_idx < len(prev_cells) else None
            if not c or not parse_shift_window(c["baseShift"]):
                break
            worked_tail_prev += 1

        locked_off_idx = next(
            (i for i, c in enumerate(existing_cells) if c.get("locked") and c.get("baseShift") == "L"),
            -1,
        )
        all_days_locked = all(c.get("locked") for c in existing_cells)
        locked_absence_days = sum(
            1 for c in existing_cells
            if c.get("locked") and is_non_working_state(c["baseShift"]) and c["baseShift"] != "L"
        )

        ideal_off_day = (prev_off_idx + 1) % 7 if prev_off_idx >= 0 else emp_idx % 7
        cargo_off_counts = off_count_by_cargo[emp["cargo"]]
        best_off_day = locked_off_idx if locked_off_idx >= 0 else (-1 if all_days_locked else ideal_off_day)

        if locked_off_idx < 0 and not all_days_locked:
            best_score = float("inf")
            for cand_day in range(7):
                if existing_cells[cand_day].get("locked"):
                    continue
                worked_until_off = worked_tail_prev
                for k in range(cand_day):
                    if existing_cells[k].get("locked") and is_non_working_state(existing_cells[k]["baseShift"]):
                        worked_until_off = 0
                    else:
                        worked_until_off += 1

                penalty = cargo_off_counts[cand_day] * 12
                if worked_until_off > rules["max_consecutive_work_days"]:
                    penalty += 5000 + (worked_until_off - rules["max_consecutive_work_days"]) * 500
                elif locked_absence_days == 0 and worked_until_off < rules["min_work_days_between_off"]:
                    penalty += 4000 + (rules["min_work_days_between_off"] - worked_until_off) * 500
                if cand_day != ideal_off_day:
                    penalty += 2
                penalty += abs(worked_until_off - 6)
                if penalty < best_score:
                    best_score = penalty
                    best_off_day = cand_day

        if best_off_day >= 0 and not existing_cells[best_off_day].get("locked"):
            cargo_off_counts[best_off_day] += 1
        elif locked_off_idx >= 0:
            cargo_off_counts[locked_off_idx] += 1

        sunday_work_shift = None
        sunday_work_offset = -1
        sunday_crossed_off = False
        for pd_idx in range(6, -1, -1):
            c = prev_cells[pd_idx] if pd_idx < len(prev_cells) else None
            if not c:
                continue
            if c["baseShift"] == "L":
                sunday_crossed_off = True
                continue
            if parse_shift_window(c["baseShift"]):
                sunday_work_shift = c["baseShift"]
                sunday_work_offset = pd_idx - 7
                break
            break

        pre_days = [d for d in range(7) if d != best_off_day and not existing_cells[d].get("locked") and (best_off_day < 0 or d < best_off_day)]
        post_days = [d for d in range(7) if d != best_off_day and not existing_cells[d].get("locked") and best_off_day >= 0 and d > best_off_day]

        first_unlocked_pre = pre_days[0] if pre_days else -1
        pre_touches_absence = (
            first_unlocked_pre > 0
            and any(existing_cells[k].get("locked") and is_non_working_state(existing_cells[k]["baseShift"]) for k in range(first_unlocked_pre))
        )

        sunday_shift_enabled = bool(
            sunday_work_shift
            and any(get_shift_target_for_day(role_configs.get(emp["cargo"], {}), sunday_work_shift, d) > 0 for d in pre_days)
        )
        can_keep_sunday_on_mon = bool(
            sunday_work_shift
            and sunday_shift_enabled
            and (
                pre_touches_absence
                or is_transition_valid(
                    sunday_work_shift,
                    sunday_work_offset,
                    sunday_work_shift,
                    first_unlocked_pre if first_unlocked_pre >= 0 else 0,
                    sunday_crossed_off,
                    rules,
                )
            )
        )

        if (
            rules["rotate_after_day_off"]
            and not sunday_crossed_off
            and not pre_touches_absence
            and sunday_work_shift
            and can_keep_sunday_on_mon
            and best_off_day > 0
            and pre_days
        ):
            pre_libre_shift = sunday_work_shift
        elif pre_days:
            pre_libre_shift = pick_best_shift_for_days(
                emp["cargo"],
                active_pool,
                pre_days,
                sunday_work_shift or primary_prev,
                lambda cand: (
                    pre_touches_absence
                    or not sunday_work_shift
                    or is_transition_valid(
                        sunday_work_shift,
                        sunday_work_offset,
                        cand,
                        first_unlocked_pre if first_unlocked_pre >= 0 else 0,
                        sunday_crossed_off,
                        rules,
                    )
                ),
            )
        else:
            pre_libre_shift = next_target_shift

        post_libre_shift = pre_libre_shift
        if post_days:
            first_post = post_days[0]
            ref_before_l = pre_libre_shift if pre_days else (sunday_work_shift or primary_prev)
            last_pre_day = pre_days[-1] if pre_days else sunday_work_offset
            crossed_before_post = True if (0 <= best_off_day < first_post) else sunday_crossed_off
            post_libre_shift = pick_best_shift_for_days(
                emp["cargo"],
                active_pool,
                post_days,
                ref_before_l,
                lambda cand: (
                    not ref_before_l
                    or pre_touches_absence
                    or is_transition_valid(ref_before_l, last_pre_day, cand, first_post, crossed_before_post, rules)
                ),
            )

        week_row = []
        counts_by_day = shift_count_by_cargo[emp["cargo"]]
        for d in range(7):
            locked_cell = existing_cells[d] if existing_cells[d].get("locked") else None
            if locked_cell:
                week_row.append(dict(locked_cell))
                b = locked_cell.get("baseShift", "")
                if b and parse_shift_window(b):
                    counts_by_day[d][b] = counts_by_day[d].get(b, 0) + 1
            elif d == best_off_day:
                week_row.append(parse_raw_cell("L"))
            elif best_off_day < 0 or d < best_off_day:
                week_row.append(parse_raw_cell(pre_libre_shift))
                counts_by_day[d][pre_libre_shift] = counts_by_day[d].get(pre_libre_shift, 0) + 1
            else:
                week_row.append(parse_raw_cell(post_libre_shift))
                counts_by_day[d][post_libre_shift] = counts_by_day[d].get(post_libre_shift, 0) + 1

        new_assign[emp["id"]] = week_row

    # Paso 2: Reparar metas diarias por cargo (priorizando 19:00 de Técnico directo)
    cargos_presentes = list(dict.fromkeys(e["cargo"] for e in employees if e["activo"]))
    for cargo in cargos_presentes:
        active_pool = get_active_pool(cargo)
        ordered_pool = (
            sorted(active_pool, key=lambda s: 0 if is_tech_shift_no_at(s) else 1)
            if cargo == "Técnico de Operaciones"
            else active_pool
        )
        cfg = role_configs.get(cargo, {})
        cargo_emps = [e for e in employees if e["activo"] and e["cargo"] == cargo]
        counts_by_day = shift_count_by_cargo[cargo]

        for d in range(7):
            for s_need in ordered_pool:
                target_need = get_shift_target_for_day(cfg, s_need, d)
                allow_borrow_no_at = cargo == "Técnico de Operaciones" and is_tech_shift_no_at(s_need)
                while counts_by_day[d].get(s_need, 0) < target_need:
                    donor_shifts = [
                        s_cand for s_cand in counts_by_day[d].keys()
                        if s_cand != s_need and (
                            counts_by_day[d].get(s_cand, 0) > get_shift_target_for_day(cfg, s_cand, d)
                            or (allow_borrow_no_at and counts_by_day[d].get(s_cand, 0) >= 1)
                        )
                    ]
                    donor_shifts.sort(
                        key=lambda s_c: (
                            counts_by_day[d].get(s_c, 0) - get_shift_target_for_day(cfg, s_c, d),
                            counts_by_day[d].get(s_c, 0),
                        ),
                        reverse=True,
                    )
                    if not donor_shifts:
                        break
                    moved = False
                    for extra_s in donor_shifts:
                        for emp in cargo_emps:
                            row = new_assign[emp["id"]]
                            c = row[d]
                            if c["baseShift"] != extra_s or c.get("locked"):
                                continue
                            prev_cells = prev_week["assignments"].get(emp["id"], [])

                            prev_s, prev_off, prev_crossed = None, d - 1, False
                            for k in range(d - 1, -1, -1):
                                if row[k]["baseShift"] == "L":
                                    prev_crossed = True
                                    continue
                                if parse_shift_window(row[k]["baseShift"]):
                                    prev_s, prev_off = row[k]["baseShift"], k
                                    break
                                break
                            if not prev_s and len(prev_cells) == 7:
                                for k in range(6, -1, -1):
                                    if prev_cells[k]["baseShift"] == "L":
                                        prev_crossed = True
                                        continue
                                    if parse_shift_window(prev_cells[k]["baseShift"]):
                                        prev_s, prev_off = prev_cells[k]["baseShift"], k - 7
                                        break
                                    break
                            if not is_transition_valid(prev_s, prev_off, s_need, d, prev_crossed, rules):
                                continue

                            next_s, next_off, next_crossed = None, d + 1, False
                            for k in range(d + 1, 7):
                                if row[k]["baseShift"] == "L":
                                    next_crossed = True
                                    continue
                                if parse_shift_window(row[k]["baseShift"]):
                                    next_s, next_off = row[k]["baseShift"], k
                                    break
                                break
                            if next_s and not is_transition_valid(s_need, d, next_s, next_off, next_crossed, rules):
                                continue

                            new_assign[emp["id"]][d] = parse_raw_cell(s_need)
                            counts_by_day[d][extra_s] = counts_by_day[d].get(extra_s, 1) - 1
                            counts_by_day[d][s_need] = counts_by_day[d].get(s_need, 0) + 1
                            moved = True
                            break
                        if moved:
                            break
                    if not moved:
                        break

    base_week = {
        "id": f"week_{new_start}",
        "label": build_week_label(new_start),
        "start_date": new_start,
        "assignments": new_assign,
    }
    if rules["auto_assign_at"]:
        return auto_assign_tecnico_at(base_week, employees, role_configs, rules, prev_week)
    return base_week, 0


def analyze_rotation_and_rest(
    curr_week: dict,
    prev_week: Optional[dict],
    employees: List[dict],
    role_configs: dict,
    rules: dict,
) -> List[dict]:
    results = []
    for emp in employees:
        prev_cells = prev_week["assignments"].get(emp["id"], [parse_raw_cell("") for _ in range(7)]) if prev_week else [parse_raw_cell("") for _ in range(7)]
        curr_cells = curr_week["assignments"].get(emp["id"], [parse_raw_cell("") for _ in range(7)])
        pool = (role_configs.get(emp["cargo"], {}) or {}).get("rotation_pool", [])

        prev_off_idx = next((i for i, c in enumerate(prev_cells) if c.get("baseShift") == "L"), None)
        curr_off_idx = next((i for i, c in enumerate(curr_cells) if c.get("baseShift") == "L"), None)

        combined = (prev_cells + curr_cells) if prev_week else list(curr_cells)
        offset_curr = 7 if prev_week else 0

        max_streak_14, streak = 0, 0
        for c in combined:
            if parse_shift_window(c.get("baseShift", "")):
                streak += 1
                max_streak_14 = max(max_streak_14, streak)
            else:
                streak = 0

        consec_before_off = None
        if curr_off_idx is not None:
            cnt = 0
            for i in range(offset_curr + curr_off_idx - 1, -1, -1):
                if parse_shift_window(combined[i].get("baseShift", "")):
                    cnt += 1
                else:
                    break
            consec_before_off = cnt

        shift_before_off, before_idx = "", -1
        shift_after_off, after_idx = "", -1
        if curr_off_idx is not None:
            for i in range(offset_curr + curr_off_idx - 1, -1, -1):
                b = combined[i].get("baseShift", "")
                if b == "L":
                    continue
                if parse_shift_window(b):
                    shift_before_off, before_idx = b, i
                    break
                break
            for i in range(offset_curr + curr_off_idx + 1, len(combined)):
                b = combined[i].get("baseShift", "")
                if b == "L":
                    continue
                if parse_shift_window(b):
                    shift_after_off, after_idx = b, i
                    break
                break

        rest_across_off = None
        if shift_before_off and shift_after_off and before_idx >= 0 and after_idx >= 0:
            rest_across_off = compute_transition_rest_hours(
                combined[before_idx], before_idx, combined[after_idx], after_idx
            )

        first_work_curr = next((c["baseShift"] for c in curr_cells if parse_shift_window(c.get("baseShift", ""))), "")
        if not emp["activo"] or not first_work_curr:
            rot_status = "ABSENCE"
            rot_detail = "Inactivo en rotación" if not emp["activo"] else "Sin turnos (novedad)"
        elif len(pool) <= 1:
            rot_status = "FIXED_ROLE"
            rot_detail = f"Turno fijo ({first_work_curr})"
        elif shift_before_off and shift_after_off and shift_before_off != shift_after_off:
            rot_status = "ROTATED_AFTER_L"
            rot_detail = f"Rotó tras L ({DIAS_SEMANA[curr_off_idx]}): {shift_before_off} → {shift_after_off}"
        elif prev_week:
            prev_primary = get_primary_shift(prev_cells, "")
            if prev_primary and prev_primary != first_work_curr:
                rot_status = "ROTATED_ON_MONDAY"
                rot_detail = f"Rotó desde inicio: {prev_primary} → {first_work_curr}"
            else:
                rot_status = "SAME_SHIFT"
                rot_detail = f"Mantuvo {first_work_curr} (igual a sem. anterior)"
        else:
            rot_status = "FIXED_ROLE"
            rot_detail = f"Semana base ({first_work_curr})"

        if rot_status == "ABSENCE":
            off_interval_ok = True
        elif consec_before_off is not None:
            off_interval_ok = rules["min_work_days_between_off"] <= consec_before_off <= rules["max_consecutive_work_days"]
        else:
            off_interval_ok = max_streak_14 <= rules["max_consecutive_work_days"]

        rest_across_ok = True if rest_across_off is None else (rest_across_off >= rules["min_rest_day_off"])

        results.append({
            "emp": emp,
            "prev_cells": prev_cells,
            "curr_cells": curr_cells,
            "prev_off_idx": prev_off_idx,
            "curr_off_idx": curr_off_idx,
            "consec_before_off": consec_before_off,
            "max_streak_14": max_streak_14,
            "rest_across_off": rest_across_off,
            "rot_status": rot_status,
            "rot_detail": rot_detail,
            "off_interval_ok": off_interval_ok,
            "rest_across_ok": rest_across_ok,
        })
    return results


# ==============================================================================
# IMPORTACIÓN Y EXPORTACIÓN EXCEL (.xlsx / Copiar y Pegar)
# ==============================================================================

def _normalize_cargo(raw_cargo: str) -> str:
    clean = str(raw_cargo or "").strip()
    lower = clean.lower()
    for c in CARGOS_ORDEN:
        if c.lower() == lower:
            return c
    if "analista" in lower:
        return "Analista de Operaciones"
    if "alistamiento" in lower:
        return "Auxiliar de Alistamiento"
    if "auxiliar" in lower:
        return "Auxiliar de Operaciones"
    if "técnico" in lower or "tecnico" in lower:
        return "Técnico de Operaciones"
    if "líder" in lower or "lider" in lower:
        return "Operador Líder"
    if "conductor" in lower:
        return "Conductor de Operaciones"
    return "Auxiliar de Operaciones"


def _parse_rows_to_roster(rows: List[List[str]], start_iso: str) -> Tuple[List[dict], dict]:
    header_idx = -1
    for i, r in enumerate(rows[:20]):
        joined = " ".join(str(x or "").upper() for x in r)
        if ("CODIGO" in joined or "CÓDIGO" in joined) and "NOMBRE" in joined:
            header_idx = i
            break
    if header_idx < 0:
        raise ValueError("No se encontró encabezado con columnas CODIGO, NOMBRE, CARGO y los 7 días.")

    headers = [str(x or "").strip().upper() for x in rows[header_idx]]
    idx_cod = next((i for i, h in enumerate(headers) if "COD" in h), -1)
    idx_nom = next((i for i, h in enumerate(headers) if "NOM" in h), -1)
    idx_car = next((i for i, h in enumerate(headers) if "CARG" in h or "PUESTO" in h), -1)

    day_keys = [("LUN", 0), ("MAR", 1), ("MIE", 2), ("MIÉ", 2), ("JUE", 3), ("VIE", 4), ("SAB", 5), ("SÁB", 5), ("DOM", 6)]
    day_cols = [-1] * 7
    for i, h in enumerate(headers):
        for prefix, d_idx in day_keys:
            if h.startswith(prefix) and day_cols[d_idx] < 0:
                day_cols[d_idx] = i
    if any(dc < 0 for dc in day_cols) and idx_car >= 0:
        day_cols = [idx_car + 1 + k for k in range(7)]

    employees = []
    assignments = {}
    for r in rows[header_idx + 1:]:
        cod = str(r[idx_cod] if 0 <= idx_cod < len(r) else "").strip()
        nom = str(r[idx_nom] if 0 <= idx_nom < len(r) else "").strip()
        if not cod or not nom or cod.upper() == "CODIGO":
            continue
        car_raw = str(r[idx_car] if 0 <= idx_car < len(r) else "").strip()
        cargo = _normalize_cargo(car_raw)
        cells = []
        has_at = False
        has_any = False
        for d_col in day_cols:
            raw_val = str(r[d_col] if 0 <= d_col < len(r) else "").strip()
            c = parse_raw_cell(raw_val)
            if c["isAT"]:
                has_at = True
            if c["raw"]:
                has_any = True
            cells.append(c)
        emp_id = f"emp_{cod}"
        employees.append({
            "id": emp_id,
            "codigo": cod,
            "nombre": nom,
            "cargo": cargo,
            "puede_at": cargo in ("Analista de Operaciones", "Auxiliar de Operaciones") and has_at,
            "activo": has_any,
        })
        assignments[emp_id] = cells

    if not employees:
        raise ValueError("No se detectaron filas de colaboradores válidas.")
    return employees, {
        "id": f"week_{start_iso}",
        "label": build_week_label(start_iso),
        "start_date": start_iso,
        "assignments": assignments,
    }


def import_excel_bytes(data: bytes, start_iso: str) -> Tuple[List[dict], dict]:
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    ws = wb.active
    rows = [[("" if cell is None else str(cell)) for cell in row] for row in ws.iter_rows(values_only=True)]
    return _parse_rows_to_roster(rows, start_iso)


def import_pasted_text(text: str, start_iso: str) -> Tuple[List[dict], dict]:
    lines = [ln for ln in re.split(r"\r?\n", text.strip()) if ln.strip()]
    rows = [ln.split("\t") if "\t" in ln else [c.strip() for c in ln.split(",")] for ln in lines]
    return _parse_rows_to_roster(rows, start_iso)


def export_excel_bytes(week: dict, employees: List[dict]) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Horario Semanal"

    thin = Side(style="thin", color="CBD5E1")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    header_row = ["CODIGO", "NOMBRE", "CARGO"] + [
        f"{DIAS_SEMANA[d]} ({format_short_date_es(add_days_iso(week['start_date'], d))})"
        for d in range(7)
    ]
    ws.append(header_row)
    for col_idx in range(1, 11):
        c = ws.cell(row=1, column=col_idx)
        c.font = Font(bold=True, color="FFFFFF", size=10)
        c.fill = PatternFill("solid", fgColor="1E293B")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border

    sorted_emps = sorted(
        employees,
        key=lambda e: (CARGOS_ORDEN.index(e["cargo"]) if e["cargo"] in CARGOS_ORDEN else 99, e["nombre"]),
    )

    for r_offset, emp in enumerate(sorted_emps, start=2):
        cells = week["assignments"].get(emp["id"], [parse_raw_cell("") for _ in range(7)])
        row_vals = [emp["codigo"], emp["nombre"], emp["cargo"]] + [c.get("raw", "") for c in cells]
        ws.append(row_vals)
        for col_idx in range(1, 11):
            c = ws.cell(row=r_offset, column=col_idx)
            c.border = border
            if col_idx <= 3:
                c.font = Font(bold=(col_idx == 2), size=10, color="0F172A")
                c.alignment = Alignment(horizontal="center" if col_idx == 1 else "left", vertical="center")
            else:
                cell_obj = cells[col_idx - 4]
                _, fg_hex, bg_argb = get_cell_colors(cell_obj)
                c.fill = PatternFill("solid", fgColor=bg_argb)
                c.font = Font(bold=bool(cell_obj.get("isAT") or cell_obj.get("baseShift") in ("L", "VACACIONES", "SUSPEND")),
                              size=10, color=fg_hex.lstrip("#"))
                c.alignment = Alignment(horizontal="center", vertical="center")

    widths = [11, 24, 24] + [16] * 7
    for idx, w in enumerate(widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(idx)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ==============================================================================
# ESTADO INICIAL Y HELPERS DE NOVEDADES EN SESSION_STATE
# ==============================================================================

def create_seed_state():
    employees = []
    assignments = {}
    for cod, nom, cargo, p_at, act, shifts in SEED_ROWS:
        emp_id = f"emp_{cod}"
        employees.append({
            "id": emp_id,
            "codigo": cod,
            "nombre": nom,
            "cargo": cargo,
            "puede_at": p_at,
            "activo": act,
        })
        assignments[emp_id] = [parse_raw_cell(s) for s in shifts]

    week0 = {
        "id": "week_2026-10-26",
        "label": "Semana 26 oct — 1 nov",
        "start_date": "2026-10-26",
        "assignments": assignments,
    }
    return employees, [week0]


def init_state():
    if "roster_initialized" not in st.session_state:
        emps, wks = create_seed_state()
        st.session_state.employees = emps
        st.session_state.weeks = wks
        st.session_state.active_week_idx = 0
        st.session_state.role_configs = {
            k: {
                "rotation_pool": list(v["rotation_pool"]),
                "daily_targets": dict(v["daily_targets"]),
                "saturday_targets": dict(v.get("saturday_targets", {})),
                "sunday_targets": dict(v.get("sunday_targets", {})),
            }
            for k, v in DEFAULT_ROLE_CONFIGS.items()
        }
        st.session_state.rules = dict(DEFAULT_RULES)
        st.session_state.scheduled_novedades = []
        st.session_state.rot_mode = "preview"
        st.session_state.notice = None
        st.session_state.roster_initialized = True


init_state()


def re_rotate_active_week():
    idx = st.session_state.active_week_idx
    curr = st.session_state.weeks[idx]
    prev = st.session_state.weeks[idx - 1] if idx > 0 else curr
    regen, at_cnt = generate_next_week(
        prev,
        st.session_state.employees,
        st.session_state.role_configs,
        st.session_state.rules,
        target_start_iso=curr["start_date"],
        preserve_locked_from=curr,
        scheduled_novedades=st.session_state.scheduled_novedades,
    )
    regen["id"] = curr["id"]
    regen["label"] = curr["label"]
    st.session_state.weeks[idx] = regen
    st.session_state.notice = (
        f"Semana '{curr['label']}' re-rotada aplicando reglas, metas y novedades programadas ({at_cnt} turno(s) AT asignados)."
    )


def register_novedad(
    emp_id: str,
    tipo: str,
    start_iso: str,
    end_iso: str,
    nota: str = "",
    save_to_schedule: bool = True,
    apply_to_existing: bool = True,
    rerotate_current: bool = False,
):
    nov = {
        "id": f"nov_{emp_id}_{int(datetime.datetime.now().timestamp() * 1000)}",
        "emp_id": emp_id,
        "tipo": tipo,
        "start_iso": start_iso,
        "end_iso": end_iso,
        "nota": nota.strip(),
    }
    if save_to_schedule:
        st.session_state.scheduled_novedades.append(nov)

    total_days = 0
    if apply_to_existing:
        updated_weeks = []
        for w in st.session_state.weeks:
            w_new, applied = apply_novedad_to_week(w, nov)
            total_days += applied
            updated_weeks.append(w_new)
        st.session_state.weeks = updated_weeks

    if rerotate_current:
        re_rotate_active_week()

    emp = next((e for e in st.session_state.employees if e["id"] == emp_id), None)
    emp_label = f"{emp['nombre']} ({emp['codigo']})" if emp else emp_id
    extra = f" — aplicada en {total_days} día(s) de semanas creadas" if total_days > 0 else ""
    extra2 = " y semana actual re-rotada." if rerotate_current else "."
    st.session_state.notice = (
        f"Novedad {tipo} registrada para {emp_label} ({start_iso} al {end_iso}){extra}{extra2}"
    )


def render_novedad_form(
    key_prefix: str,
    default_week_start_iso: str,
    allow_rerotate: bool = True,
    default_rerotate: bool = False,
):
    emps = sorted(
        st.session_state.employees,
        key=lambda e: (CARGOS_ORDEN.index(e["cargo"]) if e["cargo"] in CARGOS_ORDEN else 99, e["nombre"]),
    )
    if not emps:
        st.info("No hay colaboradores registrados.")
        return

    emp_options = [e["id"] for e in emps]
    emp_labels = {e["id"]: f"{e['codigo']} — {e['nombre']} ({e['cargo']})" for e in emps}

    base_dt = datetime.date.fromisoformat(default_week_start_iso)
    start_key = f"{key_prefix}_start_val"
    end_key = f"{key_prefix}_end_val"
    if start_key not in st.session_state:
        st.session_state[start_key] = base_dt
    if end_key not in st.session_state:
        st.session_state[end_key] = base_dt + datetime.timedelta(days=4)

    p_cols = st.columns(4)
    if p_cols[0].button("Solo 1 día", key=f"{key_prefix}_p1", use_container_width=True):
        st.session_state[end_key] = st.session_state[start_key]
        st.rerun()
    if p_cols[1].button("Lun–Vie (5 días)", key=f"{key_prefix}_p5", use_container_width=True):
        st.session_state[start_key] = base_dt
        st.session_state[end_key] = base_dt + datetime.timedelta(days=4)
        st.rerun()
    if p_cols[2].button("Sáb–Dom (2 días)", key=f"{key_prefix}_p2", use_container_width=True):
        st.session_state[start_key] = base_dt + datetime.timedelta(days=5)
        st.session_state[end_key] = base_dt + datetime.timedelta(days=6)
        st.rerun()
    if p_cols[3].button("Semana completa (7d)", key=f"{key_prefix}_p7", use_container_width=True):
        st.session_state[start_key] = base_dt
        st.session_state[end_key] = base_dt + datetime.timedelta(days=6)
        st.rerun()

    c1, c2, c3, c4 = st.columns([2.2, 1.3, 1.1, 1.1])
    sel_emp = c1.selectbox(
        "Colaborador",
        options=emp_options,
        format_func=lambda x: emp_labels[x],
        key=f"{key_prefix}_emp",
    )
    sel_tipo = c2.selectbox(
        "Tipo de novedad",
        options=ESTADOS_ESPECIALES,
        format_func=lambda x: "L (Día Libre fijo)" if x == "L" else x,
        key=f"{key_prefix}_tipo",
    )
    dt_start = c3.date_input("Desde", value=st.session_state[start_key], key=f"{key_prefix}_dt_start")
    dt_end = c4.date_input("Hasta", value=st.session_state[end_key], key=f"{key_prefix}_dt_end")
    st.session_state[start_key] = dt_start
    st.session_state[end_key] = max(dt_start, dt_end)

    nota = st.text_input(
        "Observación (opcional)",
        placeholder="Ej. Vacaciones aprobadas / Permiso médico",
        key=f"{key_prefix}_nota",
    )

    o1, o2, o3 = st.columns(3)
    save_sched = o1.checkbox(
        "Guardar para futuras semanas y re-rotaciones",
        value=True,
        key=f"{key_prefix}_save",
    )
    apply_exist = o2.checkbox(
        "Aplicar ya si coincide con semanas creadas",
        value=True,
        key=f"{key_prefix}_apply",
    )
    rerotate_now = (
        o3.checkbox(
            "Re-rotar semana actual al guardar",
            value=default_rerotate,
            key=f"{key_prefix}_rerotate",
        )
        if allow_rerotate
        else False
    )

    if st.button("Guardar novedad", type="primary", key=f"{key_prefix}_submit"):
        if dt_end < dt_start:
            st.error("La fecha 'Hasta' no puede ser anterior a 'Desde'.")
        else:
            register_novedad(
                emp_id=sel_emp,
                tipo=sel_tipo,
                start_iso=dt_start.isoformat(),
                end_iso=dt_end.isoformat(),
                nota=nota,
                save_to_schedule=save_sched,
                apply_to_existing=apply_exist,
                rerotate_current=rerotate_now,
            )
            st.rerun()


# ==============================================================================
# RENDER DE ENCABEZADO Y MÉTRICAS
# ==============================================================================

st.title("Planificador de Turnos Operativos")
st.caption(
    "Rotación semanal Noche → Tarde → Mañana · Descansos 12h / 24h · Meta diaria (Lun–Vie, Sáb, Dom) · Novedades programadas · Cobertura Técnico AT"
)

if st.session_state.notice:
    st.success(st.session_state.notice)
    st.session_state.notice = None

weeks = st.session_state.weeks
employees = st.session_state.employees
role_configs = st.session_state.role_configs
rules = st.session_state.rules
scheduled_novedades = st.session_state.scheduled_novedades
safe_idx = min(max(0, st.session_state.active_week_idx), len(weeks) - 1)
st.session_state.active_week_idx = safe_idx
active_week = weeks[safe_idx]
prev_week = weeks[safe_idx - 1] if safe_idx > 0 else None

top_c1, top_c2, top_c3, top_c4 = st.columns([2.4, 0.9, 1.4, 1.3])
with top_c1:
    week_idx_sel = st.selectbox(
        "Semana activa",
        options=list(range(len(weeks))),
        index=safe_idx,
        format_func=lambda i: f"{i + 1}. {weeks[i]['label']}",
        key="top_week_selector",
    )
    if week_idx_sel != safe_idx:
        st.session_state.active_week_idx = week_idx_sel
        st.rerun()

with top_c2:
    st.write("")
    st.write("")
    if len(weeks) > 1 and safe_idx > 0:
        if st.button("Eliminar", key="top_del_week", use_container_width=True):
            st.session_state.weeks.pop(safe_idx)
            st.session_state.active_week_idx = max(0, safe_idx - 1)
            st.rerun()

with top_c3:
    st.write("")
    st.write("")
    if st.button("Generar Siguiente Semana", type="primary", use_container_width=True):
        last_w = weeks[-1]
        new_w, at_cnt = generate_next_week(
            last_w,
            employees,
            role_configs,
            rules,
            scheduled_novedades=scheduled_novedades,
        )
        st.session_state.weeks.append(new_w)
        st.session_state.active_week_idx = len(st.session_state.weeks) - 1
        st.session_state.notice = (
            f"Se generó '{new_w['label']}' con rotación Noche → Tarde → Mañana ({at_cnt} cobertura(s) AT)."
        )
        st.rerun()

with top_c4:
    st.write("")
    st.write("")
    st.download_button(
        label="Descargar Excel (.xlsx)",
        data=export_excel_bytes(active_week, employees),
        file_name=f"horario_{active_week['start_date']}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

issues = validate_schedule(active_week, employees, role_configs, rules, prev_week)
uncovered_tasks = compute_uncovered_tasks(active_week, employees, role_configs)
uncovered_total = sum(t["missing"] for t in uncovered_tasks)

active_emps_cnt = sum(1 for e in employees if e["activo"])
absents_cnt = sum(
    1 for e in employees if e["activo"]
    and any(c.get("baseShift") in ("VACACIONES", "SUSPEND", "INCAPACIDAD")
            for c in active_week["assignments"].get(e["id"], []))
)
at_cnt_total = sum(
    sum(1 for c in active_week["assignments"].get(e["id"], []) if c.get("isAT"))
    for e in employees if e["activo"]
)
rest_alerts_cnt = sum(1 for i in issues if i["type"] != "UNCOVERED_TECH")
tech_gaps_cnt = sum(1 for i in issues if i["type"] == "UNCOVERED_TECH")

m1, m2, m3, m4, m5, m6 = st.columns(6)
m1.metric("Activos", active_emps_cnt)
m2.metric("Novedades", absents_cnt)
m3.metric("Turnos AT", at_cnt_total)
m4.metric("Descansos 12h/24h", "OK" if rest_alerts_cnt == 0 else f"{rest_alerts_cnt} alerta(s)")
m5.metric("Técnicos", "100% cubierto" if tech_gaps_cnt == 0 else f"{tech_gaps_cnt} faltante(s)")
m6.metric("Tareas sin cubrir", uncovered_total)


# ==============================================================================
# PESTAÑAS PRINCIPALES
# ==============================================================================

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "Horario Semanal",
    "Rotación y Descansos",
    "Cobertura y Apoyo AT",
    "Personal y Reglas",
    "Importar / Exportar Excel",
])

# ------------------------------------------------------------------------------
# TAB 1: HORARIO SEMANAL
# ------------------------------------------------------------------------------
with tab1:
    f1, f2, f3, f4, f5 = st.columns([1.8, 1.5, 1.0, 1.1, 1.4])
    q_search = f1.text_input("Buscar código o nombre", placeholder="Ej. 671 o Guardao", key="t1_search")
    cargo_filter = f2.selectbox("Filtrar por cargo", ["TODOS"] + CARGOS_ORDEN, key="t1_cargo")
    only_alerts = f3.checkbox(
        f"Solo alertas ({len(set(i['emp_id'] for i in issues if i['emp_id']))})",
        key="t1_alerts",
    )
    show_prev_ref = f4.checkbox("Ver ref. semana anterior", value=True, key="t1_prev_ref")
    with f5:
        st.write("")
        if st.button("Auto-cubrir Técnicos (AT)", key="t1_auto_at", use_container_width=True):
            updated_w, cnt = auto_assign_tecnico_at(active_week, employees, role_configs, rules, prev_week)
            st.session_state.weeks[safe_idx] = updated_w
            st.session_state.notice = f"Cobertura AT recalculada: {cnt} turno(s) AT asignado(s)."
            st.rerun()

    emp_issue_map: Dict[Tuple[str, int], str] = {}
    emp_has_issue = set()
    for iss in issues:
        if iss["emp_id"]:
            emp_issue_map[(iss["emp_id"], iss["day_idx"])] = iss["message"]
            emp_has_issue.add(iss["emp_id"])

    rot_by_emp = {}
    if prev_week:
        for item in analyze_rotation_and_rest(active_week, prev_week, employees, role_configs, rules):
            rot_by_emp[item["emp"]["id"]] = item

    sorted_emps = sorted(
        employees,
        key=lambda e: (CARGOS_ORDEN.index(e["cargo"]) if e["cargo"] in CARGOS_ORDEN else 99, e["nombre"]),
    )
    filtered_emps = []
    q_low = q_search.strip().lower()
    for e in sorted_emps:
        if cargo_filter != "TODOS" and e["cargo"] != cargo_filter:
            continue
        if only_alerts and e["id"] not in emp_has_issue:
            continue
        if q_low and q_low not in e["codigo"].lower() and q_low not in e["nombre"].lower():
            continue
        filtered_emps.append(e)

    th_prev = ""
    if prev_week and show_prev_ref:
        th_prev = "<th>SEM. ANTERIOR</th><th>DÍAS P/ L Y ROT.</th>"

    day_headers = "".join(
        f"<th>{DIAS_SEMANA[d][:3]} {format_short_date_es(add_days_iso(active_week['start_date'], d))}</th>"
        for d in range(7)
    )

    rows_html = []
    for e in filtered_emps:
        cells = active_week["assignments"].get(e["id"], [parse_raw_cell("") for _ in range(7)])
        prev_tds = ""
        if prev_week and show_prev_ref:
            an = rot_by_emp.get(e["id"])
            if an:
                prev_dom = (an["prev_cells"][6]["raw"] or "—") if len(an["prev_cells"]) == 7 else "—"
                prev_l = DIAS_SEMANA[an["prev_off_idx"]][:3] if an["prev_off_idx"] is not None else "—"
                d_txt = f"{an['consec_before_off']}d→L" if an["consec_before_off"] is not None else "—"
                prev_tds = (
                    f"<td style='font-size:11px;color:#475569;white-space:nowrap'>Dom: <b>{html.escape(prev_dom)}</b> | L: {prev_l}</td>"
                    f"<td style='font-size:11px;color:#0f172a;white-space:nowrap'><b>{d_txt}</b> · {html.escape(an['rot_detail'].split(':')[0])}</td>"
                )
            else:
                prev_tds = "<td>—</td><td>—</td>"

        day_tds = []
        for d in range(7):
            c = cells[d] if d < len(cells) else parse_raw_cell("")
            bg, fg, _ = get_cell_colors(c)
            alert_msg = emp_issue_map.get((e["id"], d))
            border_css = "2px solid #DC2626" if alert_msg else "1px solid #E2E8F0"
            lock_mark = " 🔒" if c.get("locked") else ""
            txt = html.escape(c.get("raw") or "—")
            title_attr = html.escape(alert_msg or c.get("raw") or "Sin turno")
            day_tds.append(
                f"<td title='{title_attr}' style='background:{bg};color:{fg};border:{border_css};"
                f"text-align:center;font-weight:600;font-size:12px;white-space:nowrap'>{txt}{lock_mark}</td>"
            )

        at_tag = " <span style='color:#BE185D;font-size:10px;font-weight:700'>AT</span>" if e.get("puede_at") else ""
        opacity_css = "opacity:0.55;" if not e["activo"] else ""
        rows_html.append(
            f"<tr style='{opacity_css}'>"
            f"<td style='font-family:monospace'>{html.escape(e['codigo'])}</td>"
            f"<td style='font-weight:600;white-space:nowrap'>{html.escape(e['nombre'])}{at_tag}</td>"
            f"<td style='white-space:nowrap'>{html.escape(e['cargo'])}</td>"
            f"{prev_tds}{''.join(day_tds)}"
            f"</tr>"
        )

    table_html = f"""
    <div style="overflow-x:auto;border:1px solid #CBD5E1;border-radius:8px;background:#FFFFFF;">
      <table style="width:100%;border-collapse:collapse;font-family:sans-serif;font-size:12px;">
        <thead>
          <tr style="background:#F1F5F9;color:#0F172A;border-bottom:2px solid #CBD5E1;text-align:left;">
            <th style="padding:7px 8px;">CODIGO</th>
            <th style="padding:7px 8px;">NOMBRE</th>
            <th style="padding:7px 8px;">CARGO</th>
            {th_prev}
            {day_headers}
          </tr>
        </thead>
        <tbody>
          {''.join(rows_html)}
        </tbody>
      </table>
    </div>
    <style>
      table td, table th {{ padding: 6px 8px; border-bottom: 1px solid #E2E8F0; }}
    </style>
    """
    st.markdown(table_html, unsafe_allow_html=True)

    bc1, bc2 = st.columns([3, 1.3])
    bc1.caption(f"Mostrando {len(filtered_emps)} de {len(employees)} colaboradores.")
    with bc2:
        if st.button("Re-rotar semana actual", key="t1_rerotate", use_container_width=True):
            re_rotate_active_week()
            st.rerun()

    with st.expander("➕ Agregar novedad por rango (Vacaciones, Suspensión, Libre, Incapacidad, Permiso)", expanded=False):
        st.caption("Usa esta opción cuando necesites registrar una novedad de uno o varios días en la semana actual (o futuras semanas) y reajustar al equipo.")
        render_novedad_form(
            key_prefix="t1_nov",
            default_week_start_iso=active_week["start_date"],
            allow_rerotate=True,
            default_rerotate=False,
        )

    with st.expander("✏️ Editar turno o novedad de un colaborador (uno o varios días)", expanded=False):
        emp_edit_id = st.selectbox(
            "Colaborador",
            options=[e["id"] for e in sorted_emps],
            format_func=lambda eid: next(f"{x['codigo']} — {x['nombre']} ({x['cargo']})" for x in sorted_emps if x["id"] == eid),
            key="t1_edit_emp",
        )
        ec1, ec2, ec3 = st.columns([1.2, 1.2, 1.6])
        day_start_edit = ec1.selectbox("Desde el día", options=list(range(7)), format_func=lambda d: DIAS_SEMANA[d], key="t1_edit_day_start")
        day_end_edit = ec2.selectbox(
            "Hasta el día",
            options=[d for d in range(day_start_edit, 7)],
            format_func=lambda d: f"Solo {DIAS_SEMANA[d]}" if d == day_start_edit else f"Hasta {DIAS_SEMANA[d]} ({d - day_start_edit + 1} días)",
            key="t1_edit_day_end",
        )
        curr_cell = (active_week["assignments"].get(emp_edit_id) or [parse_raw_cell("") for _ in range(7)])[day_start_edit]
        all_opts = ["(Vacío)"] + TURNOS_CATALOGO + ESTADOS_ESPECIALES
        cur_b = curr_cell["baseShift"] if curr_cell["baseShift"] in all_opts else "(Vacío)"
        new_base = ec3.selectbox(
            "Turno o Novedad",
            options=all_opts,
            index=all_opts.index(cur_b),
            format_func=lambda x: "L (Día Libre)" if x == "L" else x,
            key="t1_edit_val",
        )
        is_novedad_sel = new_base in ESTADOS_ESPECIALES
        cc1, cc2, cc3, cc4 = st.columns(4)
        new_at = cc1.checkbox("AT (Cubre Técnico)", value=curr_cell["isAT"], disabled=is_novedad_sel, key="t1_edit_at")
        new_cap = cc2.checkbox("CAP", value=curr_cell["isCAP"], disabled=is_novedad_sel, key="t1_edit_cap")
        new_locked = cc3.checkbox("Fijar al rotar", value=(True if is_novedad_sel else bool(curr_cell.get("locked"))), key="t1_edit_lock")
        save_nov_edit = cc4.checkbox("Guardar en novedades prog.", value=is_novedad_sel, disabled=not is_novedad_sel, key="t1_edit_save_nov")

        if st.button("Guardar cambio", type="primary", key="t1_edit_save"):
            val = "" if new_base == "(Vacío)" else new_base
            cell_obj = parse_raw_cell(format_cell_raw(val, False if is_novedad_sel else new_at, False if is_novedad_sel else new_cap), locked=new_locked)
            row = st.session_state.weeks[safe_idx]["assignments"].setdefault(emp_edit_id, [parse_raw_cell("") for _ in range(7)])
            for d_i in range(day_start_edit, day_end_edit + 1):
                row[d_i] = dict(cell_obj)
            if is_novedad_sel and save_nov_edit:
                st.session_state.scheduled_novedades.append({
                    "id": f"nov_{emp_edit_id}_{int(datetime.datetime.now().timestamp() * 1000)}",
                    "emp_id": emp_edit_id,
                    "tipo": val,
                    "start_iso": add_days_iso(active_week["start_date"], day_start_edit),
                    "end_iso": add_days_iso(active_week["start_date"], day_end_edit),
                    "nota": "Agregada desde editor de horario",
                })
            st.session_state.notice = "Cambio aplicado correctamente."
            st.rerun()

    if issues:
        st.markdown("#### Alertas activas de esta semana")
        for iss in issues:
            st.warning(f"**{DIAS_SEMANA[iss['day_idx']]}:** {iss['message']}")


# ------------------------------------------------------------------------------
# TAB 2: ROTACIÓN Y DESCANSOS (ANÁLISIS + PREVIEW + NOVEDADES PREVIAS)
# ------------------------------------------------------------------------------
with tab2:
    st.subheader("Análisis de Días para Descanso (L) y Rotación Tras Días Libres")
    st.caption("Compara la semana anterior contra la semana nueva, registra novedades antes de generarla y verifica días continuos antes de L.")

    m_col1, m_col2 = st.columns([1.4, 2.6])
    with m_col1:
        mode_choice = st.radio(
            "Modo de visualización",
            options=["compare", "preview"],
            index=0 if st.session_state.rot_mode == "compare" else 1,
            format_func=lambda m: f"Semana Activa ({active_week['label']})" if m == "compare" else "Previsualizar Siguiente Semana",
            horizontal=True,
            key="t2_mode_radio",
        )
        st.session_state.rot_mode = mode_choice

    rc1, rc2, rc3, rc4 = st.columns([1, 1, 1.3, 1.5])
    new_min_work = rc1.number_input("Mín. días trabajados entre L", 3, 6, int(rules["min_work_days_between_off"]), key="t2_min_w")
    new_max_cons = rc2.number_input("Máx. días seguidos sin descanso", 5, 10, int(rules["max_consecutive_work_days"]), key="t2_max_c")
    new_rot_after = rc3.checkbox("Rotar turno después del Día Libre (L)", value=bool(rules["rotate_after_day_off"]), key="t2_rot_l")

    rules["min_work_days_between_off"] = int(new_min_work)
    rules["max_consecutive_work_days"] = int(new_max_cons)
    rules["rotate_after_day_off"] = bool(new_rot_after)

    last_week = weeks[-1]
    preview_week, preview_at = generate_next_week(
        last_week,
        employees,
        role_configs,
        rules,
        scheduled_novedades=scheduled_novedades,
    )

    with rc4:
        st.write("")
        if mode_choice == "preview":
            if st.button(f"Aplicar y Crear {preview_week['label']}", type="primary", key="t2_apply_preview", use_container_width=True):
                st.session_state.weeks.append(preview_week)
                st.session_state.active_week_idx = len(st.session_state.weeks) - 1
                st.session_state.notice = f"Se aplicó y creó '{preview_week['label']}' ({preview_at} cobertura(s) AT)."
                st.rerun()
        else:
            if st.button("Simular Próxima Semana", key="t2_go_preview", use_container_width=True):
                st.session_state.rot_mode = "preview"
                st.rerun()

    # Novedades programadas antes de generar la siguiente semana
    with st.expander(f"📅 Novedades programadas ({len(scheduled_novedades)}) — Vacaciones, Suspensión, Libre antes de generar", expanded=False):
        if scheduled_novedades:
            for nov in sorted(scheduled_novedades, key=lambda x: x["start_iso"]):
                emp = next((e for e in employees if e["id"] == nov["emp_id"]), None)
                nc1, nc2, nc3, nc4 = st.columns([2.2, 1.1, 1.8, 0.7])
                nc1.write(f"**{emp['codigo']} — {emp['nombre']}**" if emp else nov["emp_id"])
                nc2.write(f"`{nov['tipo']}`")
                nc3.write(f"{nov['start_iso']} → {nov['end_iso']}")
                if nc4.button("Quitar", key=f"t2_del_nov_{nov['id']}"):
                    st.session_state.scheduled_novedades = [x for x in scheduled_novedades if x["id"] != nov["id"]]
                    st.session_state.notice = "Novedad programada eliminada."
                    st.rerun()
        else:
            st.caption("No hay novedades programadas aún.")
        st.markdown("---")
        render_novedad_form(
            key_prefix="t2_nov",
            default_week_start_iso=preview_week["start_date"] if mode_choice == "preview" else add_days_iso(last_week["start_date"], 7),
            allow_rerotate=False,
            default_rerotate=False,
        )

    eval_curr = preview_week if mode_choice == "preview" else active_week
    eval_prev = last_week if mode_choice == "preview" else prev_week

    analyses = analyze_rotation_and_rest(eval_curr, eval_prev, employees, role_configs, rules)
    eval_issues = validate_schedule(eval_curr, employees, role_configs, rules, eval_prev)

    active_an = [a for a in analyses if a["rot_status"] != "ABSENCE"]
    rot_cnt = sum(1 for a in active_an if a["rot_status"] in ("ROTATED_AFTER_L", "ROTATED_ON_MONDAY"))
    int_bad = sum(1 for a in active_an if not a["off_interval_ok"])
    days_list = [a["consec_before_off"] for a in active_an if a["consec_before_off"] is not None]
    avg_days = round(sum(days_list) / len(days_list), 1) if days_list else 0.0
    rest_conflicts = sum(1 for i in eval_issues if i["type"] != "UNCOVERED_TECH")

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Ventana comparada", f"{eval_prev['label'] if eval_prev else 'Base'} → {eval_curr['label']}")
    k2.metric("Promedio días antes de L", f"{avg_days} días")
    k3.metric("Rotaciones aplicadas", f"{rot_cnt} / {len(active_an)}")
    k4.metric("Intervalo fuera de rango", int_bad)
    k5.metric("Conflictos 12h / 24h", rest_conflicts)

    tf1, tf2 = st.columns([1.6, 2.0])
    t2_cargo = tf1.selectbox("Filtrar cargo", ["TODOS"] + CARGOS_ORDEN, key="t2_cargo")
    t2_only_warn = tf2.checkbox("Mostrar solo sin rotación o alertas de descanso", key="t2_only_warn")

    filtered_an = [
        a for a in sorted(analyses, key=lambda x: (CARGOS_ORDEN.index(x["emp"]["cargo"]) if x["emp"]["cargo"] in CARGOS_ORDEN else 99, x["emp"]["nombre"]))
        if (t2_cargo == "TODOS" or a["emp"]["cargo"] == t2_cargo)
        and (not t2_only_warn or not a["off_interval_ok"] or not a["rest_across_ok"] or a["rot_status"] == "SAME_SHIFT")
    ]

    def strip_html(cells_list):
        chips = []
        for idx, c in enumerate(cells_list):
            bg, fg, _ = get_cell_colors(c)
            b = c.get("baseShift", "")
            short = "L" if b == "L" else "VAC" if b == "VACACIONES" else "SUSP" if b == "SUSPEND" else (b[:5] if b else "—")
            chips.append(
                f"<div style='display:inline-block;width:36px;padding:2px 1px;margin-right:2px;border-radius:3px;"
                f"border:1px solid #CBD5E1;background:{bg};color:{fg};text-align:center;font-family:monospace;font-size:10px;'>"
                f"<div style='font-size:8px;opacity:0.7'>{DIAS_SEMANA[idx][:1]}</div>{html.escape(short)}</div>"
            )
        return "".join(chips)

    comp_rows = []
    for a in filtered_an:
        e = a["emp"]
        prev_strip = strip_html(a["prev_cells"]) if eval_prev else "<span style='color:#94A3B8'>Sin semana previa</span>"
        curr_strip = strip_html(a["curr_cells"])
        d_txt = f"{a['consec_before_off']} días seg." if a["consec_before_off"] is not None else f"Sin L ({a['max_streak_14']}d)"
        d_color = "#047857" if a["off_interval_ok"] else "#B45309"
        h_txt = f"{a['rest_across_off']}h" if a["rest_across_off"] is not None else "—"
        h_color = "#047857" if a["rest_across_ok"] else "#B91C1C"
        comp_rows.append(
            f"<tr>"
            f"<td><b>{html.escape(e['codigo'])} {html.escape(e['nombre'])}</b><br><span style='color:#64748B;font-size:11px'>{html.escape(e['cargo'])}</span></td>"
            f"<td style='white-space:nowrap'>{prev_strip}</td>"
            f"<td style='white-space:nowrap'>{curr_strip}</td>"
            f"<td style='text-align:center;font-weight:700;color:{d_color}'>{d_txt}</td>"
            f"<td style='text-align:center;font-weight:700;color:{h_color}'>{h_txt}</td>"
            f"<td>{html.escape(a['rot_detail'])}</td>"
            f"</tr>"
        )

    st.markdown(
        f"""
        <div style="overflow-x:auto;border:1px solid #CBD5E1;border-radius:8px;background:#FFFFFF;">
          <table style="width:100%;border-collapse:collapse;font-family:sans-serif;font-size:12px;">
            <thead>
              <tr style="background:#F1F5F9;color:#0F172A;border-bottom:2px solid #CBD5E1;text-align:left;">
                <th style="padding:8px;">Colaborador y Cargo</th>
                <th style="padding:8px;">Semana Anterior (Lun → Dom)</th>
                <th style="padding:8px;">{'Nueva Semana Propuesta' if mode_choice == 'preview' else 'Semana Actual'} (Lun → Dom)</th>
                <th style="padding:8px;text-align:center;">Días p/ Descanso</th>
                <th style="padding:8px;text-align:center;">Horas en L</th>
                <th style="padding:8px;">Rotación tras Día Libre</th>
              </tr>
            </thead>
            <tbody>{''.join(comp_rows)}</tbody>
          </table>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ------------------------------------------------------------------------------
# TAB 3: COBERTURA, TAREAS SIN CUBRIR Y APOYO AT
# ------------------------------------------------------------------------------
with tab3:
    st.subheader("Tareas y Turnos Sin Cubrir (Faltantes de Meta)")
    st.caption("Lista exacta de los turnos que quedaron por debajo de la meta diaria configurada (Lun–Vie, Sábado o Domingo) en cualquier cargo.")

    uf_cargo = st.selectbox(
        "Filtrar faltantes por cargo",
        options=["TODOS"] + CARGOS_ORDEN,
        key="t3_uncovered_cargo",
    )
    filtered_unc = [t for t in uncovered_tasks if uf_cargo == "TODOS" or t["cargo"] == uf_cargo]
    if not filtered_unc:
        st.success("No hay tareas pendientes: todas las metas diarias configuradas están 100% cubiertas.")
    else:
        st.dataframe(
            pd.DataFrame([
                {
                    "Día": t["day_name"],
                    "Cargo": t["cargo"],
                    "Turno": t["shift"],
                    "Asignados / Meta": f"{t['assigned']} / {t['target']}",
                    "Faltantes": -t["missing"],
                    "Motivo / Observación": t["reason"],
                }
                for t in filtered_unc
            ]),
            use_container_width=True,
            hide_index=True,
        )

    st.markdown("---")
    c_head, c_btn = st.columns([3, 1.2])
    c_head.subheader("Cobertura Técnico de Operaciones (Directos + Apoyo AT)")
    c_head.caption("El turno 19:00-27:00 es exclusivo de Técnico de Operaciones directo (nunca se cubre con AT). Los turnos 11:00 y 03:00 sí admiten apoyo AT.")
    if c_btn.button("Recalcular Apoyos AT", key="t3_recalc_at", use_container_width=True):
        updated_w, cnt = auto_assign_tecnico_at(active_week, employees, role_configs, rules, prev_week)
        st.session_state.weeks[safe_idx] = updated_w
        st.session_state.notice = f"Apoyos AT recalculados: {cnt} asignación(es)."
        st.rerun()

    tech_cfg = role_configs.get("Técnico de Operaciones", {})
    tech_shifts = list(dict.fromkeys(
        list(tech_cfg.get("rotation_pool", [])) + list((tech_cfg.get("daily_targets") or {}).keys())
    ))
    tech_rows = []
    for s_code in tech_shifts:
        no_at = is_tech_shift_no_at(s_code)
        t_lv = get_shift_target_for_day(tech_cfg, s_code, 0)
        t_s = get_shift_target_for_day(tech_cfg, s_code, 5)
        t_d = get_shift_target_for_day(tech_cfg, s_code, 6)
        row_dict = {"Turno Técnico": f"{s_code} ({'Solo Técnico' if no_at else 'Directo/AT'}) | Meta L-V:{t_lv} S:{t_s} D:{t_d}"}
        for d in range(7):
            tgt = get_shift_target_for_day(tech_cfg, s_code, d)
            names = []
            for emp in employees:
                if not emp["activo"]:
                    continue
                c = (active_week["assignments"].get(emp["id"]) or [])[d]
                if emp["cargo"] == "Técnico de Operaciones" and covers_shift_window(c["baseShift"], s_code):
                    names.append(emp["nombre"])
                elif c.get("isAT") and not no_at and covers_shift_window(c["baseShift"], s_code):
                    names.append(f"AT: {emp['nombre']}")
            row_dict[DIAS_SEMANA[d][:3]] = f"{len(names)}/{tgt}: " + (", ".join(names) if names else "Sin personal")
        tech_rows.append(row_dict)
    st.dataframe(pd.DataFrame(tech_rows), use_container_width=True, hide_index=True)

    st.subheader("Dotación Diaria por Cargo y Turno (Meta L-V / S / D)")
    dot_rows = []
    for cargo in CARGOS_ORDEN:
        cfg = role_configs.get(cargo, {})
        shifts = list(dict.fromkeys(list(cfg.get("rotation_pool", [])) + list((cfg.get("daily_targets") or {}).keys())))
        for s_code in shifts:
            no_at = cargo == "Técnico de Operaciones" and is_tech_shift_no_at(s_code)
            r_dict = {
                "Cargo": cargo,
                "Turno": s_code,
                "Meta (L-V / S / D)": f"{get_shift_target_for_day(cfg, s_code, 0)} / {get_shift_target_for_day(cfg, s_code, 5)} / {get_shift_target_for_day(cfg, s_code, 6)}",
            }
            for d in range(7):
                tgt = get_shift_target_for_day(cfg, s_code, d)
                cnt = 0
                for emp in employees:
                    if not emp["activo"]:
                        continue
                    c = (active_week["assignments"].get(emp["id"]) or [])[d]
                    if emp["cargo"] == cargo and covers_shift_window(c["baseShift"], s_code):
                        cnt += 1
                    elif cargo == "Técnico de Operaciones" and not no_at and c.get("isAT") and covers_shift_window(c["baseShift"], s_code):
                        cnt += 1
                r_dict[DIAS_SEMANA[d][:3]] = f"{cnt}/{tgt}"
            dot_rows.append(r_dict)
    st.dataframe(pd.DataFrame(dot_rows), use_container_width=True, hide_index=True)

    st.subheader("Resumen Semanal de Horas, Días Libres y Turnos AT")
    hrs_rows = []
    for emp in sorted(employees, key=lambda e: (CARGOS_ORDEN.index(e["cargo"]) if e["cargo"] in CARGOS_ORDEN else 99, e["nombre"])):
        if not emp["activo"]:
            continue
        cells = active_week["assignments"].get(emp["id"], [])
        total_min, work_d, free_d, at_d = 0, 0, 0, 0
        for c in cells:
            if c.get("baseShift") == "L":
                free_d += 1
            else:
                w = parse_shift_window(c.get("baseShift", ""))
                if w:
                    total_min += w[2]
                    work_d += 1
                    if c.get("isAT"):
                        at_d += 1
        hrs_rows.append({
            "Código": emp["codigo"],
            "Colaborador": emp["nombre"],
            "Cargo": emp["cargo"],
            "Días Trab.": work_d,
            "Días L": free_d,
            "Apoyos AT": at_d,
            "Horas Semanales": round(total_min / 60.0, 1),
        })
    st.dataframe(pd.DataFrame(hrs_rows), use_container_width=True, hide_index=True)


# ------------------------------------------------------------------------------
# TAB 4: PERSONAL, REGLAS, META DIARIA Y NOVEDADES PROGRAMADAS
# ------------------------------------------------------------------------------
with tab4:
    st.subheader("Programación de Novedades (Vacaciones, Suspensión, Libre, Incapacidad, Permiso)")
    st.caption("Registra ausencias o días libres fijos por fecha antes de generar nuevas semanas o para aplicarlos de inmediato.")
    if scheduled_novedades:
        for nov in sorted(scheduled_novedades, key=lambda x: x["start_iso"]):
            emp = next((e for e in employees if e["id"] == nov["emp_id"]), None)
            nc1, nc2, nc3, nc4, nc5 = st.columns([2.2, 1.1, 1.8, 1.8, 0.7])
            nc1.write(f"**{emp['codigo']} — {emp['nombre']}**" if emp else nov["emp_id"])
            nc2.write(f"`{nov['tipo']}`")
            nc3.write(f"{nov['start_iso']} → {nov['end_iso']}")
            nc4.write(nov.get("nota") or "—")
            if nc5.button("Quitar", key=f"t4_del_nov_{nov['id']}"):
                st.session_state.scheduled_novedades = [x for x in scheduled_novedades if x["id"] != nov["id"]]
                st.session_state.notice = "Novedad programada eliminada."
                st.rerun()
    else:
        st.info("No hay novedades programadas actualmente.")

    with st.expander("➕ Programar nueva novedad", expanded=False):
        render_novedad_form(
            key_prefix="t4_nov",
            default_week_start_iso=active_week["start_date"],
            allow_rerotate=True,
            default_rerotate=False,
        )

    st.markdown("---")
    st.subheader("Reglas de descanso y rotación")
    r1, r2, r3, r4 = st.columns(4)
    rules["min_rest_between_shifts"] = int(r1.number_input("Mín. horas entre turnos", 8, 24, int(rules["min_rest_between_shifts"]), key="t4_r12"))
    rules["min_rest_day_off"] = int(r2.number_input("Mín. horas en Día Libre L", 12, 48, int(rules["min_rest_day_off"]), key="t4_r24"))
    rules["min_work_days_between_off"] = int(r3.number_input("Mín. días trabajados entre L", 3, 6, int(rules["min_work_days_between_off"]), key="t4_min_w"))
    rules["max_consecutive_work_days"] = int(r4.number_input("Máx. días seguidos sin L", 5, 10, int(rules["max_consecutive_work_days"]), key="t4_max_c"))
    rules["rotate_after_day_off"] = st.checkbox("Rotar turno después del Día Libre (L)", value=bool(rules["rotate_after_day_off"]), key="t4_rot_l")
    rules["auto_assign_at"] = st.checkbox("Auto-asignar cobertura AT al generar semana", value=bool(rules["auto_assign_at"]), key="t4_auto_at")

    st.markdown("---")
    st.subheader("Secuencia Noche → Tarde → Mañana y Meta Diaria (Lun–Vie, Sábado, Domingo) por Cargo")
    col_left, col_right = st.columns(2)
    for idx_c, cargo in enumerate(CARGOS_ORDEN):
        cfg = role_configs[cargo]
        container = col_left if idx_c % 2 == 0 else col_right
        with container:
            st.markdown(f"**{cargo}**")
            seq_txt = st.text_input(
                f"Secuencia {cargo}",
                value=", ".join(cfg["rotation_pool"]),
                key=f"t4_seq_{cargo}",
            )
            new_pool = [s.strip() for s in seq_txt.split(",") if s.strip()]
            if not new_pool:
                new_pool = list(cfg["rotation_pool"])
            cfg["rotation_pool"] = new_pool

            new_wk, new_sat, new_sun = {}, {}, {}
            h_cols = st.columns([1.4, 1, 1, 1])
            h_cols[0].caption("Turno")
            h_cols[1].caption("Lun–Vie")
            h_cols[2].caption("Sábado")
            h_cols[3].caption("Domingo")
            for s_code in new_pool:
                wk_def = int(cfg["daily_targets"].get(s_code, 1))
                sat_def = int((cfg.get("saturday_targets") or {}).get(s_code, wk_def))
                sun_def = int((cfg.get("sunday_targets") or {}).get(s_code, wk_def))
                t_cols = st.columns([1.4, 1, 1, 1])
                t_cols[0].write(f"`{s_code}`")
                new_wk[s_code] = int(t_cols[1].number_input("LV", 0, 20, wk_def, label_visibility="collapsed", key=f"tgt_wk_{cargo}_{s_code}"))
                new_sat[s_code] = int(t_cols[2].number_input("Sab", 0, 20, sat_def, label_visibility="collapsed", key=f"tgt_sat_{cargo}_{s_code}"))
                new_sun[s_code] = int(t_cols[3].number_input("Dom", 0, 20, sun_def, label_visibility="collapsed", key=f"tgt_sun_{cargo}_{s_code}"))
            cfg["daily_targets"] = new_wk
            cfg["saturday_targets"] = new_sat
            cfg["sunday_targets"] = new_sun

    if st.button("Aplicar reglas y Re-rotar semana actual", type="primary", key="t4_apply_rerotate"):
        re_rotate_active_week()
        st.rerun()

    st.markdown("---")
    st.subheader("Agregar personal nuevo para cualquier cargo")
    ac1, ac2, ac3, ac4, ac5 = st.columns([1, 2, 1.8, 1, 1])
    new_cod = ac1.text_input("Código", placeholder="Ej. 14020", key="t4_new_cod")
    new_nom = ac2.text_input("Nombre completo", placeholder="Nombre y apellido", key="t4_new_nom")
    new_car = ac3.selectbox("Cargo", CARGOS_ORDEN, key="t4_new_car")
    ac4.write("")
    new_at_flag = ac4.checkbox("Habilitado AT", key="t4_new_at")
    ac5.write("")
    if ac5.button("Agregar", key="t4_add_emp", use_container_width=True):
        if new_cod.strip() and new_nom.strip():
            eid = f"emp_{new_cod.strip()}_{int(datetime.datetime.now().timestamp())}"
            st.session_state.employees.append({
                "id": eid,
                "codigo": new_cod.strip(),
                "nombre": new_nom.strip(),
                "cargo": new_car,
                "puede_at": bool(new_at_flag and new_car in ("Analista de Operaciones", "Auxiliar de Operaciones")),
                "activo": True,
            })
            for w in st.session_state.weeks:
                w["assignments"][eid] = [parse_raw_cell("") for _ in range(7)]
            st.session_state.notice = f"Colaborador {new_nom.strip()} agregado."
            st.rerun()

    st.subheader("Directorio de personal")
    for emp in employees:
        dc1, dc2, dc3, dc4, dc5, dc6 = st.columns([0.8, 2.0, 2.0, 1.0, 1.0, 0.7])
        dc1.write(f"`{emp['codigo']}`")
        dc2.write(emp["nombre"])
        emp["cargo"] = dc3.selectbox(
            "Cargo",
            CARGOS_ORDEN,
            index=CARGOS_ORDEN.index(emp["cargo"]) if emp["cargo"] in CARGOS_ORDEN else 0,
            label_visibility="collapsed",
            key=f"dir_c_{emp['id']}",
        )
        elegible_at = emp["cargo"] in ("Analista de Operaciones", "Auxiliar de Operaciones")
        emp["puede_at"] = dc4.checkbox("AT", value=bool(emp.get("puede_at")) and elegible_at, disabled=not elegible_at, key=f"dir_at_{emp['id']}")
        emp["activo"] = dc5.checkbox("Activo", value=bool(emp["activo"]), key=f"dir_act_{emp['id']}")
        if dc6.button("🗑️", key=f"dir_del_{emp['id']}"):
            st.session_state.employees = [x for x in employees if x["id"] != emp["id"]]
            st.rerun()


# ------------------------------------------------------------------------------
# TAB 5: IMPORTAR / EXPORTAR EXCEL
# ------------------------------------------------------------------------------
with tab5:
    st.subheader("Importar Horario Base desde Excel o Pegar Tabla")
    imp_date = st.date_input(
        "Lunes de inicio de la semana importada",
        value=datetime.date.fromisoformat(active_week["start_date"]),
        key="t5_start_date",
    )

    ic1, ic2 = st.columns(2)
    with ic1:
        st.markdown("**Opción 1: Subir archivo Excel (.xlsx / .xlsm / .xls)**")
        up_file = st.file_uploader("Sube un archivo Excel con columnas CODIGO, NOMBRE, CARGO y los 7 días", type=["xlsx", "xlsm", "xls"], key="t5_uploader")
        if up_file is not None:
            if st.button("Cargar archivo Excel", type="primary", key="t5_load_file"):
                try:
                    new_emps, new_wk = import_excel_bytes(up_file.read(), imp_date.isoformat())
                    st.session_state.employees = new_emps
                    st.session_state.weeks = [new_wk]
                    st.session_state.active_week_idx = 0
                    st.session_state.notice = f"Se importaron {len(new_emps)} colaboradores desde '{up_file.name}'."
                    st.rerun()
                except Exception as ex:
                    st.error(f"Error al importar Excel: {ex}")

    with ic2:
        st.markdown("**Opción 2: Copiar y pegar desde Excel (Ctrl + C / Ctrl + V)**")
        pasted = st.text_area(
            "Selecciona en Excel desde CODIGO hasta DOMINGO, copia y pega aquí:",
            height=170,
            key="t5_paste_area",
        )
        if st.button("Cargar tabla pegada", type="primary", key="t5_load_paste"):
            if not pasted.strip():
                st.error("Pega primero las filas copiadas desde Excel.")
            else:
                try:
                    new_emps, new_wk = import_pasted_text(pasted, imp_date.isoformat())
                    st.session_state.employees = new_emps
                    st.session_state.weeks = [new_wk]
                    st.session_state.active_week_idx = 0
                    st.session_state.notice = f"Se importaron {len(new_emps)} colaboradores desde la tabla pegada."
                    st.rerun()
                except Exception as ex:
                    st.error(f"Error al interpretar la tabla pegada: {ex}")

    st.markdown("---")
    if st.button("Restaurar plantilla inicial", key="t5_reset"):
        emps, wks = create_seed_state()
        st.session_state.employees = emps
        st.session_state.weeks = wks
        st.session_state.active_week_idx = 0
        st.session_state.scheduled_novedades = []
        st.session_state.notice = "Se restauró la plantilla inicial del horario."
        st.rerun()
