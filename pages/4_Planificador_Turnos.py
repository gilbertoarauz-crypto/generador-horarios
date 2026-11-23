import io
import re
from datetime import date, timedelta
import pandas as pd
import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

try:
    st.set_page_config(
        page_title="Planificador de Turnos Operativos",
        layout="wide",
    )
except Exception:
    pass

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

ESTADOS_ESPECIALES = ["L", "VACACIONES", "SUSPEND", "INCAPACIDAD", "PERMISO"]

MESES_ES = [
    "ene", "feb", "mar", "abr", "may", "jun",
    "jul", "ago", "sep", "oct", "nov", "dic",
]

DEFAULT_RULES = {
    "min_rest_shifts": 12,
    "min_rest_day_off": 24,
    "days_off_per_week": 1,
    "min_work_days_between_off": 5,
    "max_consecutive_work_days": 7,
    "rotate_after_day_off": True,
    "auto_assign_at": True,
}

# Secuencia ajustada a Noche -> Tarde -> Mañana
DEFAULT_ROLE_CONFIGS = {
    "Analista de Operaciones": {
        "rotation_pool": ["19:00-27:00", "11:00-19:00", "06:00-15:00", "03:00-11:00"],
        "daily_targets": {"19:00-27:00": 1, "11:00-19:00": 1, "06:00-15:00": 1, "03:00-11:00": 1},
    },
    "Auxiliar de Operaciones": {
        "rotation_pool": ["19:00-27:00", "23:00-31:00", "11:00-19:00", "03:00-11:00"],
        "daily_targets": {"19:00-27:00": 1, "23:00-31:00": 1, "11:00-19:00": 2, "03:00-11:00": 2},
    },
    "Operador Líder": {
        "rotation_pool": ["20:00-28:00"],
        "daily_targets": {"20:00-28:00": 1},
    },
    "Técnico de Operaciones": {
        "rotation_pool": ["19:00-27:00", "11:00-19:00", "03:00-11:00"],
        "daily_targets": {"19:00-27:00": 1, "11:00-19:00": 1, "03:00-11:00": 1},
    },
    "Auxiliar de Alistamiento": {
        "rotation_pool": ["22:00-30:00", "20:00-28:00", "08:00-17:00"],
        "daily_targets": {"22:00-30:00": 1, "20:00-28:00": 2, "08:00-17:00": 1},
    },
    "Conductor de Operaciones": {
        "rotation_pool": ["11:00-18:00", "03:00-10:00"],
        "daily_targets": {"11:00-18:00": 1, "03:00-10:00": 1},
    },
}

SEED_ROWS = [
    ("669", "Roberto Velarde", "Analista de Operaciones", False, True, ["L", "06:00-15:00", "06:00-15:00", "06:00-15:00 CAP", "08:00-17:00", "11:00-19:00", "08:00-17:00"]),
    ("12520", "Keily Agrazal", "Analista de Operaciones", True, True, ["11:00-19:00", "11:00-19:00", "11:00-19:00", "11:00-19:00", "L", "19:00-27:00", "19:00-27:00"]),
    ("13563", "Kathleen Batista", "Analista de Operaciones", True, True, ["VACACIONES", "VACACIONES", "VACACIONES", "VACACIONES", "VACACIONES", "L", "03:00-11:00"]),
    ("3332", "Kevin Pitty", "Analista de Operaciones", False, True, ["19:00-27:00", "19:00-27:00", "19:00-27:00", "19:00-27:00", "19:00-27:00", "L", "11:00-19:00"]),
    ("12896", "Alexis Mojica", "Analista de Operaciones", True, True, ["03:00-11:00 AT", "03:00-11:00 AT", "03:00-11:00 AT", "03:00-11:00", "03:00-11:00 AT", "03:00-11:00", "L"]),
    ("3393", "David Vega", "Auxiliar de Operaciones", True, True, ["L", "SUSPEND", "11:00-19:00", "11:00-19:00", "11:00-27:00", "11:00-19:00", "11:00-19:00 AT"]),
    ("4211", "Dariel Cardales", "Auxiliar de Operaciones", True, True, ["03:00-11:00", "03:00-11:00", "03:00-11:00", "03:00-11:00 AT", "SUSPEND", "03:00-11:00", "L"]),
    ("4607", "Nelson Mojica", "Auxiliar de Operaciones", True, True, ["11:00-19:00", "11:00-19:00", "11:00-19:00", "11:00-19:00 AT/CAP", "11:00-19:00", "11:00-19:00 AT", "L"]),
    ("12717", "Francisco Delgado", "Auxiliar de Operaciones", True, True, ["VACACIONES", "VACACIONES", "VACACIONES", "VACACIONES", "VACACIONES", "03:00-11:00 AT", "03:00-11:00"]),
    ("13580", "Juan Patiño", "Auxiliar de Operaciones", False, True, ["19:00-27:00", "19:00-27:00", "19:00-27:00", "19:00-27:00", "L", "23:00-35:00", "23:00-31:00"]),
    ("13488", "Edwin Aranda", "Auxiliar de Operaciones", True, True, ["11:00-19:00", "11:00-19:00", "L", "03:00-11:00", "03:00-11:00", "19:00-27:00", "19:00-27:00"]),
    ("12644", "Idelfonso Sánchez", "Auxiliar de Operaciones", False, True, ["03:00-11:00", "03:00-11:00", "03:00-11:00", "03:00-11:00", "03:00-11:00", "L", "03:00-11:00"]),
    ("81569", "Kathleen Amat", "Auxiliar de Operaciones", False, True, ["23:00-31:00", "23:00-31:00", "23:00-31:00", "23:00-31:00", "23:00-31:00", "L", "11:00-19:00"]),
    ("759", "Daniel Solis", "Operador Líder", False, True, ["20:00-28:00", "20:00-28:00", "20:00-28:00", "20:00-28:00", "20:00-28:00", "L", "20:00-28:00"]),
    ("10043", "Gilberto Torrero", "Operador Líder", False, True, ["20:00-28:00", "20:00-28:00", "20:00-28:00", "20:00-28:00", "L", "20:00-28:00", "20:00-28:00"]),
    ("671", "Luis Guardao", "Técnico de Operaciones", False, True, ["03:00-11:00", "03:00-11:00", "03:00-11:00 CAP", "03:00-11:00", "L", "19:00-27:00", "19:00-27:00"]),
    ("12560", "Estefany Villamil", "Técnico de Operaciones", False, True, ["19:00-27:00", "19:00-27:00", "19:00-27:00", "19:00-27:00", "19:00-27:00", "L", "11:00-19:00"]),
    ("11383", "Kathia Gonzalez", "Técnico de Operaciones", False, True, ["11:00-19:00", "11:00-19:00", "11:00-19:00", "11:00-19:00", "11:00-19:00", "11:00-19:00", "L"]),
    ("13479", "Edier Barrios", "Técnico de Operaciones", False, True, ["11:00-19:00", "11:00-19:00", "11:00-19:00", "L", "03:00-19:00", "03:00-11:00", "03:00-11:00"]),
    ("590", "Julio Aguilar", "Auxiliar de Alistamiento", False, True, ["20:00-28:00", "20:00-28:00", "20:00-28:00", "20:00-28:00", "20:00-28:00", "L", "20:00-28:00"]),
    ("7144", "Darjil Ortega", "Auxiliar de Alistamiento", False, True, ["22:00-30:00", "22:00-30:00", "22:00-30:00", "L", "20:00-28:00", "20:00-28:00", "20:00-28:00"]),
    ("11584", "Alberto Puchie", "Auxiliar de Alistamiento", False, True, ["08:00-17:00", "08:00-17:00", "L", "22:00-30:00", "22:00-30:00", "20:00-28:00", "20:00-28:00"]),
    ("11589", "Nelson Carrasquilla", "Auxiliar de Alistamiento", False, True, ["VACACIONES", "VACACIONES", "VACACIONES", "VACACIONES", "VACACIONES", "08:00-17:00", "08:00-17:00"]),
    ("12933", "Jose Canto", "Auxiliar de Alistamiento", False, True, ["20:00-28:00", "20:00-28:00", "20:00-28:00", "20:00-28:00", "L", "22:00-30:00", "22:00-30:00"]),
    ("12936", "Eric Batista", "Auxiliar de Alistamiento", False, False, ["", "", "", "", "", "", ""]),
    ("13618", "Isidro Concepcion", "Conductor de Operaciones", False, True, ["11:00-18:00", "11:00-18:00", "11:00-18:00", "11:00-18:00", "11:00-18:00", "11:00-18:00", "L"]),
    ("13635", "Ricardo Palacio", "Conductor de Operaciones", False, True, ["03:00-10:00", "03:00-10:00", "03:00-10:00", "03:00-10:00", "03:00-10:00", "L", "11:00-18:00"]),
]


def parse_raw_cell(raw_text: str, locked: bool = False) -> dict:
    clean = re.sub(r"\s+", " ", str(raw_text or "").strip().upper())
    if not clean:
        return {"raw": "", "base": "", "is_at": False, "is_cap": False, "locked": locked}

    tokens = clean.split(" ")
    first = tokens[0]
    rest = " ".join(tokens[1:])
    is_at = bool(re.search(r"\bAT\b|AT/CAP|CAP/AT", clean))
    is_cap = bool(re.search(r"\bCAP\b|AT/CAP|CAP/AT", clean))

    if first in ("L", "LIBRE", "VACACIONES", "VAC", "SUSPEND", "SUSPENDIDO", "INCAPACIDAD", "PERMISO"):
        norm = "L" if first == "LIBRE" else "VACACIONES" if first == "VAC" else "SUSPEND" if first == "SUSPENDIDO" else first
        return {
            "raw": norm,
            "base": norm,
            "is_at": False,
            "is_cap": False,
            "locked": locked or norm in ("VACACIONES", "SUSPEND"),
        }

    suffix = []
    if is_at and is_cap:
        suffix.append("AT/CAP")
    else:
        if is_at:
            suffix.append("AT")
        if is_cap:
            suffix.append("CAP")

    if suffix:
        final_raw = f"{first} {' '.join(suffix)}"
    elif rest:
        final_raw = clean
    else:
        final_raw = first

    return {"raw": final_raw, "base": first, "is_at": is_at, "is_cap": is_cap, "locked": locked}


def format_cell_raw(base: str, is_at: bool = False, is_cap: bool = False) -> str:
    b = str(base or "").strip().upper()
    if not b:
        return ""
    if b in ESTADOS_ESPECIALES:
        return b
    if is_at and is_cap:
        return f"{b} AT/CAP"
    if is_at:
        return f"{b} AT"
    if is_cap:
        return f"{b} CAP"
    return b


def parse_shift_window(base_shift: str):
    m = re.match(r"^(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})$", str(base_shift or "").strip().upper())
    if not m:
        return None
    sh, sm, eh, em = map(int, m.groups())
    if sm < 0 or sm > 59 or em < 0 or em > 59:
        return None
    start_min = sh * 60 + sm
    end_min = eh * 60 + em
    if end_min <= start_min:
        end_min += 1440
    return start_min, end_min, end_min - start_min


def is_non_working(base_shift: str) -> bool:
    return str(base_shift or "").strip().upper() in ("", "L", "VACACIONES", "SUSPEND", "INCAPACIDAD", "PERMISO")


def covers_shift_window(assigned_shift: str, required_shift: str) -> bool:
    if assigned_shift == required_shift:
        return True
    wa = parse_shift_window(assigned_shift)
    wr = parse_shift_window(required_shift)
    if not wa or not wr:
        return False
    return wa[0] <= wr[0] and wa[1] >= wr[1]


def compute_rest_hours(prev_cell: dict, prev_day_idx: int, next_cell: dict, next_day_idx: int):
    if not prev_cell or not next_cell:
        return None
    wp = parse_shift_window(prev_cell["base"])
    wn = parse_shift_window(next_cell["base"])
    if not wp or not wn:
        return None
    end_abs = prev_day_idx * 1440 + wp[1]
    start_abs = next_day_idx * 1440 + wn[0]
    return round((start_abs - end_abs) / 60.0, 1)


def get_cell_colors(cell: dict) -> tuple[str, str]:
    base = cell["base"].strip().upper()
    if base == "SUSPEND":
        return "DC2626", "FFFFFF"
    if base == "VACACIONES":
        return "FFFFFF", "DC2626"
    if base in ("INCAPACIDAD", "PERMISO"):
        return "FFE4E6", "9F1239"
    if base == "L":
        return "FCE4D6", "C00000"
    if cell["is_at"] or base in ("03:00-19:00", "11:00-27:00", "23:00-35:00"):
        return "F4B8E4", "0F172A"
    if base.startswith("03:00-"):
        return "92D050", "0F172A"
    if base.startswith("11:00-"):
        return "FFC000", "0F172A"
    if base.startswith("19:00-"):
        return "B4C6E7", "0F172A"
    if base.startswith("23:00-"):
        return "FFF2CC", "0F172A"
    if base.startswith("20:00-"):
        return "00B0F0", "0F172A"
    if base.startswith("22:00-"):
        return "A6A6A6", "0F172A"
    if base in ("06:00-15:00", "08:00-17:00"):
        return "FFFFFF", "0F172A"
    return "F8FAFC", "334155"


def add_days_iso(iso_date: str, days: int) -> str:
    dt = date.fromisoformat(iso_date) + timedelta(days=days)
    return dt.isoformat()


def format_short_es(iso_date: str) -> str:
    dt = date.fromisoformat(iso_date)
    return f"{dt.day}-{MESES_ES[dt.month - 1]}"


def build_week_label(start_iso: str) -> str:
    return f"Semana {format_short_es(start_iso)} — {format_short_es(add_days_iso(start_iso, 6))}"


def get_primary_shift(cells: list, fallback: str) -> str:
    counts = {}
    for c in cells:
        if c and parse_shift_window(c["base"]):
            counts[c["base"]] = counts.get(c["base"], 0) + 1
    if not counts:
        return fallback
    return max(counts.items(), key=lambda x: x[1])[0]


def is_transition_valid(last_shift, last_offset, cand_shift, cand_day, crossed_off, rules):
    if not last_shift:
        return True
    wp = parse_shift_window(last_shift)
    wn = parse_shift_window(cand_shift)
    if not wp or not wn:
        return True
    rest_h = (cand_day * 1440 + wn[0] - (last_offset * 1440 + wp[1])) / 60.0
    req = rules["min_rest_day_off"] if crossed_off else rules["min_rest_shifts"]
    return rest_h >= req


def validate_schedule(week: dict, employees: list, role_configs: dict, rules: dict, prev_week=None) -> list:
    issues = []
    for emp in employees:
        if not emp["activo"]:
            continue
        cells = week["assignments"].get(emp["id"], [parse_raw_cell("") for _ in range(7)])
        prev_cells = prev_week["assignments"].get(emp["id"], []) if prev_week else []

        last_work_cell = None
        last_work_offset = -1
        saw_off = False
        streak = 0

        if len(prev_cells) == 7:
            for pd_idx in range(6, -1, -1):
                c = prev_cells[pd_idx]
                if c["base"] == "L":
                    saw_off = True
                    continue
                if parse_shift_window(c["base"]):
                    last_work_cell = c
                    last_work_offset = pd_idx - 7
                    break
                if is_non_working(c["base"]):
                    break
            for pd_idx in range(6, -1, -1):
                if parse_shift_window(prev_cells[pd_idx]["base"]):
                    streak += 1
                else:
                    break

        dias_libres = 0
        dias_trabajados = 0
        streak_alerted = False

        for d in range(7):
            cell = cells[d]
            if cell["base"] == "L":
                dias_libres += 1
                saw_off = True
                streak = 0
                continue
            if is_non_working(cell["base"]):
                last_work_cell = None
                saw_off = False
                streak = 0
                continue

            if parse_shift_window(cell["base"]):
                dias_trabajados += 1
                streak += 1
                if not streak_alerted and streak > rules["max_consecutive_work_days"]:
                    streak_alerted = True
                    issues.append({
                        "employee_id": emp["id"],
                        "day_index": d,
                        "severity": "warning",
                        "type": "CONSECUTIVE_DAYS",
                        "message": f"{emp['nombre']} ({emp['codigo']}): {streak} días continuos sin Día Libre L (máx. {rules['max_consecutive_work_days']}).",
                    })
                if last_work_cell:
                    rh = compute_rest_hours(last_work_cell, last_work_offset, cell, d)
                    if rh is not None:
                        if saw_off and rh < rules["min_rest_day_off"]:
                            issues.append({
                                "employee_id": emp["id"],
                                "day_index": d,
                                "severity": "error",
                                "type": "REST_24H_LIBRE",
                                "message": f"{emp['nombre']} ({emp['codigo']}): solo {rh}h tras Día Libre L (mín. {rules['min_rest_day_off']}h).",
                            })
                        elif not saw_off and rh < rules["min_rest_shifts"]:
                            issues.append({
                                "employee_id": emp["id"],
                                "day_index": d,
                                "severity": "error",
                                "type": "REST_12H",
                                "message": f"{emp['nombre']} ({emp['codigo']}): {rh}h entre {last_work_cell['base']} y {cell['base']} (mín. {rules['min_rest_shifts']}h).",
                            })
                last_work_cell = cell
                last_work_offset = d
                saw_off = False

        if dias_trabajados >= 6 and dias_libres < rules["days_off_per_week"]:
            issues.append({
                "employee_id": emp["id"],
                "day_index": 6,
                "severity": "warning",
                "type": "NO_DAY_OFF",
                "message": f"{emp['nombre']} ({emp['codigo']}) tiene {dias_libres} día(s) libre(s) asignado(s) (requerido: {rules['days_off_per_week']}).",
            })

    tech_targets = role_configs.get("Técnico de Operaciones", {}).get("daily_targets", {})
    for d in range(7):
        for shift_code, target in tech_targets.items():
            if target <= 0:
                continue
            covered = 0
            for emp in employees:
                if not emp["activo"]:
                    continue
                cell = week["assignments"].get(emp["id"], [parse_raw_cell("") for _ in range(7)])[d]
                if not cell["base"]:
                    continue
                if emp["cargo"] == "Técnico de Operaciones" and covers_shift_window(cell["base"], shift_code):
                    covered += 1
                elif cell["is_at"] and covers_shift_window(cell["base"], shift_code):
                    covered += 1
            if covered < target:
                issues.append({
                    "employee_id": None,
                    "day_index": d,
                    "severity": "warning",
                    "type": "UNCOVERED_TECH",
                    "message": f"{DIAS_SEMANA[d]}: Turno {shift_code} de Técnico de Operaciones sin cobertura completa ({covered}/{target}).",
                })
    return issues


def auto_assign_tecnico_at(week: dict, employees: list, role_configs: dict, rules: dict, prev_week=None):
    next_assign = {eid: [dict(c) for c in row] for eid, row in week["assignments"].items()}

    for emp in employees:
        if not emp["activo"] or emp["cargo"] == "Técnico de Operaciones":
            continue
        row = next_assign.get(emp["id"], [])
        for d in range(len(row)):
            if row[d]["is_at"] and not row[d].get("locked"):
                row[d] = parse_raw_cell(format_cell_raw(row[d]["base"], False, row[d]["is_cap"]))

    tech_targets = role_configs.get("Técnico de Operaciones", {}).get("daily_targets", {})
    at_totals = {e["id"]: sum(1 for c in next_assign.get(e["id"], []) if c["is_at"]) for e in employees}
    assigned = 0

    for d in range(7):
        for shift_code, target in tech_targets.items():
            if target <= 0:
                continue
            while True:
                covered = 0
                for emp in employees:
                    if not emp["activo"]:
                        continue
                    c = next_assign.get(emp["id"], [parse_raw_cell("") for _ in range(7)])[d]
                    if (emp["cargo"] == "Técnico de Operaciones" or c["is_at"]) and covers_shift_window(c["base"], shift_code):
                        covered += 1
                if covered >= target:
                    break

                same_cands = [
                    e for e in employees
                    if e["activo"] and e["puede_at"]
                    and e["cargo"] in ("Auxiliar de Operaciones", "Analista de Operaciones")
                    and not next_assign[e["id"]][d]["is_at"]
                    and not next_assign[e["id"]][d].get("locked")
                    and next_assign[e["id"]][d]["base"] == shift_code
                ]
                same_cands.sort(key=lambda e: (at_totals[e["id"]], 0 if e["cargo"] == "Auxiliar de Operaciones" else 1))
                if same_cands:
                    chosen = same_cands[0]
                    old = next_assign[chosen["id"]][d]
                    next_assign[chosen["id"]][d] = parse_raw_cell(format_cell_raw(old["base"], True, old["is_cap"]))
                    at_totals[chosen["id"]] += 1
                    assigned += 1
                    continue

                flex_cands = []
                for e in employees:
                    if not (e["activo"] and e["puede_at"] and e["cargo"] in ("Auxiliar de Operaciones", "Analista de Operaciones")):
                        continue
                    row = next_assign[e["id"]]
                    cell = row[d]
                    if cell["is_at"] or cell.get("locked") or is_non_working(cell["base"]):
                        continue

                    prev_s, prev_off, crossed = None, d - 1, False
                    for k in range(d - 1, -1, -1):
                        if row[k]["base"] == "L":
                            crossed = True
                            continue
                        if parse_shift_window(row[k]["base"]):
                            prev_s, prev_off = row[k]["base"], k
                            break
                        break
                    if not prev_s and d == 0 and prev_week:
                        p_row = prev_week["assignments"].get(e["id"], [])
                        for k in range(len(p_row) - 1, -1, -1):
                            if p_row[k]["base"] == "L":
                                crossed = True
                                continue
                            if parse_shift_window(p_row[k]["base"]):
                                prev_s, prev_off = p_row[k]["base"], k - 7
                                break
                            break
                    if not is_transition_valid(prev_s, prev_off, shift_code, d, crossed, rules):
                        continue

                    next_s, next_off, next_crossed = None, d + 1, False
                    for k in range(d + 1, 7):
                        if row[k]["base"] == "L":
                            next_crossed = True
                            continue
                        if parse_shift_window(row[k]["base"]):
                            next_s, next_off = row[k]["base"], k
                            break
                        break
                    if next_s and not is_transition_valid(shift_code, d, next_s, next_off, next_crossed, rules):
                        continue
                    flex_cands.append(e)

                flex_cands.sort(key=lambda e: at_totals[e["id"]])
                if flex_cands:
                    chosen = flex_cands[0]
                    old = next_assign[chosen["id"]][d]
                    next_assign[chosen["id"]][d] = parse_raw_cell(format_cell_raw(shift_code, True, old["is_cap"]))
                    at_totals[chosen["id"]] += 1
                    assigned += 1
                    continue
                break

    return {"id": week["id"], "label": week["label"], "start_iso": week["start_iso"], "assignments": next_assign}, assigned


def generate_next_week(prev_week: dict, employees: list, role_configs: dict, rules: dict, target_start_iso=None):
    new_start = target_start_iso or add_days_iso(prev_week["start_iso"], 7)
    new_assign = {}
    off_count_cargo = {c: [0] * 7 for c in CARGOS_ORDEN}

    for idx, emp in enumerate(employees):
        prev_cells = prev_week["assignments"].get(emp["id"], [])
        if not emp["activo"]:
            new_assign[emp["id"]] = [parse_raw_cell("") for _ in range(7)]
            continue

        pool = role_configs.get(emp["cargo"], {}).get("rotation_pool") or ["08:00-17:00"]
        primary_prev = get_primary_shift(prev_cells, pool[0])
        pool_idx = pool.index(primary_prev) if primary_prev in pool else -1
        next_target = pool[(pool_idx + 1) % len(pool)] if pool_idx >= 0 else pool[idx % len(pool)]

        prev_off_idx = next((i for i, c in enumerate(prev_cells) if c["base"] == "L"), -1)
        worked_tail = 0
        for pd_idx in range(len(prev_cells) - 1, -1, -1):
            if parse_shift_window(prev_cells[pd_idx]["base"]):
                worked_tail += 1
            else:
                break

        ideal_off = (prev_off_idx + 1) % 7 if prev_off_idx >= 0 else idx % 7
        cargo_counts = off_count_cargo.setdefault(emp["cargo"], [0] * 7)
        best_off = ideal_off
        best_score = 10**9

        for cand in range(7):
            worked_until = worked_tail + cand
            penalty = cargo_counts[cand] * 12
            if worked_until > rules["max_consecutive_work_days"]:
                penalty += 5000 + (worked_until - rules["max_consecutive_work_days"]) * 500
            elif worked_until < rules["min_work_days_between_off"]:
                penalty += 4000 + (rules["min_work_days_between_off"] - worked_until) * 500
            if cand != ideal_off:
                penalty += 2
            penalty += abs(worked_until - 6)
            if penalty < best_score:
                best_score = penalty
                best_off = cand
        cargo_counts[best_off] += 1

        sun_shift, sun_off, sun_crossed = None, -1, False
        for pd_idx in range(len(prev_cells) - 1, -1, -1):
            c = prev_cells[pd_idx]
            if c["base"] == "L":
                sun_crossed = True
                continue
            if parse_shift_window(c["base"]):
                sun_shift, sun_off = c["base"], pd_idx - 7
                break
            break

        can_mon_next = is_transition_valid(sun_shift, sun_off, next_target, 0, sun_crossed, rules)
        can_keep_sun = is_transition_valid(sun_shift, sun_off, sun_shift, 0, sun_crossed, rules) if sun_shift else True

        if rules["rotate_after_day_off"] and not sun_crossed and sun_shift and can_keep_sun and best_off > 0:
            pre_shift = sun_shift
        elif can_mon_next or not sun_shift:
            pre_shift = next_target
        elif can_keep_sun:
            pre_shift = sun_shift
        else:
            pre_shift = next_target

        if pre_shift != sun_shift and not rules["rotate_after_day_off"]:
            desired_post = pre_shift
        elif pool_idx >= 0 and pre_shift == primary_prev:
            desired_post = next_target
        else:
            idx_pre = pool.index(pre_shift) if pre_shift in pool else -1
            desired_post = pool[(idx_pre + 1) % len(pool)] if idx_pre >= 0 else next_target

        post_shift = desired_post
        if best_off > 0:
            if not is_transition_valid(pre_shift, best_off - 1, desired_post, best_off + 1, True, rules):
                comp = next((c for c in pool if is_transition_valid(pre_shift, best_off - 1, c, best_off + 1, True, rules)), None)
                if comp:
                    post_shift = comp
        else:
            if not is_transition_valid(sun_shift, sun_off, desired_post, 1, True, rules):
                comp = next((c for c in pool if is_transition_valid(sun_shift, sun_off, c, 1, True, rules)), None)
                if comp:
                    post_shift = comp

        row = []
        for d in range(7):
            if d == best_off:
                row.append(parse_raw_cell("L"))
            elif d < best_off:
                row.append(parse_raw_cell(pre_shift))
            else:
                row.append(parse_raw_cell(post_shift))
        new_assign[emp["id"]] = row

    base_week = {
        "id": f"week_{new_start}",
        "label": build_week_label(new_start),
        "start_iso": new_start,
        "assignments": new_assign,
    }
    if rules["auto_assign_at"]:
        return auto_assign_tecnico_at(base_week, employees, role_configs, rules, prev_week)
    return base_week, 0


def analyze_rotation_and_rest(curr_week: dict, prev_week, employees: list, role_configs: dict, rules: dict) -> list:
    results = []
    for emp in employees:
        prev_cells = prev_week["assignments"].get(emp["id"], [parse_raw_cell("") for _ in range(7)]) if prev_week else [parse_raw_cell("") for _ in range(7)]
        curr_cells = curr_week["assignments"].get(emp["id"], [parse_raw_cell("") for _ in range(7)])
        pool = role_configs.get(emp["cargo"], {}).get("rotation_pool", [])

        prev_primary = get_primary_shift(prev_cells, "") if prev_week else ""
        prev_sun = prev_cells[6]["raw"] or "—"
        prev_off = next((i for i, c in enumerate(prev_cells) if c["base"] == "L"), None)
        curr_off = next((i for i, c in enumerate(curr_cells) if c["base"] == "L"), None)

        combined = (prev_cells + curr_cells) if prev_week else list(curr_cells)
        offset_curr = 7 if prev_week else 0

        max_streak, streak = 0, 0
        for c in combined:
            if parse_shift_window(c["base"]):
                streak += 1
                max_streak = max(max_streak, streak)
            else:
                streak = 0

        days_before_off = None
        if curr_off is not None:
            cnt = 0
            for i in range(offset_curr + curr_off - 1, -1, -1):
                if parse_shift_window(combined[i]["base"]):
                    cnt += 1
                else:
                    break
            days_before_off = cnt

        shift_before, idx_before, shift_after, idx_after = "", -1, "", -1
        if curr_off is not None:
            for i in range(offset_curr + curr_off - 1, -1, -1):
                if combined[i]["base"] == "L":
                    continue
                if parse_shift_window(combined[i]["base"]):
                    shift_before, idx_before = combined[i]["base"], i
                    break
                break
            for i in range(offset_curr + curr_off + 1, len(combined)):
                if combined[i]["base"] == "L":
                    continue
                if parse_shift_window(combined[i]["base"]):
                    shift_after, idx_after = combined[i]["base"], i
                    break
                break

        rest_across = compute_rest_hours(combined[idx_before], idx_before, combined[idx_after], idx_after) if (shift_before and shift_after) else None
        first_work = next((c["base"] for c in curr_cells if parse_shift_window(c["base"])), "")

        if not emp["activo"] or not first_work:
            status = "ABSENCE"
            detail = "Inactivo" if not emp["activo"] else "Sin turnos / Novedad"
        elif len(pool) <= 1:
            status = "FIXED_ROLE"
            detail = f"Turno fijo ({first_work})"
        elif curr_off is not None and shift_before and shift_after and shift_before != shift_after:
            status = "ROTATED_AFTER_L"
            detail = f"Rotó tras L ({DIAS_SEMANA[curr_off][:3]}): {shift_before} → {shift_after}"
        elif prev_primary and first_work and prev_primary != first_work:
            status = "ROTATED_ON_MONDAY"
            detail = f"Rotó inicio: {prev_primary} → {first_work}"
        else:
            status = "SAME_SHIFT"
            detail = f"Mantuvo {first_work}"

        if days_before_off is None:
            interval_ok = max_streak <= rules["max_consecutive_work_days"]
        else:
            interval_ok = (rules["min_work_days_between_off"] <= days_before_off <= rules["max_consecutive_work_days"]) and (max_streak <= rules["max_consecutive_work_days"])

        rest_ok = rest_across is None or rest_across >= rules["min_rest_day_off"]
        results.append({
            "emp": emp,
            "prev_primary": prev_primary,
            "prev_sun": prev_sun,
            "prev_off": prev_off,
            "curr_off": curr_off,
            "days_before_off": days_before_off,
            "max_streak": max_streak,
            "rest_across": rest_across,
            "status": status,
            "detail": detail,
            "interval_ok": interval_ok,
            "rest_ok": rest_ok,
        })
    return results


def export_excel_bytes(week: dict, employees: list) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Horario Semanal"
    ws.freeze_panes = "D3"

    thin = Side(style="thin", color="334155")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.append(["", "", ""] + DIAS_SEMANA)
    date_headers = [format_short_es(add_days_iso(week["start_iso"], d)) for d in range(7)]
    ws.append(["CODIGO", "NOMBRE", "CARGO"] + date_headers)

    for r_idx, fill_hex in [(1, "F1F5F9"), (2, "E2E8F0")]:
        for col in range(1, 11):
            c = ws.cell(row=r_idx, column=col)
            c.font = Font(name="Calibri", size=11, bold=True, color="0F172A")
            c.alignment = Alignment(horizontal="center", vertical="middle")
            c.border = border
            c.fill = PatternFill(start_color=fill_hex, end_color=fill_hex, fill_type="solid")

    ordered = sorted(employees, key=lambda e: CARGOS_ORDEN.index(e["cargo"]) if e["cargo"] in CARGOS_ORDEN else 99)
    for emp in ordered:
        cells = week["assignments"].get(emp["id"], [parse_raw_cell("") for _ in range(7)])
        cod_val = int(emp["codigo"]) if emp["codigo"].isdigit() else emp["codigo"]
        ws.append([cod_val, emp["nombre"], emp["cargo"]] + [c["raw"] for c in cells])
        r_num = ws.max_row
        for col in range(1, 4):
            c = ws.cell(row=r_num, column=col)
            c.font = Font(name="Calibri", size=11, color="0F172A")
            c.alignment = Alignment(horizontal="center", vertical="middle")
            c.border = border
        for d in range(7):
            fill_hex, font_hex = get_cell_colors(cells[d])
            c = ws.cell(row=r_num, column=4 + d)
            c.font = Font(name="Calibri", size=10, bold=True, color=font_hex)
            c.fill = PatternFill(start_color=fill_hex, end_color=fill_hex, fill_type="solid")
            c.alignment = Alignment(horizontal="center", vertical="middle")
            c.border = border

    widths = [11, 25, 25] + [16] * 7
    for i, w in enumerate(widths, start=1):
        col_letter = ws.cell(row=1, column=i).column_letter
        ws.column_dimensions[col_letter].width = w

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def normalize_cargo(raw: str) -> str:
    s = str(raw or "").upper().strip()
    if "ANALISTA" in s:
        return "Analista de Operaciones"
    if "AUXILIAR" in s and "ALISTAMIENTO" in s:
        return "Auxiliar de Alistamiento"
    if "AUXILIAR" in s:
        return "Auxiliar de Operaciones"
    if "LIDER" in s or "LÍDER" in s:
        return "Operador Líder"
    if "TECNICO" in s or "TÉCNICO" in s:
        return "Técnico de Operaciones"
    if "CONDUCTOR" in s:
        return "Conductor de Operaciones"
    return "Auxiliar de Operaciones"


def import_matrix_rows(matrix: list, start_iso: str):
    header_idx = 0
    for i in range(min(12, len(matrix))):
        upper_row = [str(x or "").strip().upper() for x in matrix[i]]
        if any(x in ("CODIGO", "CÓDIGO") for x in upper_row) and "NOMBRE" in upper_row:
            header_idx = i
            break

    hdr = [str(x or "").strip().upper() for x in matrix[header_idx]]
    has_hdr = any(x in ("CODIGO", "CÓDIGO") for x in hdr) and "NOMBRE" in hdr
    idx_cod = next((i for i, x in enumerate(hdr) if x in ("CODIGO", "CÓDIGO")), 0) if has_hdr else 0
    idx_nom = hdr.index("NOMBRE") if has_hdr and "NOMBRE" in hdr else 1
    idx_car = hdr.index("CARGO") if has_hdr and "CARGO" in hdr else 2
    data_start = header_idx + 1 if has_hdr else 0

    if has_hdr and data_start < len(matrix):
        next_row = matrix[data_start]
        if not str(next_row[idx_cod] if idx_cod < len(next_row) else "").strip():
            data_start += 1

    day_start = max(idx_car, idx_nom) + 1
    emps, assigns = [], {}
    for r in range(data_start, len(matrix)):
        row = matrix[r]
        cod = re.sub(r"\.0$", "", str(row[idx_cod] if idx_cod < len(row) else "").strip())
        nom = str(row[idx_nom] if idx_nom < len(row) else "").strip()
        if not cod and not nom:
            continue
        cargo = normalize_cargo(row[idx_car] if idx_car < len(row) else "")
        shifts, has_any, has_at = [], False, False
        for d in range(7):
            val = str(row[day_start + d] if (day_start + d) < len(row) else "").strip()
            p = parse_raw_cell(val)
            if p["base"]:
                has_any = True
            if p["is_at"]:
                has_at = True
            shifts.append(p)
        eid = f"emp_{cod or r}"
        emps.append({"id": eid, "codigo": cod or str(r), "nombre": nom or f"Colab {cod}", "cargo": cargo, "puede_at": has_at, "activo": has_any})
        assigns[eid] = shifts

    if not emps:
        raise ValueError("No se encontraron colaboradores válidos en el archivo o texto.")
    return emps, {"id": f"week_{start_iso}", "label": build_week_label(start_iso), "start_iso": start_iso, "assignments": assigns}


def init_state():
    if "roster_ready" in st.session_state:
        return
    emps, assigns = [], {}
    for cod, nom, car, pat, act, shs in SEED_ROWS:
        eid = f"emp_{cod}"
        emps.append({"id": eid, "codigo": cod, "nombre": nom, "cargo": car, "puede_at": pat, "activo": act})
        assigns[eid] = [parse_raw_cell(s) for s in shs]
    st.session_state.roster_employees = emps
    st.session_state.roster_weeks = [{
        "id": "week_2026-10-26",
        "label": "Semana 26 oct — 1 nov",
        "start_iso": "2026-10-26",
        "assignments": assigns,
    }]
    st.session_state.roster_active_idx = 0
    st.session_state.roster_rules = dict(DEFAULT_RULES)
    st.session_state.roster_configs = {
        k: {"rotation_pool": list(v["rotation_pool"]), "daily_targets": dict(v["daily_targets"])}
        for k, v in DEFAULT_ROLE_CONFIGS.items()
    }
    st.session_state.roster_ready = True


init_state()

employees = st.session_state.roster_employees
weeks = st.session_state.roster_weeks
rules = st.session_state.roster_rules
role_configs = st.session_state.roster_configs
active_idx = min(st.session_state.roster_active_idx, len(weeks) - 1)
active_week = weeks[active_idx]
prev_week = weeks[active_idx - 1] if active_idx > 0 else None

st.title("Planificador de Turnos Operativos")
st.caption("Rotación semanal Noche → Tarde → Mañana · Descansos 12h / 24h · Cobertura Técnico AT")

top_c1, top_c2, top_c3 = st.columns([3, 1.4, 1.4])
with top_c1:
    week_labels = [f"{i+1}. {w['label']}" for i, w in enumerate(weeks)]
    sel = st.selectbox("Semana activa", week_labels, index=active_idx)
    new_idx = week_labels.index(sel)
    if new_idx != active_idx:
        st.session_state.roster_active_idx = new_idx
        st.rerun()
with top_c2:
    st.write("")
    if st.button("Generar Siguiente Semana", type="primary", use_container_width=True):
        nw, at_cnt = generate_next_week(weeks[-1], employees, role_configs, rules)
        weeks.append(nw)
        st.session_state.roster_active_idx = len(weeks) - 1
        st.success(f"Generada {nw['label']} ({at_cnt} apoyos AT).")
        st.rerun()
with top_c3:
    st.write("")
    st.download_button(
        "Descargar Excel (.xlsx)",
        data=export_excel_bytes(active_week, employees),
        file_name=f"Horario_Operaciones_{active_week['start_iso']}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

issues = validate_schedule(active_week, employees, role_configs, rules, prev_week)
rest_alerts = [i for i in issues if i["type"] != "UNCOVERED_TECH"]
tech_gaps = [i for i in issues if i["type"] == "UNCOVERED_TECH"]

m1, m2, m3, m4 = st.columns(4)
m1.metric("Personal activo", sum(1 for e in employees if e["activo"]))
m2.metric("Turnos AT asignados", sum(1 for e in employees for c in active_week["assignments"].get(e["id"], []) if c["is_at"]))
m3.metric("Alertas descanso / racha", len(rest_alerts))
m4.metric("Faltantes Técnico", len(tech_gaps))

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "Horario Semanal",
    "Rotación y Descansos",
    "Cobertura y Apoyo AT",
    "Personal y Reglas",
    "Importar / Exportar Excel",
])

with tab1:
    fc1, fc2, fc3 = st.columns([2, 1.5, 1.5])
    with fc1:
        filtro_cargo = st.selectbox("Filtrar por cargo", ["TODOS"] + CARGOS_ORDEN, key="f_cargo")
    with fc2:
        show_prev_cols = st.checkbox("Ver ref. semana anterior", value=bool(prev_week), disabled=(prev_week is None))
    with fc3:
        if st.button("Auto-cubrir Técnicos (AT)", use_container_width=True):
            upd_w, cnt = auto_assign_tecnico_at(active_week, employees, role_configs, rules, prev_week)
            weeks[active_idx] = upd_w
            st.success(f"Apoyos AT actualizados ({cnt} asignados).")
            st.rerun()

    ordered_emps = sorted(employees, key=lambda e: CARGOS_ORDEN.index(e["cargo"]) if e["cargo"] in CARGOS_ORDEN else 99)
    if filtro_cargo != "TODOS":
        ordered_emps = [e for e in ordered_emps if e["cargo"] == filtro_cargo]

    analysis_map = {a["emp"]["id"]: a for a in analyze_rotation_and_rest(active_week, prev_week, employees, role_configs, rules)} if prev_week else {}
    day_cols = [f"{DIAS_SEMANA[d][:3]} {format_short_es(add_days_iso(active_week['start_iso'], d))}" for d in range(7)]

    rows_view = []
    cell_style_map = {}
    for emp in ordered_emps:
        row_dict = {"CODIGO": emp["codigo"], "NOMBRE": emp["nombre"], "CARGO": emp["cargo"]}
        if prev_week and show_prev_cols:
            an = analysis_map.get(emp["id"])
            prev_l = DIAS_SEMANA[an["prev_off"]][:3] if an and an["prev_off"] is not None else "—"
            row_dict["SEM ANT (Dom / L)"] = f"{an['prev_sun']} | L:{prev_l}" if an else "—"
            d_txt = f"{an['days_before_off']}d→L" if an and an["days_before_off"] is not None else "—"
            row_dict["DIAS P/ L / ROT."] = f"{d_txt} · {an['detail']}" if an else "—"
        cells = active_week["assignments"].get(emp["id"], [parse_raw_cell("") for _ in range(7)])
        for d, col_name in enumerate(day_cols):
            row_dict[col_name] = cells[d]["raw"]
            bg, fg = get_cell_colors(cells[d])
            cell_style_map[(emp["codigo"], col_name)] = f"background-color: #{bg}; color: #{fg}; font-weight: 600;"
        rows_view.append(row_dict)

    df_view = pd.DataFrame(rows_view)

    def style_matrix_row(s):
        styles = []
        cod = s["CODIGO"]
        for col_name in s.index:
            styles.append(cell_style_map.get((cod, col_name), ""))
        return styles

    st.dataframe(df_view.style.apply(style_matrix_row, axis=1), use_container_width=True, height=520, hide_index=True)

    with st.expander("Editar turno de un colaborador"):
        ec1, ec2, ec3, ec4 = st.columns([2, 1.2, 1.5, 1])
        emp_names = [f"{e['codigo']} - {e['nombre']} ({e['cargo']})" for e in ordered_emps]
        if emp_names:
            sel_e_txt = ec1.selectbox("Colaborador", emp_names)
            sel_e = ordered_emps[emp_names.index(sel_e_txt)]
            sel_d = ec2.selectbox("Día", DIAS_SEMANA)
            d_idx = DIAS_SEMANA.index(sel_d)
            curr_c = active_week["assignments"][sel_e["id"]][d_idx]
            opts = TURNOS_CATALOGO + ESTADOS_ESPECIALES
            idx_base = opts.index(curr_c["base"]) if curr_c["base"] in opts else 0
            new_base = ec3.selectbox("Turno / Estado", opts, index=idx_base)
            new_at = ec4.checkbox("AT (Cubre Técnico)", value=curr_c["is_at"], disabled=(new_base in ESTADOS_ESPECIALES))
            new_cap = st.checkbox("CAP (Capacitación)", value=curr_c["is_cap"], disabled=(new_base in ESTADOS_ESPECIALES))
            if st.button("Guardar cambio en celda"):
                active_week["assignments"][sel_e["id"]][d_idx] = parse_raw_cell(format_cell_raw(new_base, new_at, new_cap), locked=True)
                st.success("Turno actualizado.")
                st.rerun()

    if rest_alerts:
        st.warning("Alertas de descanso (12h entre turnos / 24h en L / días seguidos):")
        for a in rest_alerts:
            st.write(f"- **{DIAS_SEMANA[a['day_index']]}:** {a['message']}")
    else:
        st.success("Sin conflictos de descanso 12h/24h ni rachas largas esta semana.")

with tab2:
    mode_rot = st.radio(
        "Modo de análisis",
        ["Comparar Semana Activa vs Anterior", "Simular Siguiente Semana antes de crear"],
        horizontal=True,
    )
    if mode_rot.startswith("Simular"):
        sim_w, sim_at = generate_next_week(weeks[-1], employees, role_configs, rules)
        eval_curr, eval_prev = sim_w, weeks[-1]
        if st.button(f"Aplicar y Crear {sim_w['label']}", type="primary"):
            weeks.append(sim_w)
            st.session_state.roster_active_idx = len(weeks) - 1
            st.success(f"Semana {sim_w['label']} creada.")
            st.rerun()
    else:
        eval_curr, eval_prev = active_week, prev_week

    st.caption(f"Ventana analizada: {(eval_prev['label'] + ' → ') if eval_prev else ''}{eval_curr['label']}")
    rot_data = analyze_rotation_and_rest(eval_curr, eval_prev, employees, role_configs, rules)
    rot_rows = []
    for r in rot_data:
        emp = r["emp"]
        prev_str = " | ".join([c["raw"] or "—" for c in eval_prev["assignments"].get(emp["id"], [])]) if eval_prev else "Sin semana previa"
        curr_str = " | ".join([c["raw"] or "—" for c in eval_curr["assignments"].get(emp["id"], [])])
        rot_rows.append({
            "Código": emp["codigo"],
            "Nombre": emp["nombre"],
            "Cargo": emp["cargo"],
            "Semana Anterior (Lun-Dom)": prev_str,
            "Semana Analizada (Lun-Dom)": curr_str,
            "Días seguidos p/ L": f"{r['days_before_off']}d" if r["days_before_off"] is not None else f"Sin L ({r['max_streak']}d)",
            "Horas en L": f"{r['rest_across']}h" if r["rest_across"] is not None else "—",
            "Rotación tras Día Libre": r["detail"],
            "Estado": "OK" if (r["interval_ok"] and r["rest_ok"]) else "Revisar",
        })
    st.dataframe(pd.DataFrame(rot_rows), use_container_width=True, hide_index=True)

with tab3:
    st.subheader("Cobertura diaria de Técnico de Operaciones (Directos + Apoyo AT)")
    tech_targets = role_configs["Técnico de Operaciones"]["daily_targets"]
    cov_rows = []
    for sh, tgt in tech_targets.items():
        row_c = {"Turno Técnico": sh, "Meta": tgt}
        for d in range(7):
            directs = [e["nombre"] for e in employees if e["activo"] and e["cargo"] == "Técnico de Operaciones" and covers_shift_window(active_week["assignments"][e["id"]][d]["base"], sh)]
            ats = [f"AT: {e['nombre']}" for e in employees if e["activo"] and e["cargo"] != "Técnico de Operaciones" and active_week["assignments"][e["id"]][d]["is_at"] and covers_shift_window(active_week["assignments"][e["id"]][d]["base"], sh)]
            tot = len(directs) + len(ats)
            row_c[DIAS_SEMANA[d][:3]] = f"{tot}/{tgt} (" + (", ".join(directs + ats) or "Sin cubrir") + ")"
        cov_rows.append(row_c)
    st.dataframe(pd.DataFrame(cov_rows), use_container_width=True, hide_index=True)

with tab4:
    st.subheader("Reglas de descanso y rotación")
    rc1, rc2, rc3, rc4 = st.columns(4)
    rules["min_rest_shifts"] = int(rc1.number_input("Mín. horas entre turnos", 8, 24, rules["min_rest_shifts"]))
    rules["min_rest_day_off"] = int(rc2.number_input("Mín. horas en Día Libre L", 12, 48, rules["min_rest_day_off"]))
    rules["min_work_days_between_off"] = int(rc3.number_input("Mín. días trabajados entre L", 3, 6, rules["min_work_days_between_off"]))
    rules["max_consecutive_work_days"] = int(rc4.number_input("Máx. días seguidos sin L", 5, 10, rules["max_consecutive_work_days"]))
    rules["rotate_after_day_off"] = st.checkbox("Rotar turno después del Día Libre (L)", value=rules["rotate_after_day_off"])
    rules["auto_assign_at"] = st.checkbox("Auto-asignar cobertura AT al generar semana", value=rules["auto_assign_at"])

    st.subheader("Secuencia Noche → Tarde → Mañana por cargo")
    for cargo in CARGOS_ORDEN:
        cfg = role_configs[cargo]
        new_seq = st.text_input(f"Secuencia {cargo}", value=", ".join(cfg["rotation_pool"]), key=f"seq_{cargo}")
        parsed_seq = [x.strip() for x in new_seq.split(",") if x.strip()]
        if parsed_seq:
            cfg["rotation_pool"] = parsed_seq
            cfg["daily_targets"] = {t: cfg["daily_targets"].get(t, 1) for t in parsed_seq}

    st.subheader("Agregar personal nuevo para cualquier cargo")
    with st.form("form_nuevo_colaborador", clear_on_submit=True):
        nc1, nc2, nc3, nc4 = st.columns([1, 2, 2, 1])
        n_cod = nc1.text_input("Código")
        n_nom = nc2.text_input("Nombre completo")
        n_car = nc3.selectbox("Cargo", CARGOS_ORDEN)
        n_at = nc4.checkbox("Habilitado AT")
        if st.form_submit_button("Agregar colaborador"):
            if n_cod.strip() and n_nom.strip():
                eid = f"emp_{n_cod.strip()}_{len(employees)}"
                employees.append({
                    "id": eid,
                    "codigo": n_cod.strip(),
                    "nombre": n_nom.strip(),
                    "cargo": n_car,
                    "puede_at": n_at and n_car in ("Analista de Operaciones", "Auxiliar de Operaciones"),
                    "activo": True,
                })
                for w in weeks:
                    w["assignments"][eid] = [parse_raw_cell("") for _ in range(7)]
                st.success(f"Agregado {n_nom.strip()} ({n_car}).")
                st.rerun()

    st.subheader("Directorio actual")
    df_dir = pd.DataFrame([
        {"Código": e["codigo"], "Nombre": e["nombre"], "Cargo": e["cargo"], "Habilitado AT": e["puede_at"], "Activo": e["activo"]}
        for e in employees
    ])
    edited_dir = st.data_editor(
        df_dir,
        column_config={"Cargo": st.column_config.SelectboxColumn("Cargo", options=CARGOS_ORDEN)},
        use_container_width=True,
        hide_index=True,
        key="dir_editor",
    )
    if st.button("Aplicar cambios del directorio"):
        for idx, row in edited_dir.iterrows():
            employees[idx]["cargo"] = row["Cargo"]
            employees[idx]["puede_at"] = bool(row["Habilitado AT"]) and row["Cargo"] in ("Analista de Operaciones", "Auxiliar de Operaciones")
            employees[idx]["activo"] = bool(row["Activo"])
        st.success("Directorio actualizado.")
        st.rerun()

with tab5:
    st.subheader("Importar semana desde Excel (.xlsx) o pegar desde Excel")
    imp_date = st.date_input("Fecha del Lunes de la semana importada", value=date.fromisoformat(active_week["start_iso"]))
    up_file = st.file_uploader("Subir archivo Excel (.xlsx / .xlsm)", type=["xlsx", "xlsm"])
    if up_file is not None and st.button("Cargar archivo Excel subido"):
        wb_in = load_workbook(up_file, data_only=True)
        ws_in = wb_in.active
        mat = [["" if v is None else v for v in row] for row in ws_in.iter_rows(values_only=True)]
        new_emps, new_w = import_matrix_rows(mat, imp_date.isoformat())
        st.session_state.roster_employees = new_emps
        st.session_state.roster_weeks = [new_w]
        st.session_state.roster_active_idx = 0
        st.success(f"Importados {len(new_emps)} colaboradores.")
        st.rerun()

    paste_txt = st.text_area("O pega aquí las celdas copiadas desde Excel (Ctrl + C / Ctrl + V)")
    if st.button("Cargar tabla pegada") and paste_txt.strip():
        mat = [line.split("\t") for line in paste_txt.splitlines() if line.strip()]
        new_emps, new_w = import_matrix_rows(mat, imp_date.isoformat())
        st.session_state.roster_employees = new_emps
        st.session_state.roster_weeks = [new_w]
        st.session_state.roster_active_idx = 0
        st.success(f"Importados {len(new_emps)} colaboradores desde texto pegado.")
        st.rerun()
