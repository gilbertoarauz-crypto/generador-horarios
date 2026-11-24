import html
import io
import re
import unicodedata
import datetime as dt
import pandas as pd
import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Planificador de Turnos Operativos", layout="wide")

DIAS_SEMANA = ["LUN", "MAR", "MIÉ", "JUE", "VIE", "SÁB", "DOM"]
MESES_ES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]
CARGOS = [
    "Analista de Operaciones",
    "Auxiliar de Operaciones",
    "Operador Líder",
    "Técnico de Operaciones",
    "Auxiliar de Alistamiento",
    "Conductor de Operaciones",
]
NON_WORKING = {"L", "VACACIONES", "SUSPEND", "INCAPACIDAD", "PERMISO"}

DEFAULT_REST_RULES = {
    "min_rest_between_hours": 12,
    "min_rest_off_hours": 24,
    "days_off_per_week": 1,
    "min_work_between_off": 5,
    "max_consecutive_days": 7,
    "rotate_after_day_off": True,
    "auto_assign_at": True,
}

# Secuencia de rotación: Noche -> Tarde -> Mañana
DEFAULT_ROLE_CONFIGS = {
    "Analista de Operaciones": {
        "rotation_pool": ["19:00-27:00", "11:00-19:00", "06:00-15:00", "03:00-11:00"],
        "daily_targets": {"19:00-27:00": 1, "11:00-19:00": 1, "06:00-15:00": 1, "03:00-11:00": 1},
        "saturday_targets": {},
        "sunday_targets": {},
    },
    "Auxiliar de Operaciones": {
        "rotation_pool": ["19:00-27:00", "23:00-31:00", "11:00-19:00", "03:00-11:00"],
        "daily_targets": {"19:00-27:00": 1, "23:00-31:00": 1, "11:00-19:00": 2, "03:00-11:00": 2},
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
        "daily_targets": {"19:00-27:00": 1, "11:00-19:00": 1, "03:00-11:00": 1},
        "saturday_targets": {},
        "sunday_targets": {},
    },
    "Auxiliar de Alistamiento": {
        "rotation_pool": ["22:00-30:00", "20:00-28:00", "08:00-17:00"],
        "daily_targets": {"22:00-30:00": 1, "20:00-28:00": 2, "08:00-17:00": 1},
        "saturday_targets": {},
        "sunday_targets": {},
    },
    "Conductor de Operaciones": {
        "rotation_pool": ["11:00-18:00", "03:00-10:00"],
        "daily_targets": {"11:00-18:00": 1, "03:00-10:00": 1},
        "saturday_targets": {},
        "sunday_targets": {},
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


def is_tech_shift_no_at(shift_code: str) -> bool:
    return str(shift_code or "").strip().startswith("19:00")


def get_shift_target_for_day(cfg: dict, shift_code: str, day_idx: int) -> int:
    if not cfg:
        return 0
    if day_idx == 5:
        sat = cfg.get("saturday_targets") or {}
        if shift_code in sat:
            return int(sat[shift_code])
    elif day_idx == 6:
        sun = cfg.get("sunday_targets") or {}
        if shift_code in sun:
            return int(sun[shift_code])
    return int((cfg.get("daily_targets") or {}).get(shift_code, 0))


def normalize_cargo(raw: str) -> str:
    s = unicodedata.normalize("NFKD", str(raw or "")).encode("ascii", "ignore").decode("ascii").upper().strip()
    if "ANALISTA" in s:
        return "Analista de Operaciones"
    if "AUXILIAR" in s and "ALISTAMIENTO" in s:
        return "Auxiliar de Alistamiento"
    if "AUXILIAR" in s:
        return "Auxiliar de Operaciones"
    if "LIDER" in s:
        return "Operador Líder"
    if "TECNICO" in s:
        return "Técnico de Operaciones"
    if "CONDUCTOR" in s:
        return "Conductor de Operaciones"
    return "Auxiliar de Operaciones"


def parse_raw_cell(raw_text: str, locked: bool = False) -> dict:
    clean = re.sub(r"\s+", " ", str(raw_text or "").strip().upper())
    if not clean:
        return {"raw": "", "base": "", "is_at": False, "is_cap": False, "locked": locked}

    tokens = clean.split(" ")
    first = tokens[0]
    rest = " ".join(tokens[1:])
    is_at = bool(re.search(r"\bAT\b|AT/CAP|CAP/AT", clean))
    is_cap = bool(re.search(r"\bCAP\b|AT/CAP|CAP/AT", clean))

    if first in {"L", "LIBRE", "VACACIONES", "VAC", "SUSPEND", "SUSPENDIDO", "INCAPACIDAD", "PERMISO"}:
        norm = "L" if first == "LIBRE" else ("VACACIONES" if first == "VAC" else ("SUSPEND" if first == "SUSPENDIDO" else first))
        return {
            "raw": norm,
            "base": norm,
            "is_at": False,
            "is_cap": False,
            "locked": locked or norm in {"VACACIONES", "SUSPEND"},
        }

    suffix = []
    if is_at and is_cap:
        suffix.append("AT/CAP")
    else:
        if is_at:
            suffix.append("AT")
        if is_cap:
            suffix.append("CAP")

    final_raw = f"{first} {' '.join(suffix)}" if suffix else (clean if rest else first)
    return {"raw": final_raw, "base": first, "is_at": is_at, "is_cap": is_cap, "locked": locked}


def format_cell_raw(base: str, is_at: bool = False, is_cap: bool = False) -> str:
    b = base.strip().upper()
    if not b:
        return ""
    if b in NON_WORKING:
        return b
    if is_at and is_cap:
        return f"{b} AT/CAP"
    if is_at:
        return f"{b} AT"
    if is_cap:
        return f"{b} CAP"
    return b


def parse_shift_window(base_shift: str):
    m = re.match(r"^(\d{1,2}):(\d{2})-(\d{1,2}):(\d{2})$", str(base_shift or "").strip())
    if not m:
        return None
    sh, sm, eh, em = map(int, m.groups())
    return (sh + sm / 60.0, eh + em / 60.0)


def compute_rest_hours(prev_base: str, prev_day: int, next_base: str, next_day: int):
    w1 = parse_shift_window(prev_base)
    w2 = parse_shift_window(next_base)
    if not w1 or not w2:
        return None
    end_prev = prev_day * 24.0 + w1[1]
    start_next = next_day * 24.0 + w2[0]
    return round((start_next - end_prev) * 10) / 10.0


def is_transition_valid(prev_base, prev_day: int, next_base: str, next_day: int, crossed_off: bool, rules: dict) -> bool:
    if not prev_base:
        return True
    rest = compute_rest_hours(prev_base, prev_day, next_base, next_day)
    if rest is None:
        return True
    min_req = rules["min_rest_off_hours"] if crossed_off else rules["min_rest_between_hours"]
    return rest >= min_req


def get_primary_shift(cells: list, fallback: str = "") -> str:
    counts = {}
    for c in cells:
        b = c.get("base", "")
        if parse_shift_window(b):
            counts[b] = counts.get(b, 0) + 1
    if not counts:
        return fallback
    return max(counts.items(), key=lambda kv: kv[1])[0]


def covers_shift_window(assigned_base: str, target_base: str) -> bool:
    if assigned_base == target_base:
        return True
    w_a = parse_shift_window(assigned_base)
    w_t = parse_shift_window(target_base)
    if not w_a or not w_t:
        return False
    return w_a[0] <= w_t[0] and w_a[1] >= w_t[1]


def format_date_es(d: dt.date) -> str:
    return f"{d.day}-{MESES_ES[d.month - 1]}"


def build_week_label(start_date: dt.date) -> str:
    end_date = start_date + dt.timedelta(days=6)
    return f"Semana {format_date_es(start_date)} — {format_date_es(end_date)}"


def get_week_day_headers(start_date: dt.date) -> list:
    return [f"{DIAS_SEMANA[i]} {format_date_es(start_date + dt.timedelta(days=i))}" for i in range(7)]


def compute_uncovered_tasks(week: dict, employees: list, role_configs: dict) -> list:
    uncovered = []
    active = [e for e in employees if e["activo"]]
    for cargo in CARGOS:
        cfg = role_configs.get(cargo, {})
        shifts = list(dict.fromkeys(cfg.get("rotation_pool", []) + list((cfg.get("daily_targets") or {}).keys())))
        for d in range(7):
            for sc in shifts:
                tgt = get_shift_target_for_day(cfg, sc, d)
                if tgt <= 0:
                    continue
                no_at = (cargo == "Técnico de Operaciones" and is_tech_shift_no_at(sc))
                assigned = 0
                for e in active:
                    cell = week["assignments"].get(e["id"], [{}] * 7)[d]
                    b = cell.get("base", "")
                    if not b:
                        continue
                    if e["cargo"] == cargo and covers_shift_window(b, sc):
                        assigned += 1
                    elif cargo == "Técnico de Operaciones" and (not no_at) and cell.get("is_at") and covers_shift_window(b, sc):
                        assigned += 1
                if assigned < tgt:
                    uncovered.append({
                        "day_idx": d,
                        "DÍA": DIAS_SEMANA[d],
                        "CARGO": cargo,
                        "TURNO": sc,
                        "ASIGNADOS / META": f"{assigned} / {tgt}",
                        "FALTANTES": tgt - assigned,
                        "MOTIVO": "Requiere Técnico directo (19:00 no admite AT)" if no_at else "Personal insuficiente, novedades o descanso 12h/24h",
                    })
    return uncovered


def import_matrix_to_week(matrix: list, start_date: dt.date, employees: list, weeks: list):
    if not matrix:
        raise ValueError("No se encontraron filas válidas para importar.")

    header_idx = -1
    for i in range(min(12, len(matrix))):
        row_up = [str(c or "").strip().upper() for c in matrix[i]]
        if any(c in {"CODIGO", "CÓDIGO"} for c in row_up) and any(c == "NOMBRE" for c in row_up):
            header_idx = i
            break

    has_header = header_idx != -1
    hdr_row = [str(c or "").strip().upper() for c in matrix[header_idx]] if has_header else []
    idx_cod = next((i for i, c in enumerate(hdr_row) if c in {"CODIGO", "CÓDIGO"}), 0) if has_header else 0
    idx_nom = next((i for i, c in enumerate(hdr_row) if c == "NOMBRE"), 1) if has_header else 1
    idx_car = next((i for i, c in enumerate(hdr_row) if c == "CARGO"), 2) if has_header else 2
    day_start = max(idx_car, idx_nom) + 1

    data_start = header_idx + 1 if has_header else 0
    if has_header and data_start < len(matrix):
        nxt = matrix[data_start]
        c_val = str(nxt[idx_cod] if len(nxt) > idx_cod else "").strip()
        d_val = str(nxt[day_start] if len(nxt) > day_start else "").strip()
        if not c_val and d_val:
            data_start += 1

    emp_by_code = {e["codigo"]: e for e in employees}
    imp_assign = {e["id"]: [parse_raw_cell("") for _ in range(7)] for e in employees}
    imported_count = 0

    for r_i in range(data_start, len(matrix)):
        r = matrix[r_i]
        cod = str(r[idx_cod] if len(r) > idx_cod else "").strip()
        if cod.endswith(".0"):
            cod = cod[:-2]
        nom = str(r[idx_nom] if len(r) > idx_nom else "").strip()
        if not cod and not nom:
            continue
        if cod.upper() in {"CODIGO", "CÓDIGO"}:
            continue

        car_raw = str(r[idx_car] if len(r) > idx_car else "").strip()
        cells = [parse_raw_cell(r[day_start + d] if len(r) > day_start + d else "") for d in range(7)]
        has_work = any(c["base"] for c in cells)
        has_at = any(c["is_at"] for c in cells)

        if cod in emp_by_code:
            emp = emp_by_code[cod]
            if car_raw:
                emp["cargo"] = normalize_cargo(car_raw)
            if has_at:
                emp["puede_at"] = True
            emp["activo"] = has_work or emp["activo"]
            imp_assign[emp["id"]] = cells
        else:
            eid = f"emp_{cod or r_i}_{len(employees)+1}"
            new_emp = {
                "id": eid,
                "codigo": cod or str(r_i),
                "nombre": nom or f"Colaborador {cod}",
                "cargo": normalize_cargo(car_raw),
                "puede_at": has_at,
                "activo": has_work,
            }
            employees.append(new_emp)
            emp_by_code[new_emp["codigo"]] = new_emp
            for w in weeks:
                w["assignments"][eid] = [parse_raw_cell("") for _ in range(7)]
            imp_assign[eid] = cells

        imported_count += 1

    if imported_count == 0:
        raise ValueError("No se detectaron colaboradores. Verifica columnas CODIGO, NOMBRE, CARGO y los 7 días.")

    return {
        "id": f"week_{start_date.isoformat()}",
        "label": build_week_label(start_date),
        "start_date": start_date,
        "assignments": imp_assign,
    }, imported_count


def auto_assign_tecnico_at(week: dict, employees: list, role_configs: dict, rules: dict, prev_week=None):
    assignments = {eid: [dict(c) for c in row] for eid, row in week["assignments"].items()}

    for emp in employees:
        if not emp["activo"] or emp["cargo"] == "Técnico de Operaciones":
            continue
        row = assignments.get(emp["id"], [])
        for d in range(7):
            if d < len(row) and row[d].get("is_at") and not row[d].get("locked"):
                row[d] = parse_raw_cell(format_cell_raw(row[d]["base"], False, row[d].get("is_cap", False)))

    tech_cfg = role_configs.get("Técnico de Operaciones", {})
    tech_shifts = list(dict.fromkeys(tech_cfg.get("rotation_pool", []) + list((tech_cfg.get("daily_targets") or {}).keys())))
    at_totals = {e["id"]: sum(1 for c in assignments.get(e["id"], []) if c.get("is_at")) for e in employees}
    assigned_count = 0

    for d in range(7):
        for shift_code in tech_shifts:
            target = get_shift_target_for_day(tech_cfg, shift_code, d)
            if target <= 0 or is_tech_shift_no_at(shift_code):
                continue

            def count_covered():
                cnt = 0
                for e in employees:
                    if not e["activo"]:
                        continue
                    cell = assignments.get(e["id"], [{}] * 7)[d]
                    if e["cargo"] == "Técnico de Operaciones" and covers_shift_window(cell.get("base", ""), shift_code):
                        cnt += 1
                    elif cell.get("is_at") and covers_shift_window(cell.get("base", ""), shift_code):
                        cnt += 1
                return cnt

            covered = count_covered()
            while covered < target:
                same_cands = [
                    e for e in employees
                    if e["activo"]
                    and e["puede_at"]
                    and e["cargo"] in {"Auxiliar de Operaciones", "Analista de Operaciones"}
                    and not assignments[e["id"]][d].get("is_at")
                    and not assignments[e["id"]][d].get("locked")
                    and assignments[e["id"]][d].get("base") == shift_code
                ]
                same_cands.sort(key=lambda e: (at_totals.get(e["id"], 0), 0 if e["cargo"] == "Auxiliar de Operaciones" else 1))

                if same_cands:
                    ch = same_cands[0]
                    old = assignments[ch["id"]][d]
                    assignments[ch["id"]][d] = parse_raw_cell(format_cell_raw(old["base"], True, old.get("is_cap", False)))
                    at_totals[ch["id"]] = at_totals.get(ch["id"], 0) + 1
                    assigned_count += 1
                    covered += 1
                    continue

                flex_cands = []
                for e in employees:
                    if not (e["activo"] and e["puede_at"] and e["cargo"] in {"Auxiliar de Operaciones", "Analista de Operaciones"}):
                        continue
                    row = assignments.get(e["id"], [])
                    cell = row[d]
                    if cell.get("is_at") or cell.get("locked") or cell.get("base", "") in NON_WORKING or not cell.get("base"):
                        continue

                    prev_s, prev_off, crossed = None, d - 1, False
                    for k in range(d - 1, -1, -1):
                        b = row[k].get("base", "")
                        if b == "L":
                            crossed = True
                            continue
                        if parse_shift_window(b):
                            prev_s, prev_off = b, k
                            break
                        break
                    if not prev_s and d == 0 and prev_week and e["id"] in prev_week["assignments"]:
                        prow = prev_week["assignments"][e["id"]]
                        for k in range(6, -1, -1):
                            b = prow[k].get("base", "")
                            if b == "L":
                                crossed = True
                                continue
                            if parse_shift_window(b):
                                prev_s, prev_off = b, k - 7
                                break
                            break
                    if not is_transition_valid(prev_s, prev_off, shift_code, d, crossed, rules):
                        continue

                    next_s, next_off, next_crossed = None, d + 1, False
                    for k in range(d + 1, 7):
                        b = row[k].get("base", "")
                        if b == "L":
                            next_crossed = True
                            continue
                        if parse_shift_window(b):
                            next_s, next_off = b, k
                            break
                        break
                    if next_s and not is_transition_valid(shift_code, d, next_s, next_off, next_crossed, rules):
                        continue

                    flex_cands.append(e)

                def flex_sort_key(emp):
                    cur_s = assignments[emp["id"]][d].get("base", "")
                    cur_cnt = sum(
                        1 for o in employees
                        if o["activo"] and o["cargo"] == emp["cargo"] and assignments[o["id"]][d].get("base") == cur_s
                    )
                    cur_tgt = get_shift_target_for_day(role_configs.get(emp["cargo"], {}), cur_s, d)
                    has_surplus = 0 if cur_cnt > cur_tgt else 1
                    return (has_surplus, at_totals.get(emp["id"], 0), 0 if emp["cargo"] == "Auxiliar de Operaciones" else 1)

                flex_cands.sort(key=flex_sort_key)
                if flex_cands:
                    ch = flex_cands[0]
                    old = assignments[ch["id"]][d]
                    assignments[ch["id"]][d] = parse_raw_cell(format_cell_raw(shift_code, True, old.get("is_cap", False)))
                    at_totals[ch["id"]] = at_totals.get(ch["id"], 0) + 1
                    assigned_count += 1
                    covered += 1
                    continue
                break

    out_week = dict(week)
    out_week["assignments"] = assignments
    return out_week, assigned_count


def generate_next_week(
    prev_week: dict,
    employees: list,
    role_configs: dict,
    rules: dict,
    target_start_date: dt.date = None,
    preserve_locked_from: dict = None,
):
    new_start = target_start_date or (prev_week["start_date"] + dt.timedelta(days=7))
    new_assignments = {}

    off_counts = {}
    shift_counts = {}
    for e in employees:
        if e["cargo"] not in off_counts:
            off_counts[e["cargo"]] = [0] * 7
            shift_counts[e["cargo"]] = [{} for _ in range(7)]

    def get_active_pool(cargo: str) -> list:
        cfg = role_configs.get(cargo, {})
        pool = cfg.get("rotation_pool", ["08:00-17:00"]) or ["08:00-17:00"]
        filtered = [
            s for s in pool
            if get_shift_target_for_day(cfg, s, 0) > 0
            or get_shift_target_for_day(cfg, s, 5) > 0
            or get_shift_target_for_day(cfg, s, 6) > 0
        ]
        return filtered if filtered else pool

    def pick_best_shift_for_days(cargo: str, active_pool: list, day_indices: list, prev_shift_ref: str, is_valid_fn) -> str:
        cfg = role_configs.get(cargo, {})
        counts_by_day = shift_counts[cargo]
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
                target = get_shift_target_for_day(cfg, cand, d)
                cur = counts_by_day[d].get(cand, 0)
                if cur < target:
                    deficit_score += 450 if (cargo == "Técnico de Operaciones" and is_tech_shift_no_at(cand)) else 200
                else:
                    surplus_penalty += (cur - target + 1) * 70

            rot_bonus = 0
            if cand == ideal_next:
                rot_bonus += 35
            if len(active_pool) > 1 and cand == prev_shift_ref:
                rot_bonus -= 45

            total_score = deficit_score - surplus_penalty + rot_bonus
            if total_score > best_score:
                best_score = total_score
                best_shift = cand

        return best_shift if best_score != float("-inf") else ideal_next

    for emp_idx, emp in enumerate(employees):
        prev_cells = prev_week["assignments"].get(emp["id"], [])
        existing_cells = (preserve_locked_from or {}).get("assignments", {}).get(emp["id"], [])

        if not emp["activo"]:
            new_assignments[emp["id"]] = [
                dict(existing_cells[d]) if d < len(existing_cells) and existing_cells[d].get("locked") else parse_raw_cell("")
                for d in range(7)
            ]
            continue

        cfg = role_configs.get(emp["cargo"], {})
        active_pool = get_active_pool(emp["cargo"])
        primary_prev = get_primary_shift(prev_cells, active_pool[0])
        pool_idx = active_pool.index(primary_prev) if primary_prev in active_pool else -1
        next_target = active_pool[(pool_idx + 1) % len(active_pool)] if pool_idx >= 0 else active_pool[emp_idx % len(active_pool)]

        prev_off_idx = next((i for i, c in enumerate(prev_cells) if c.get("base") == "L"), -1)
        worked_tail = 0
        for pd_idx in range(len(prev_cells) - 1, -1, -1):
            if parse_shift_window(prev_cells[pd_idx].get("base", "")):
                worked_tail += 1
            else:
                break

        locked_off_idx = next(
            (i for i, c in enumerate(existing_cells) if c.get("locked") and c.get("base") == "L"),
            -1,
        )
        ideal_off = (prev_off_idx + 1) % 7 if prev_off_idx >= 0 else emp_idx % 7
        cargo_offs = off_counts[emp["cargo"]]
        if locked_off_idx >= 0:
            best_off = locked_off_idx
        else:
            best_off = ideal_off
            best_score = float("inf")
            for cand_day in range(7):
                if cand_day < len(existing_cells) and existing_cells[cand_day].get("locked"):
                    continue
                worked_until = worked_tail + cand_day
                pen = cargo_offs[cand_day] * 12
                if worked_until > rules["max_consecutive_days"]:
                    pen += 5000 + (worked_until - rules["max_consecutive_days"]) * 500
                elif worked_until < rules["min_work_between_off"]:
                    pen += 4000 + (rules["min_work_between_off"] - worked_until) * 500
                if cand_day != ideal_off:
                    pen += 2
                pen += abs(worked_until - 6)
                if pen < best_score:
                    best_score = pen
                    best_off = cand_day
        cargo_offs[best_off] += 1

        sunday_shift, sunday_off, sunday_crossed = None, -1, False
        for pd_idx in range(len(prev_cells) - 1, -1, -1):
            b = prev_cells[pd_idx].get("base", "")
            if b == "L":
                sunday_crossed = True
                continue
            if parse_shift_window(b):
                sunday_shift = b
                sunday_off = pd_idx - 7
                break
            break

        pre_days = [d for d in range(best_off)]
        post_days = [d for d in range(best_off + 1, 7)]

        sunday_enabled = any(get_shift_target_for_day(cfg, sunday_shift, d) > 0 for d in pre_days) if sunday_shift else False
        can_keep_sun = (
            is_transition_valid(sunday_shift, sunday_off, sunday_shift, 0, sunday_crossed, rules)
            if (sunday_shift and sunday_enabled)
            else False
        )

        if rules["rotate_after_day_off"] and not sunday_crossed and sunday_shift and can_keep_sun and best_off > 0:
            pre_shift = sunday_shift
        elif pre_days:
            pre_shift = pick_best_shift_for_days(
                emp["cargo"],
                active_pool,
                pre_days,
                sunday_shift or primary_prev,
                lambda cand: (not sunday_shift) or is_transition_valid(sunday_shift, sunday_off, cand, 0, sunday_crossed, rules),
            )
        else:
            pre_shift = next_target

        post_shift = pre_shift
        if post_days:
            ref_before_l = pre_shift if best_off > 0 else (sunday_shift or primary_prev)
            if best_off > 0:
                post_shift = pick_best_shift_for_days(
                    emp["cargo"],
                    active_pool,
                    post_days,
                    ref_before_l,
                    lambda cand: is_transition_valid(pre_shift, best_off - 1, cand, best_off + 1, True, rules),
                )
            else:
                post_shift = pick_best_shift_for_days(
                    emp["cargo"],
                    active_pool,
                    post_days,
                    ref_before_l,
                    lambda cand: (not sunday_shift) or is_transition_valid(sunday_shift, sunday_off, cand, 1, True, rules),
                )

        row = []
        counts_by_day = shift_counts[emp["cargo"]]
        for d in range(7):
            locked_cell = existing_cells[d] if d < len(existing_cells) and existing_cells[d].get("locked") else None
            if locked_cell:
                row.append(dict(locked_cell))
                b_lk = locked_cell.get("base", "")
                if parse_shift_window(b_lk):
                    counts_by_day[d][b_lk] = counts_by_day[d].get(b_lk, 0) + 1
            elif d == best_off:
                row.append(parse_raw_cell("L"))
            elif d < best_off:
                row.append(parse_raw_cell(pre_shift))
                counts_by_day[d][pre_shift] = counts_by_day[d].get(pre_shift, 0) + 1
            else:
                row.append(parse_raw_cell(post_shift))
                counts_by_day[d][post_shift] = counts_by_day[d].get(post_shift, 0) + 1
        new_assignments[emp["id"]] = row

    cargos_presentes = []
    for e in employees:
        if e["activo"] and e["cargo"] not in cargos_presentes:
            cargos_presentes.append(e["cargo"])

    for cargo in cargos_presentes:
        cfg = role_configs.get(cargo, {})
        active_pool = get_active_pool(cargo)
        ordered_pool = (
            sorted(active_pool, key=lambda s: 0 if is_tech_shift_no_at(s) else 1)
            if cargo == "Técnico de Operaciones"
            else active_pool
        )
        cargo_emps = [e for e in employees if e["activo"] and e["cargo"] == cargo]
        counts_by_day = shift_counts[cargo]

        for d in range(7):
            for s_need in ordered_pool:
                target_need = get_shift_target_for_day(cfg, s_need, d)
                allow_borrow_no_at = (cargo == "Técnico de Operaciones" and is_tech_shift_no_at(s_need))

                while counts_by_day[d].get(s_need, 0) < target_need:
                    donor_shifts = [
                        s_cand for s_cand, cur_cnt in counts_by_day[d].items()
                        if s_cand != s_need and (cur_cnt > get_shift_target_for_day(cfg, s_cand, d) or (allow_borrow_no_at and cur_cnt >= 1))
                    ]
                    donor_shifts.sort(
                        key=lambda s_cand: (
                            counts_by_day[d].get(s_cand, 0) - get_shift_target_for_day(cfg, s_cand, d),
                            counts_by_day[d].get(s_cand, 0),
                        ),
                        reverse=True,
                    )
                    if not donor_shifts:
                        break

                    moved = False
                    for extra_shift in donor_shifts:
                        for emp in cargo_emps:
                            row = new_assignments[emp["id"]]
                            cell = row[d]
                            if cell.get("base") != extra_shift or cell.get("locked"):
                                continue

                            prev_cells = prev_week["assignments"].get(emp["id"], [])
                            prev_s, prev_off, prev_crossed = None, d - 1, False
                            for k in range(d - 1, -1, -1):
                                b = row[k].get("base", "")
                                if b == "L":
                                    prev_crossed = True
                                    continue
                                if parse_shift_window(b):
                                    prev_s, prev_off = b, k
                                    break
                                break
                            if not prev_s and prev_cells:
                                for k in range(len(prev_cells) - 1, -1, -1):
                                    b = prev_cells[k].get("base", "")
                                    if b == "L":
                                        prev_crossed = True
                                        continue
                                    if parse_shift_window(b):
                                        prev_s, prev_off = b, k - 7
                                        break
                                    break
                            if not is_transition_valid(prev_s, prev_off, s_need, d, prev_crossed, rules):
                                continue

                            next_s, next_off, next_crossed = None, d + 1, False
                            for k in range(d + 1, 7):
                                b = row[k].get("base", "")
                                if b == "L":
                                    next_crossed = True
                                    continue
                                if parse_shift_window(b):
                                    next_s, next_off = b, k
                                    break
                                break
                            if next_s and not is_transition_valid(s_need, d, next_s, next_off, next_crossed, rules):
                                continue

                            new_assignments[emp["id"]][d] = parse_raw_cell(s_need)
                            counts_by_day[d][extra_shift] = counts_by_day[d].get(extra_shift, 1) - 1
                            counts_by_day[d][s_need] = counts_by_day[d].get(s_need, 0) + 1
                            moved = True
                            break
                        if moved:
                            break
                    if not moved:
                        break

    new_week = {
        "id": f"week_{new_start.isoformat()}",
        "label": build_week_label(new_start),
        "start_date": new_start,
        "assignments": new_assignments,
    }

    if rules.get("auto_assign_at", True):
        return auto_assign_tecnico_at(new_week, employees, role_configs, rules, prev_week)
    return new_week, 0


def validate_week_rest(week: dict, employees: list, rules: dict, prev_week=None) -> list:
    alerts = []
    for emp in employees:
        if not emp["activo"]:
            continue
        curr = week["assignments"].get(emp["id"], [])
        prev = prev_week["assignments"].get(emp["id"], []) if prev_week else []
        prev_shift, prev_day, crossed_off = None, -1, False

        for pd_idx, cell in enumerate(prev):
            b = cell.get("base", "")
            if b == "L":
                crossed_off = True
            elif parse_shift_window(b):
                prev_shift, prev_day, crossed_off = b, pd_idx - 7, False
            elif b:
                prev_shift, crossed_off = None, True

        for d_idx, cell in enumerate(curr):
            b = cell.get("base", "")
            if b == "L":
                crossed_off = True
                continue
            if not parse_shift_window(b):
                prev_shift, crossed_off = None, True
                continue
            if prev_shift:
                rest = compute_rest_hours(prev_shift, prev_day, b, d_idx)
                req = rules["min_rest_off_hours"] if crossed_off else rules["min_rest_between_hours"]
                if rest is not None and rest < req:
                    alerts.append(f"{emp['nombre']} ({DIAS_SEMANA[d_idx]}): {rest}h < {req}h ({prev_shift} → {b})")
            prev_shift, prev_day, crossed_off = b, d_idx, False

        streak = 0
        for c in (prev + curr):
            if parse_shift_window(c.get("base", "")):
                streak += 1
                if streak > rules["max_consecutive_days"]:
                    alerts.append(f"{emp['nombre']}: racha de {streak} días seguidos > {rules['max_consecutive_days']}")
                    break
            else:
                streak = 0
    return alerts


def analyze_rotation_details(curr_week: dict, prev_week, employees: list, role_configs: dict, rules: dict) -> list:
    items = []
    for emp in employees:
        if not emp["activo"]:
            continue
        prev_cells = prev_week["assignments"].get(emp["id"], [parse_raw_cell("")] * 7) if prev_week else []
        curr_cells = curr_week["assignments"].get(emp["id"], [parse_raw_cell("")] * 7)
        pool = role_configs.get(emp["cargo"], {}).get("rotation_pool", [])
        prev_off = next((i for i, c in enumerate(prev_cells) if c.get("base") == "L"), None)
        curr_off = next((i for i, c in enumerate(curr_cells) if c.get("base") == "L"), None)

        combined = (prev_cells + curr_cells) if prev_week else curr_cells
        offset_curr = 7 if prev_week else 0

        max_consec_14, streak = 0, 0
        for c in combined:
            if parse_shift_window(c.get("base", "")):
                streak += 1
                max_consec_14 = max(max_consec_14, streak)
            else:
                streak = 0

        before_l, after_l, days_before_l, rest_l = "", "", None, None
        if curr_off is not None:
            cnt = 0
            for i in range(offset_curr + curr_off - 1, -1, -1):
                if parse_shift_window(combined[i].get("base", "")):
                    cnt += 1
                else:
                    break
            days_before_l = cnt

            b_idx, a_idx = -1, -1
            for i in range(offset_curr + curr_off - 1, -1, -1):
                b = combined[i].get("base", "")
                if b == "L":
                    continue
                if parse_shift_window(b):
                    before_l, b_idx = b, i
                    break
                break
            for i in range(offset_curr + curr_off + 1, len(combined)):
                b = combined[i].get("base", "")
                if b == "L":
                    continue
                if parse_shift_window(b):
                    after_l, a_idx = b, i
                    break
                break
            if b_idx >= 0 and a_idx >= 0:
                rest_l = compute_rest_hours(before_l, b_idx, after_l, a_idx)

        first_curr = next((c.get("base") for c in curr_cells if parse_shift_window(c.get("base", ""))), "")
        prev_primary = get_primary_shift(prev_cells, "") if prev_week else ""

        if not first_curr:
            status, detail = "ABSENCE", "Sin turnos asignados / Novedad"
        elif len(pool) <= 1:
            status, detail = "FIXED_ROLE", f"Turno fijo ({first_curr})"
        elif curr_off is not None and before_l and after_l and before_l != after_l:
            status, detail = "ROTATED_AFTER_L", f"Rotó tras L ({DIAS_SEMANA[curr_off]}): {before_l} → {after_l}"
        elif prev_primary and first_curr != prev_primary:
            status, detail = "ROTATED_ON_MONDAY", f"Rotó desde inicio: {prev_primary} → {first_curr}"
        else:
            status = "SAME_SHIFT"
            detail = f"Mantuvo {first_curr} (igual a sem. anterior)" if prev_primary else f"Turno continuo {first_curr}"

        if days_before_l is None:
            off_ok = max_consec_14 <= rules["max_consecutive_days"]
        else:
            off_ok = (rules["min_work_between_off"] <= days_before_l <= rules["max_consecutive_days"]) and (max_consec_14 <= rules["max_consecutive_days"])
        rest_ok = (rest_l is None) or (rest_l >= rules["min_rest_off_hours"])

        items.append({
            "emp": emp,
            "prev_cells": prev_cells,
            "curr_cells": curr_cells,
            "prev_off": prev_off,
            "curr_off": curr_off,
            "days_before_l": days_before_l,
            "max_consec_14": max_consec_14,
            "rest_l": rest_l,
            "status": status,
            "detail": detail,
            "off_ok": off_ok,
            "rest_ok": rest_ok,
        })
    return items


def shift_badge_colors(cell: dict):
    b = str(cell.get("base", "")).strip().upper()
    if b == "SUSPEND":
        return ("#DC2626", "#FFFFFF", "#B91C1C")
    if b == "VACACIONES":
        return ("#FFFFFF", "#DC2626", "#E2E8F0")
    if b in {"INCAPACIDAD", "PERMISO"}:
        return ("#FFF1F2", "#9F1239", "#FECDD3")
    if b == "L":
        return ("#FCE4D6", "#C00000", "#FED7AA")
    if cell.get("is_at") or b in {"03:00-19:00", "11:00-27:00", "23:00-35:00"}:
        return ("#FBCFE8", "#0F172A", "#F472B6")
    if b.startswith("03:00-"):
        return ("#92D050", "#0F172A", "#65A30D")
    if b.startswith("11:00-"):
        return ("#FFC000", "#0F172A", "#D97706")
    if b.startswith("19:00-"):
        return ("#B4C6E7", "#0F172A", "#93C5FD")
    if b.startswith("23:00-"):
        return ("#FFF2CC", "#0F172A", "#FDE68A")
    if b.startswith("20:00-"):
        return ("#00B0F0", "#0F172A", "#0284C7")
    if b.startswith("22:00-"):
        return ("#A6A6A6", "#0F172A", "#64748B")
    if b in {"06:00-15:00", "08:00-17:00"}:
        return ("#FFFFFF", "#0F172A", "#CBD5E1")
    return ("#F8FAFC", "#475569", "#E2E8F0")


def build_mini_strip_html(cells: list) -> str:
    badges = []
    for idx in range(7):
        cell = cells[idx] if idx < len(cells) else {"base": "", "raw": ""}
        b = cell.get("base", "")
        bg, fg, br = shift_badge_colors(cell)
        short_txt = "L" if b == "L" else ("VAC" if b == "VACACIONES" else ("SUSP" if b == "SUSPEND" else (b[:5] if b else "—")))
        tip = html.escape(f"{DIAS_SEMANA[idx]}: {cell.get('raw') or 'Sin turno'}")
        d_ltr = DIAS_SEMANA[idx][0]
        badges.append(
            f'<div title="{tip}" style="background:{bg};color:{fg};border:1px solid {br};border-radius:4px;padding:3px 2px;text-align:center;font-family:monospace;font-size:10px;line-height:1.15;font-weight:600;">'
            f'<div style="font-size:9px;opacity:0.7;">{d_ltr}</div><div>{html.escape(short_txt)}</div></div>'
        )
    return f'<div style="display:grid;grid-template-columns:repeat(7,minmax(32px,1fr));gap:3px;">{"".join(badges)}</div>'


def export_excel_bytes(week: dict, employees: list) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = week["label"][:31]
    headers = ["CODIGO", "NOMBRE", "CARGO"] + get_week_day_headers(week["start_date"])

    thin = Side(style="thin", color="CBD5E1")
    border = Border(top=thin, bottom=thin, left=thin, right=thin)
    hdr_fill = PatternFill("solid", fgColor="1E293B")
    hdr_font = Font(bold=True, color="FFFFFF", size=10)

    ws.append(headers)
    for col_idx in range(1, len(headers) + 1):
        c = ws.cell(row=1, column=col_idx)
        c.fill = hdr_fill
        c.font = hdr_font
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border

    shift_fills = {
        "L": "FFE4E6",
        "VACACIONES": "EDE9FE",
        "SUSPEND": "F1F5F9",
        "INCAPACIDAD": "FEF3C7",
        "PERMISO": "E0F2FE",
        "19:00-27:00": "DBEAFE",
        "23:00-31:00": "FEF3C7",
        "11:00-19:00": "FEF08A",
        "03:00-11:00": "DCFCE7",
        "06:00-15:00": "F8FAFC",
    }

    for emp in employees:
        if not emp["activo"]:
            continue
        cells = week["assignments"].get(emp["id"], [])
        row_vals = [emp["codigo"], emp["nombre"], emp["cargo"]] + [c.get("raw", "") for c in cells]
        ws.append(row_vals)
        r_idx = ws.max_row
        for c_idx in range(1, len(headers) + 1):
            cell_xl = ws.cell(row=r_idx, column=c_idx)
            cell_xl.border = border
            if c_idx <= 3:
                cell_xl.font = Font(bold=(c_idx == 2), size=10)
                cell_xl.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell_obj = cells[c_idx - 4]
                bg = "FCE7F3" if cell_obj.get("is_at") else shift_fills.get(cell_obj.get("base", ""), "FFFFFF")
                cell_xl.fill = PatternFill("solid", fgColor=bg)
                cell_xl.font = Font(bold=True, size=9)
                cell_xl.alignment = Alignment(horizontal="center", vertical="center")

    for col_idx, w in enumerate([10, 22, 24] + [16] * 7, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


def init_state():
    if "roster_initialized" in st.session_state:
        return
    employees = []
    assignments = {}
    for code, name, cargo, puede_at, activo, shifts in SEED_ROWS:
        eid = f"emp_{code}"
        employees.append({
            "id": eid,
            "codigo": code,
            "nombre": name,
            "cargo": cargo,
            "puede_at": puede_at,
            "activo": activo,
        })
        assignments[eid] = [parse_raw_cell(s) for s in shifts]

    base_start = dt.date(2026, 10, 26)
    w1 = {
        "id": f"week_{base_start.isoformat()}",
        "label": build_week_label(base_start),
        "start_date": base_start,
        "assignments": assignments,
    }
    st.session_state.employees = employees
    st.session_state.rules = dict(DEFAULT_REST_RULES)
    st.session_state.role_configs = {
        k: {
            "rotation_pool": list(v["rotation_pool"]),
            "daily_targets": dict(v["daily_targets"]),
            "saturday_targets": dict(v.get("saturday_targets", {})),
            "sunday_targets": dict(v.get("sunday_targets", {})),
        }
        for k, v in DEFAULT_ROLE_CONFIGS.items()
    }
    w2, _ = generate_next_week(w1, employees, st.session_state.role_configs, st.session_state.rules)
    st.session_state.weeks = [w1, w2]
    st.session_state.active_week_idx = 1
    st.session_state.rot_mode = "compare"
    st.session_state.roster_toast = ""
    st.session_state.roster_initialized = True


init_state()

employees = st.session_state.employees
rules = st.session_state.rules
role_configs = st.session_state.role_configs
weeks = st.session_state.weeks
week_idx = max(0, min(st.session_state.active_week_idx, len(weeks) - 1))
st.session_state.active_week_idx = week_idx
active_week = weeks[week_idx]
prev_week = weeks[week_idx - 1] if week_idx > 0 else None


def re_rotate_active_week():
    base_ref = prev_week if prev_week is not None else active_week
    regen, at_cnt = generate_next_week(
        base_ref,
        employees,
        role_configs,
        rules,
        target_start_date=active_week["start_date"],
        preserve_locked_from=active_week,
    )
    regen["id"] = active_week["id"]
    regen["label"] = active_week["label"]
    st.session_state.weeks[week_idx] = regen
    st.session_state.roster_toast = (
        f"Semana '{active_week['label']}' re-rotada aplicando reglas, metas de Sáb/Dom y 19:00 solo Técnico ({at_cnt} AT asignados)."
    )


st.title("Planificador de Turnos Operativos")
st.caption("Rotación semanal Noche → Tarde → Mañana · Descansos 12h / 24h · Metas L–V / Sáb / Dom · 19:00 Técnico sin AT")

if st.session_state.get("roster_toast"):
    st.success(st.session_state.roster_toast)
    st.session_state.roster_toast = ""

c1, c2, c3, c4 = st.columns([3, 1.1, 1.9, 1.9])
with c1:
    week_options = [f"{i+1}. {w['label']}" for i, w in enumerate(weeks)]
    selected_idx = st.selectbox(
        "Semana activa",
        range(len(weeks)),
        index=week_idx,
        format_func=lambda i: week_options[i],
    )
    if selected_idx != week_idx:
        st.session_state.active_week_idx = selected_idx
        st.rerun()

with c2:
    st.write("")
    st.write("")
    if st.button("Eliminar", disabled=(len(weeks) <= 1), use_container_width=True):
        st.session_state.weeks.pop(week_idx)
        st.session_state.active_week_idx = max(0, week_idx - 1)
        st.session_state.roster_toast = "Semana eliminada."
        st.rerun()

with c3:
    st.write("")
    st.write("")
    if st.button("Generar Siguiente Semana", type="primary", use_container_width=True):
        next_w, at_cnt = generate_next_week(weeks[-1], employees, role_configs, rules)
        st.session_state.weeks.append(next_w)
        st.session_state.active_week_idx = len(st.session_state.weeks) - 1
        st.session_state.roster_toast = f"Se generó '{next_w['label']}' respetando Meta diaria y rotación ({at_cnt} turnos AT; 19:00 solo Técnico)."
        st.rerun()

with c4:
    st.write("")
    st.write("")
    st.download_button(
        "Descargar Excel (.xlsx)",
        data=export_excel_bytes(active_week, employees),
        file_name=f"Horario_{active_week['start_date'].isoformat()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

active_emps = [e for e in employees if e["activo"]]
nov_count = sum(
    1 for e in active_emps for c in active_week["assignments"].get(e["id"], [])
    if c.get("base") in {"VACACIONES", "SUSPEND", "INCAPACIDAD", "PERMISO"}
)
at_count = sum(1 for e in active_emps for c in active_week["assignments"].get(e["id"], []) if c.get("is_at"))
alerts = validate_week_rest(active_week, employees, rules, prev_week)
uncovered_list = compute_uncovered_tasks(active_week, employees, role_configs)
tech_missing = sum(u["FALTANTES"] for u in uncovered_list if u["CARGO"] == "Técnico de Operaciones")
total_missing_tasks = sum(u["FALTANTES"] for u in uncovered_list)

m1, m2, m3, m4, m5, m6 = st.columns(6)
m1.metric("Activos", len(active_emps))
m2.metric("Novedades", nov_count)
m3.metric("Turnos AT", at_count)
m4.metric("Descansos 12h/24h", "OK" if not alerts else f"{len(alerts)} alerta(s)")
m5.metric("Técnicos", "100% Cubierto" if tech_missing == 0 else f"{tech_missing} faltante(s)")
m6.metric("Tareas sin cubrir", "0 (Completo)" if total_missing_tasks == 0 else f"{total_missing_tasks} faltante(s)")

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "Horario Semanal",
    "Rotación y Descansos",
    "Cobertura y Apoyo AT",
    "Personal y Reglas",
    "Importar / Exportar Excel",
])


def style_shift_val(val: str) -> str:
    v = str(val).strip().upper()
    if "AT" in v:
        return "background-color: #FCE7F3; color: #9D174D; font-weight: 700;"
    if v == "L":
        return "background-color: #FFE4E6; color: #BE123C; font-weight: 700;"
    if v == "VACACIONES":
        return "background-color: #EDE9FE; color: #5B21B6; font-weight: 700;"
    if v == "SUSPEND":
        return "background-color: #F1F5F9; color: #334155; font-weight: 700;"
    if v.startswith("19:00-27:00"):
        return "background-color: #DBEAFE; color: #1E40AF; font-weight: 600;"
    if v.startswith("23:00-31:00") or v.startswith("22:00-30:00"):
        return "background-color: #FEF3C7; color: #92400E; font-weight: 600;"
    if v.startswith("11:00-19:00") or v.startswith("11:00-18:00") or v.startswith("20:00-28:00"):
        return "background-color: #FEF08A; color: #854D0E; font-weight: 600;"
    if v.startswith("03:00-11:00") or v.startswith("03:00-10:00") or v.startswith("08:00-17:00"):
        return "background-color: #DCFCE7; color: #166534; font-weight: 600;"
    return ""


with tab1:
    f1, f2, f3, f4, f5 = st.columns([2.1, 1.8, 1.3, 1.4, 1.4])
    with f1:
        search_q = st.text_input("Buscar código o nombre", placeholder="Ej. 671 o Guardao")
    with f2:
        role_filter = st.selectbox("Filtrar por cargo", ["TODOS"] + CARGOS)
    with f3:
        st.write("")
        st.write("")
        only_alerts = st.checkbox(f"Solo alertas ({len(alerts)})")
    with f4:
        st.write("")
        st.write("")
        if st.button("Auto-cubrir Técnicos (AT)", key="btn_at_top", use_container_width=True):
            updated_w, cnt = auto_assign_tecnico_at(active_week, employees, role_configs, rules, prev_week)
            st.session_state.weeks[week_idx] = updated_w
            st.session_state.roster_toast = f"Cobertura AT aplicada: {cnt} turno(s) asignado(s) (19:00 exclusivo Técnico)."
            st.rerun()
    with f5:
        st.write("")
        st.write("")
        if st.button("Re-rotar semana actual", key="btn_rerotar_top", type="primary", use_container_width=True):
            re_rotate_active_week()
            st.rerun()

    day_cols = get_week_day_headers(active_week["start_date"])
    rot_items_t1 = analyze_rotation_details(active_week, prev_week, employees, role_configs, rules)
    rot_map = {item["emp"]["codigo"]: item for item in rot_items_t1}
    alert_names = {a.split(":")[0].split(" (")[0] for a in alerts}

    grid_rows = []
    for emp in active_emps:
        if role_filter != "TODOS" and emp["cargo"] != role_filter:
            continue
        if search_q.strip() and search_q.strip().lower() not in f"{emp['codigo']} {emp['nombre']}".lower():
            continue
        if only_alerts and emp["nombre"] not in alert_names:
            continue
        cells = active_week["assignments"].get(emp["id"], [parse_raw_cell("")] * 7)
        r_info = rot_map.get(emp["codigo"], {})
        prev_cells = prev_week["assignments"].get(emp["id"], []) if prev_week else []
        prev_dom = prev_cells[6].get("raw", "—") if len(prev_cells) == 7 else "—"
        prev_l = DIAS_SEMANA[r_info["prev_off"]] if r_info.get("prev_off") is not None else "—"
        d_pl = f"{r_info['days_before_l']}d" if r_info.get("days_before_l") is not None else "—"
        row_dict = {
            "CODIGO": emp["codigo"],
            "NOMBRE": emp["nombre"],
            "CARGO": emp["cargo"],
            "SEM ANT": f"Dom: {prev_dom} | L: {prev_l}",
            "DÍAS P/L Y ROT.": f"{d_pl} · {r_info.get('detail', '')}",
        }
        for d in range(7):
            row_dict[day_cols[d]] = cells[d].get("raw", "") if d < len(cells) else ""
        grid_rows.append(row_dict)

    grid_df = pd.DataFrame(grid_rows)
    if grid_df.empty:
        st.info("No hay colaboradores para el filtro actual.")
    else:
        def apply_grid_styles(row):
            return [style_shift_val(row[col]) if col in day_cols else "" for col in row.index]

        st.dataframe(grid_df.style.apply(apply_grid_styles, axis=1), use_container_width=True, height=560)

    c_bot_l, c_bot_r = st.columns([3, 1.5])
    c_bot_l.caption(f"Mostrando {len(grid_df)} de {len(active_emps)} colaboradores activos.")
    with c_bot_r:
        if st.button("Re-rotar semana actual", key="btn_rerotar_bot", use_container_width=True):
            re_rotate_active_week()
            st.rerun()

    if alerts:
        with st.expander(f"Ver {len(alerts)} alerta(s) de descanso / racha", expanded=False):
            for a in alerts:
                st.warning(a)

    with st.expander("Editar turno o novedad de un colaborador"):
        e1, e2, e3, e4 = st.columns([2, 1.5, 1.5, 1.2])
        emp_edit = e1.selectbox("Colaborador", active_emps, format_func=lambda e: f"{e['codigo']} - {e['nombre']}")
        day_edit = e2.selectbox("Día", range(7), format_func=lambda d: day_cols[d])
        cur_val = active_week["assignments"][emp_edit["id"]][day_edit].get("raw", "")
        new_val = e3.text_input("Turno / Novedad (ej. 11:00-19:00, 11:00-19:00 AT, L, VACACIONES, SUSPEND)", value=cur_val)
        lock_cell = e4.checkbox("Fijar al rotar", value=active_week["assignments"][emp_edit["id"]][day_edit].get("locked", False))
        if st.button("Guardar turno"):
            active_week["assignments"][emp_edit["id"]][day_edit] = parse_raw_cell(new_val, locked=lock_cell)
            st.session_state.roster_toast = f"Turno actualizado para {emp_edit['nombre']} ({day_cols[day_edit]})."
            st.rerun()

with tab2:
    with st.container(border=True):
        h_l, h_r = st.columns([2.2, 1.8])
        with h_l:
            st.subheader("Análisis de Días para Descanso (L) y Rotación Tras Días Libres")
            st.caption("Compara la semana anterior contra la semana nueva para verificar cuántos días continuos trabaja cada colaborador antes de su día libre y cómo rota de turno después de descansar.")
        with h_r:
            mode_options = [f"Semana Activa ({active_week['label']})", "Previsualizar Siguiente Semana"]
            mode_idx = 1 if st.session_state.get("rot_mode") == "preview" else 0
            chosen_mode = st.radio("Modo de análisis", mode_options, index=mode_idx, horizontal=True, label_visibility="collapsed")
            st.session_state.rot_mode = "preview" if chosen_mode == "Previsualizar Siguiente Semana" else "compare"

        is_preview = st.session_state.rot_mode == "preview"
        if is_preview:
            preview_base = weeks[-1]
            preview_gen, preview_at = generate_next_week(preview_base, employees, role_configs, rules)
            eval_curr, eval_prev = preview_gen, preview_base
        else:
            eval_curr, eval_prev = active_week, prev_week

        rc1, rc2, rc3, rc4 = st.columns([1.2, 1.2, 1.4, 1.4])
        rules["min_work_between_off"] = int(rc1.number_input("Mín. días trabajados entre L", 3, 6, rules["min_work_between_off"], key="tab2_min_l"))
        rules["max_consecutive_days"] = int(rc2.number_input("Máx. días seguidos sin descanso", 5, 10, rules["max_consecutive_days"], key="tab2_max_l"))
        with rc3:
            st.write("")
            st.write("")
            rules["rotate_after_day_off"] = st.checkbox("Rotar turno después del Día Libre (L)", value=rules["rotate_after_day_off"], key="tab2_rot_l")
        with rc4:
            st.write("")
            st.write("")
            if is_preview:
                if st.button(f"Aplicar y Crear {eval_curr['label']}", type="primary", use_container_width=True, key="btn_apply_preview_t2"):
                    st.session_state.weeks.append(eval_curr)
                    st.session_state.active_week_idx = len(st.session_state.weeks) - 1
                    st.session_state.rot_mode = "compare"
                    st.session_state.roster_toast = f"Se aplicó y creó '{eval_curr['label']}' ({preview_at} cobertura(s) AT)."
                    st.rerun()
            else:
                if st.button("Simular Próxima Semana", use_container_width=True, key="btn_sim_next_t2"):
                    st.session_state.rot_mode = "preview"
                    st.rerun()

        rot_details = analyze_rotation_details(eval_curr, eval_prev, employees, role_configs, rules)
        eval_alerts = validate_week_rest(eval_curr, employees, rules, eval_prev)
        active_rot = [r for r in rot_details if r["status"] != "ABSENCE"]
        rotated_cnt = sum(1 for r in active_rot if r["status"] in {"ROTATED_AFTER_L", "ROTATED_ON_MONDAY"})
        interval_bad = sum(1 for r in active_rot if not r["off_ok"])
        days_list = [r["days_before_l"] for r in active_rot if r["days_before_l"] is not None]
        avg_days = round(sum(days_list) / len(days_list), 1) if days_list else 0.0

        st.divider()
        km1, km2, km3, km4, km5 = st.columns([2, 1, 1, 1, 1])
        km1.metric("Ventana comparada", f"{eval_prev['label']} → {eval_curr['label']}" if eval_prev else eval_curr["label"])
        km2.metric("Promedio días antes de L", f"{avg_days} días")
        km3.metric("Rotaciones aplicadas", f"{rotated_cnt} / {len(active_rot)}")
        km4.metric("Intervalo fuera de rango", interval_bad)
        km5.metric("Conflictos 12h / 24h", len(eval_alerts))

    rf1, rf2, rf3 = st.columns([1.8, 2.2, 2])
    cargo_rot_filter = rf1.selectbox("Filtrar por cargo", ["TODOS"] + CARGOS, format_func=lambda x: f"Todos los cargos ({len(active_emps)})" if x == "TODOS" else x, key="rot_cargo_flt")
    with rf2:
        st.write("")
        st.write("")
        only_no_rot = st.checkbox("Mostrar solo sin rotación o alertas de descanso", key="rot_only_alerts")
    with rf3:
        st.write("")
        st.write("")
        st.caption(f"Rango saludable: **{rules['min_work_between_off']}–{rules['max_consecutive_days']} días** · Descanso L ≥ **{rules['min_rest_off_hours']}h**")

    filtered_rot = [
        r for r in rot_details
        if (cargo_rot_filter == "TODOS" or r["emp"]["cargo"] == cargo_rot_filter)
        and (not only_no_rot or (not r["off_ok"] or not r["rest_ok"] or r["status"] == "SAME_SHIFT"))
    ]

    if not filtered_rot:
        st.info("No hay colaboradores que coincidan con el filtro seleccionado.")
    else:
        curr_col_title = "Nueva Semana Propuesta (Lun → Dom)" if is_preview else "Semana Actual (Lun → Dom)"
        rows_html = []
        for item in filtered_rot:
            emp = item["emp"]
            prev_strip = build_mini_strip_html(item["prev_cells"]) if eval_prev else '<span style="color:#94a3b8;">Sin semana previa</span>'
            curr_strip = build_mini_strip_html(item["curr_cells"])

            if item["status"] == "ABSENCE":
                dias_html = '<span style="color:#94a3b8;">—</span>'
            elif item["days_before_l"] is not None:
                c_ok = "#047857" if item["off_ok"] else "#B45309"
                p_lbl = f"{DIAS_SEMANA[item['prev_off']]} → " if item["prev_off"] is not None else ""
                c_lbl = f"L {DIAS_SEMANA[item['curr_off']]}" if item["curr_off"] is not None else ""
                dias_html = f'<div style="font-weight:700;color:{c_ok};">{item["days_before_l"]} días seg.</div><div style="font-size:10px;color:#64748B;">{p_lbl}{c_lbl}</div>'
            else:
                dias_html = f'<span style="color:#B45309;font-weight:700;">Sin L ({item["max_consec_14"]}d)</span>'

            if item["rest_l"] is not None:
                r_col = "#047857" if item["rest_ok"] else "#B91C1C"
                horas_html = f'<span style="font-weight:700;color:{r_col};">{item["rest_l"]}h</span>'
            else:
                horas_html = '<span style="color:#94a3b8;">—</span>'

            det_col = "#B45309" if item["status"] == "SAME_SHIFT" else "#1E293B"
            warn_extra = (
                f'<div style="font-size:11px;color:#B45309;">Fuera de rango ({rules["min_work_between_off"]}–{rules["max_consecutive_days"]} días antes de L)</div>'
                if (not item["off_ok"] and item["status"] != "ABSENCE") else ""
            )

            rows_html.append(
                f'<tr style="border-bottom:1px solid #E2E8F0;">'
                f'<td style="padding:8px 10px;vertical-align:top;">'
                f'<div style="font-weight:600;color:#0F172A;"><span style="color:#64748B;font-family:monospace;margin-right:6px;">{html.escape(emp["codigo"])}</span>{html.escape(emp["nombre"])}</div>'
                f'<div style="font-size:11px;color:#64748B;">{html.escape(emp["cargo"])}</div></td>'
                f'<td style="padding:6px 8px;vertical-align:middle;">{prev_strip}</td>'
                f'<td style="padding:6px 8px;vertical-align:middle;">{curr_strip}</td>'
                f'<td style="padding:6px 8px;text-align:center;vertical-align:middle;font-family:monospace;">{dias_html}</td>'
                f'<td style="padding:6px 8px;text-align:center;vertical-align:middle;font-family:monospace;">{horas_html}</td>'
                f'<td style="padding:6px 10px;vertical-align:middle;color:{det_col};font-weight:500;">{html.escape(item["detail"])}{warn_extra}</td>'
                f'</tr>'
            )

        table_html = (
            '<div style="overflow-x:auto;max-height:580px;overflow-y:auto;border:1px solid #CBD5E1;border-radius:8px;background:#FFFFFF;">'
            '<table style="width:100%;border-collapse:collapse;font-size:12px;">'
            '<thead><tr style="background:#F1F5F9;border-bottom:2px solid #CBD5E1;color:#1E293B;position:sticky;top:0;z-index:1;">'
            '<th style="padding:10px;text-align:left;min-width:190px;">Colaborador y Cargo</th>'
            '<th style="padding:10px;text-align:left;min-width:255px;">Semana Anterior (Lun → Dom)</th>'
            f'<th style="padding:10px;text-align:left;min-width:255px;">{curr_col_title}</th>'
            '<th style="padding:10px;text-align:center;min-width:115px;">Días p/ Descanso</th>'
            '<th style="padding:10px;text-align:center;min-width:85px;">Horas en L</th>'
            '<th style="padding:10px;text-align:left;min-width:220px;">Rotación tras Día Libre</th>'
            '</tr></thead>'
            f'<tbody>{"".join(rows_html)}</tbody></table></div>'
        )
        st.markdown(table_html, unsafe_allow_html=True)

with tab3:
    st.subheader("Tareas y Turnos Sin Cubrir (Faltantes de Meta)")
    st.caption("Muestra todas las tareas o turnos que quedaron por debajo de la meta configurada (Lun–Vie, Sábado o Domingo) en cualquier cargo.")
    flt_falt = st.selectbox(
        "Filtrar faltantes por cargo",
        ["TODOS"] + CARGOS,
        format_func=lambda x: f"Todos los cargos ({total_missing_tasks} faltante(s))" if x == "TODOS" else x,
        key="flt_uncovered_cargo",
    )
    filtered_uncovered = [u for u in uncovered_list if flt_falt == "TODOS" or u["CARGO"] == flt_falt]
    if not filtered_uncovered:
        st.success("No hay tareas pendientes: todas las metas diarias configuradas están 100% cubiertas.")
    else:
        df_unc = pd.DataFrame([
            {
                "DÍA": u["DÍA"],
                "CARGO": u["CARGO"],
                "TURNO": u["TURNO"],
                "ASIGNADOS / META": u["ASIGNADOS / META"],
                "FALTANTES": f"-{u['FALTANTES']}",
                "MOTIVO / OBSERVACIÓN": u["MOTIVO"],
            }
            for u in filtered_uncovered
        ])
        st.dataframe(df_unc, use_container_width=True)

    st.divider()
    st.subheader("Cobertura Diaria por Cargo y Turno vs Meta (L–V / Sáb / Dom)")
    st.caption("Nota: En Técnico de Operaciones, el turno 19:00-27:00 es cubierto únicamente por Técnicos directos (no admite AT).")
    cargo_cov_sel = st.selectbox("Seleccionar cargo para revisar cobertura", CARGOS, key="cov_cargo_sel")
    cfg_sel = role_configs[cargo_cov_sel]
    pool_cov = list(dict.fromkeys(cfg_sel["rotation_pool"] + list(cfg_sel["daily_targets"].keys())))
    day_cols_cov = get_week_day_headers(active_week["start_date"])
    cov_rows = []
    for sc in pool_cov:
        tgt_lv = get_shift_target_for_day(cfg_sel, sc, 0)
        tgt_sab = get_shift_target_for_day(cfg_sel, sc, 5)
        tgt_dom = get_shift_target_for_day(cfg_sel, sc, 6)
        no_at = (cargo_cov_sel == "Técnico de Operaciones" and is_tech_shift_no_at(sc))
        label_turno = f"{sc} (Solo Técnico)" if no_at else sc
        row_c = {"TURNO": label_turno, "META (L-V / S / D)": f"{tgt_lv} / {tgt_sab} / {tgt_dom}"}
        for d in range(7):
            tgt_d = get_shift_target_for_day(cfg_sel, sc, d)
            cnt = 0
            for e in active_emps:
                cell = active_week["assignments"][e["id"]][d]
                if e["cargo"] == cargo_cov_sel and covers_shift_window(cell.get("base", ""), sc):
                    cnt += 1
                elif cargo_cov_sel == "Técnico de Operaciones" and (not no_at) and cell.get("is_at") and covers_shift_window(cell.get("base", ""), sc):
                    cnt += 1
            row_c[day_cols_cov[d]] = f"{cnt} / {tgt_d}"
        cov_rows.append(row_c)
    st.dataframe(pd.DataFrame(cov_rows), use_container_width=True)

    st.subheader("Personal Habilitado para Cobertura Técnico (AT)")
    at_rows = []
    for e in employees:
        if not e["activo"] or not e["puede_at"]:
            continue
        at_days = sum(1 for c in active_week["assignments"][e["id"]] if c.get("is_at"))
        at_rows.append({
            "CÓDIGO": e["codigo"],
            "COLABORADOR": e["nombre"],
            "CARGO BASE": e["cargo"],
            "TURNOS AT ESTA SEMANA": at_days,
        })
    st.dataframe(pd.DataFrame(at_rows), use_container_width=True)

with tab4:
    st.subheader("Reglas Automáticas de Descanso y Rotación")
    r1, r2, r3, r4, r5 = st.columns(5)
    rules["min_rest_between_hours"] = int(r1.number_input("Mín. horas entre turnos", 8, 24, rules["min_rest_between_hours"]))
    rules["min_rest_off_hours"] = int(r2.number_input("Mín. horas en Día Libre L", 12, 48, rules["min_rest_off_hours"]))
    rules["days_off_per_week"] = int(r3.number_input("Días libres por semana", 1, 2, rules["days_off_per_week"]))
    rules["min_work_between_off"] = int(r4.number_input("Mín. días trabajados entre L", 3, 7, rules["min_work_between_off"]))
    rules["max_consecutive_days"] = int(r5.number_input("Máx. días seguidos sin L", 5, 10, rules["max_consecutive_days"]))
    rules["rotate_after_day_off"] = st.checkbox("Rotar turno después del Día Libre (L)", value=rules["rotate_after_day_off"])
    rules["auto_assign_at"] = st.checkbox("Auto-asignar cobertura AT al generar semana (excluye 19:00)", value=rules["auto_assign_at"])

    st.divider()
    st.subheader("Ciclo de Rotación (Noche → Tarde → Mañana) y Meta Diaria por Cargo")
    st.caption("Ajusta cuántas personas necesitas en Lun–Vie, Sábado y Domingo por cada turno (pon 0 donde no requieras). Luego pulsa 'Aplicar reglas y Re-rotar semana actual'.")
    cols_cfg = st.columns(2)
    for idx_c, cargo in enumerate(CARGOS):
        with cols_cfg[idx_c % 2]:
            with st.container(border=True):
                st.markdown(f"**{cargo}**")
                cfg = role_configs[cargo]
                cfg.setdefault("saturday_targets", {})
                cfg.setdefault("sunday_targets", {})
                pool_str = st.text_input(
                    "Secuencia de rotación semanal (Noche → Tarde → Mañana, separada por coma)",
                    value=", ".join(cfg["rotation_pool"]),
                    key=f"pool_{cargo}",
                )
                new_pool = [s.strip() for s in pool_str.split(",") if s.strip()]
                if new_pool:
                    cfg["rotation_pool"] = new_pool
                target_shifts = list(dict.fromkeys(cfg["rotation_pool"] + list(cfg["daily_targets"].keys())))
                if target_shifts:
                    hdr_t = st.columns([1.4, 1, 1, 1])
                    hdr_t[0].caption("Turno")
                    hdr_t[1].caption("Lun–Vie")
                    hdr_t[2].caption("Sábado")
                    hdr_t[3].caption("Domingo")
                    for s_code in target_shifts:
                        row_t = st.columns([1.4, 1, 1, 1])
                        row_t[0].markdown(f"`{s_code}`")
                        wk_val = int(
                            row_t[1].number_input(
                                "Lun-Vie",
                                min_value=0,
                                max_value=20,
                                value=int(cfg["daily_targets"].get(s_code, 0)),
                                key=f"tgt_wk_{cargo}_{s_code}",
                                label_visibility="collapsed",
                            )
                        )
                        cfg["daily_targets"][s_code] = wk_val
                        sat_default = int(cfg["saturday_targets"].get(s_code, wk_val))
                        cfg["saturday_targets"][s_code] = int(
                            row_t[2].number_input(
                                "Sáb",
                                min_value=0,
                                max_value=20,
                                value=sat_default,
                                key=f"tgt_sat_{cargo}_{s_code}",
                                label_visibility="collapsed",
                            )
                        )
                        sun_default = int(cfg["sunday_targets"].get(s_code, wk_val))
                        cfg["sunday_targets"][s_code] = int(
                            row_t[3].number_input(
                                "Dom",
                                min_value=0,
                                max_value=20,
                                value=sun_default,
                                key=f"tgt_sun_{cargo}_{s_code}",
                                label_visibility="collapsed",
                            )
                        )

    if st.button("Aplicar reglas y Re-rotar semana actual", key="btn_apply_rerotar_tab4", type="primary"):
        re_rotate_active_week()
        st.rerun()

    st.divider()
    st.subheader("Agregar personal nuevo para cualquier cargo")
    with st.form("add_emp_form", clear_on_submit=True):
        a1, a2, a3, a4 = st.columns([1.2, 2.2, 2, 1])
        new_code = a1.text_input("Código", placeholder="Ej. 14025")
        new_name = a2.text_input("Nombre completo", placeholder="Ej. Carlos Mendoza")
        new_cargo = a3.selectbox("Cargo", CARGOS)
        new_at = a4.checkbox("Habilitado AT", value=(new_cargo in {"Auxiliar de Operaciones", "Analista de Operaciones"}))
        if st.form_submit_button("Agregar colaborador", type="primary"):
            if new_code.strip() and new_name.strip():
                eid = f"emp_{new_code.strip()}_{len(employees)+1}"
                employees.append({
                    "id": eid,
                    "codigo": new_code.strip(),
                    "nombre": new_name.strip(),
                    "cargo": new_cargo,
                    "puede_at": new_at,
                    "activo": True,
                })
                for w in st.session_state.weeks:
                    w["assignments"][eid] = [parse_raw_cell("") for _ in range(7)]
                st.session_state.roster_toast = f"Colaborador {new_name.strip()} ({new_code.strip()}) agregado a {new_cargo}."
                st.rerun()
            else:
                st.error("Ingresa código y nombre.")

    st.subheader("Directorio de Personal")
    for idx_e, emp in enumerate(employees):
        d1, d2, d3, d4, d5 = st.columns([1, 2.2, 2, 1.2, 1])
        d1.text(emp["codigo"])
        d2.text(emp["nombre"])
        emp["cargo"] = d3.selectbox(
            "Cargo",
            CARGOS,
            index=CARGOS.index(emp["cargo"]) if emp["cargo"] in CARGOS else 0,
            key=f"ecargo_{emp['id']}_{idx_e}",
            label_visibility="collapsed",
        )
        emp["puede_at"] = d4.checkbox("Habilitado AT", value=emp["puede_at"], key=f"eat_{emp['id']}_{idx_e}")
        emp["activo"] = d5.checkbox("Activo", value=emp["activo"], key=f"eact_{emp['id']}_{idx_e}")

with tab5:
    st.subheader("Importar Semana Base desde Excel o Pegado Directo")
    st.caption("Compatible con columnas: CODIGO, NOMBRE, CARGO y los 7 días (LUN a DOM).")

    c_imp1, c_imp2 = st.columns([1.4, 2])
    with c_imp1:
        imp_date = st.date_input("Lunes de inicio de la semana importada", value=active_week["start_date"] + dt.timedelta(days=7))
    with c_imp2:
        imp_mode = st.radio(
            "¿Cómo cargar los datos?",
            ["Crear como nueva semana", "Reemplazar la semana activa"],
            horizontal=True,
        )

    col_up, col_paste = st.columns(2)
    with col_up:
        with st.container(border=True):
            st.markdown("**Opción 1: Subir archivo Excel (.xlsx)**")
            up = st.file_uploader("Sube un archivo Excel con CODIGO, NOMBRE, CARGO y 7 días", type=["xlsx"])
            if up is not None and st.button("Cargar archivo Excel", key="btn_imp_xlsx", type="primary", use_container_width=True):
                try:
                    wb_in = load_workbook(up, data_only=True)
                    ws_in = wb_in.active
                    matrix = [list(r) for r in ws_in.iter_rows(values_only=True)]
                    new_w, cnt = import_matrix_to_week(matrix, imp_date, employees, weeks)
                    if imp_mode == "Reemplazar la semana activa":
                        st.session_state.weeks[week_idx] = new_w
                    else:
                        st.session_state.weeks.append(new_w)
                        st.session_state.active_week_idx = len(st.session_state.weeks) - 1
                    st.session_state.roster_toast = f"Semana '{new_w['label']}' cargada desde Excel ({cnt} colaboradores)."
                    st.rerun()
                except Exception as exc:
                    st.error(f"No se pudo leer el archivo: {exc}")

    with col_paste:
        with st.container(border=True):
            st.markdown("**Opción 2: Copiar y pegar desde Excel (Ctrl + C / Ctrl + V)**")
            paste_txt = st.text_area(
                "Pega aquí las filas copiadas desde Excel (CODIGO, NOMBRE, CARGO y LUN a DOM)",
                height=170,
                placeholder="Selecciona en Excel desde CODIGO hasta DOMINGO, copia con Ctrl + C y pégalo aquí...",
                key="paste_excel_box",
            )
            if st.button("Cargar tabla pegada", key="btn_imp_paste", type="primary", use_container_width=True):
                if not paste_txt.strip():
                    st.error("Pega primero las celdas copiadas de Excel.")
                else:
                    try:
                        lines = [l for l in paste_txt.splitlines() if l.strip()]
                        sep = "\t" if any("\t" in l for l in lines) else ";"
                        matrix = [l.split(sep) for l in lines]
                        new_w, cnt = import_matrix_to_week(matrix, imp_date, employees, weeks)
                        if imp_mode == "Reemplazar la semana activa":
                            st.session_state.weeks[week_idx] = new_w
                        else:
                            st.session_state.weeks.append(new_w)
                            st.session_state.active_week_idx = len(st.session_state.weeks) - 1
                        st.session_state.roster_toast = f"Semana '{new_w['label']}' cargada desde texto pegado ({cnt} colaboradores)."
                        st.rerun()
                    except Exception as exc:
                        st.error(f"Error al leer la tabla pegada: {exc}")
