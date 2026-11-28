from __future__ import annotations

from io import BytesIO
from pathlib import Path
import re
import unicodedata
import numpy as np
import pandas as pd

EXCEL_FILENAME = "262701_DQDLC_A_v003.xlsx"
TOTAL_ROUNDS = 17
NUM_MANAGERS = 10

# Índices de filas y columnas en "Datos Originales" (0-based)
ROW_PTS_START = 5
ROW_CAP_BASE_START = 19
ROW_GOALS_START = 32
ROW_ASSISTS_START = 44
ROW_PAC_START = 56
ROW_MULT_START = 69
COL_NAME = 0
COL_ROUNDS_START = 8
COL_ROUNDS_END = 25  # slice [8:25] -> 17 jornadas

# Configuración de la hoja "Capitanes" (0-based)
CAP_ROW_START = 16
CAP_ROW_END = 35
CAP_COLS = {
    "Veces x3": 2,
    "Veces x2": 3,
    "Veces x1,5": 4,
    "Mult. Promedio": 5,
    "Cap. Top (≥5)": 7,
    "% Aporte Cap.": 8,
}

MIN_REQUIRED_ROWS = ROW_MULT_START + NUM_MANAGERS
MIN_REQUIRED_COLS = COL_ROUNDS_END


def strip_accents(text: str) -> str:
    return "".join(
        ch for ch in unicodedata.normalize("NFKD", str(text))
        if not unicodedata.combining(ch)
    )


def norm_key(text: str) -> str:
    s = strip_accents(str(text))
    return " ".join(s.strip().lower().split())


def find_sheet_name(sheet_names: list[str], target: str) -> str | None:
    target_norm = norm_key(target)
    for name in sheet_names:
        if name.strip().casefold() == target.strip().casefold():
            return name
    for name in sheet_names:
        if norm_key(name) == target_norm:
            return name
    return None


def find_excel_file(filename: str = EXCEL_FILENAME, base_dir: Path | None = None) -> Path | None:
    """Busca el archivo Excel priorizando el nombre exacto o archivos de la liga."""
    root = base_dir or Path(__file__).resolve().parent
    search_dirs = [
        root,
        root.parent,
        root.parent.parent,
        Path.cwd(),
    ]
    # Eliminar duplicados manteniendo orden
    seen: set[Path] = set()
    valid_dirs: list[Path] = []
    for d in search_dirs:
        try:
            resolved = d.resolve()
        except OSError:
            continue
        if resolved not in seen and resolved.exists() and resolved.is_dir():
            seen.add(resolved)
            valid_dirs.append(resolved)

    for folder in valid_dirs:
        exact = folder / filename
        if exact.is_file():
            return exact

    league_candidates: list[Path] = []
    other_candidates: list[Path] = []

    for folder in valid_dirs:
        for p in folder.rglob("*.xlsx"):
            if p.name.startswith("~$") or not p.is_file():
                continue
            name_lower = p.name.lower()
            if "dqdlc" in name_lower or "donqui" in name_lower:
                league_candidates.append(p)
            else:
                other_candidates.append(p)

    if league_candidates:
        return max(league_candidates, key=lambda p: p.stat().st_mtime)

    if other_candidates:
        return max(other_candidates, key=lambda p: p.stat().st_mtime)

    return None


def is_valid_number(val) -> bool:
    if pd.isna(val):
        return False
    s = str(val).strip()
    return s != "" and s.lower() not in {"nan", "none", "null", "n/a", "-", "—"}


def safe_float(val, default: float = 0.0) -> float:
    if not is_valid_number(val):
        return default
    if isinstance(val, (int, float, np.number)):
        f = float(val)
        return default if (np.isnan(f) or np.isinf(f)) else f
    try:
        cleaned = str(val).strip().replace(",", ".")
        f = float(cleaned)
        return default if (np.isnan(f) or np.isinf(f)) else f
    except (ValueError, TypeError):
        return default


def safe_int(val, default: int = 0) -> int:
    return int(round(safe_float(val, default=float(default))))


def round_num_str(rk: str) -> str:
    digits = re.sub(r"\D", "", str(rk))
    return digits or str(rk)


def parse_round_header(text: str) -> tuple[str, str] | None:
    """Detecta encabezados tipo 'JORNADA 13', 'Jornada 14' o 'J13'."""
    s = strip_accents(str(text)).strip()
    m = re.fullmatch(r"(?:jornada|j)\s*(\d{1,2})", s, flags=re.IGNORECASE)
    if m:
        num = int(m.group(1))
        return f"J{num}", f"JORNADA {num}"
    return None


def split_summary_and_fact(text: str) -> tuple[str, str]:
    """Separa el resumen principal del 'Dato curioso:' si existe."""
    parts = re.split(r"\bDato curioso:\s*", text.strip(), maxsplit=1, flags=re.IGNORECASE)
    main_txt = parts[0].strip().rstrip(" .") + "." if parts[0].strip() else ""
    fact_txt = parts[1].strip() if len(parts) > 1 else ""
    return main_txt, fact_txt


def clean_lead_name(manager_name: str, text: str) -> str:
    """Evita repetir '**Nombre:** Nombre (...)' si el texto ya comienza con el nombre."""
    if re.match(rf"^\s*{re.escape(manager_name)}\s*:", text, flags=re.IGNORECASE):
        return re.sub(
            rf"^\s*{re.escape(manager_name)}\s*:\s*",
            "",
            text,
            count=1,
            flags=re.IGNORECASE,
        ).strip()
    return text.strip()


def extract_managers_summaries(
    df_mgr_sheet: pd.DataFrame,
    fallback_round: int,
    known_managers: list[str] | None = None,
) -> dict[str, dict[str, str]]:
    """
    Lee resúmenes por jornada y mánager desde la hoja 'Mánagers' del Excel,
    soportando tanto una celda en columna A ('Nombre: texto...') como
    nombre en columna A y texto en columna B+.
    """
    summaries_by_round: dict[str, dict[str, str]] = {}
    current_rk = f"J{max(1, fallback_round)}"
    known_map = {norm_key(str(m)): str(m) for m in (known_managers or []) if str(m).strip()}

    for r_i in range(len(df_mgr_sheet)):
        row_vals = [
            str(df_mgr_sheet.iloc[r_i, c_idx]).strip()
            for c_idx in range(df_mgr_sheet.shape[1])
            if pd.notna(df_mgr_sheet.iloc[r_i, c_idx]) and str(df_mgr_sheet.iloc[r_i, c_idx]).strip()
        ]
        if not row_vals:
            continue

        first_cell = row_vals[0]
        hdr = parse_round_header(first_cell)
        if hdr:
            current_rk = hdr[0]
            summaries_by_round.setdefault(current_rk, {})
            continue

        if len(row_vals) >= 2:
            mgr_name = row_vals[0]
            summary_text = " ".join(row_vals[1:]).strip()
            canon = known_map.get(norm_key(mgr_name), mgr_name)
            if canon and summary_text:
                summaries_by_round.setdefault(current_rk, {})[canon] = summary_text
            continue

        m_colon = re.match(r"^([^:]+):\s*(.+)$", first_cell, flags=re.DOTALL)
        if m_colon:
            cand_raw = m_colon.group(1).strip()
            canon = known_map.get(norm_key(cand_raw), cand_raw)
            body = m_colon.group(2).strip()
            if canon and body:
                summaries_by_round.setdefault(current_rk, {})[canon] = body
                continue

        first_norm = norm_key(first_cell)
        for k_norm, canon in known_map.items():
            if first_norm.startswith(k_norm + " "):
                summaries_by_round.setdefault(current_rk, {})[canon] = first_cell
                break

    return summaries_by_round


def _extract_block_df(
    df_orig: pd.DataFrame,
    row_start: int,
    managers: list[str],
    round_cols: list[str],
    as_int: bool = False,
) -> pd.DataFrame:
    """Extrae un bloque de (NUM_MANAGERS x TOTAL_ROUNDS) de forma estructurada."""
    raw_block = df_orig.iloc[row_start : row_start + NUM_MANAGERS, COL_ROUNDS_START:COL_ROUNDS_END].copy()
    raw_block.index = managers
    raw_block.columns = round_cols

    parsed = pd.DataFrame(index=managers, columns=round_cols, dtype=float)
    for col in round_cols:
        parsed[col] = raw_block[col].apply(
            lambda v: float(safe_int(v)) if (is_valid_number(v) and as_int)
            else (safe_float(v) if is_valid_number(v) else np.nan)
        )
    return parsed


def load_raw_excel_data(file_bytes: bytes) -> dict:
    """
    Lee y valida el Excel de la liga en una única apertura de pd.ExcelFile.
    Devuelve matrices limpias por jornada (index=mánager, columns=J1..J17),
    datos extra de capitanes y resúmenes editoriales.
    """
    bio = BytesIO(file_bytes)
    xls = pd.ExcelFile(bio)
    sheet_names = xls.sheet_names

    orig_sheet = find_sheet_name(sheet_names, "Datos Originales")
    if orig_sheet is None:
        raise ValueError(
            f"No se encontró la hoja 'Datos Originales' en el Excel. Hojas disponibles: {sheet_names}"
        )

    df_orig = xls.parse(sheet_name=orig_sheet, header=None)
    if df_orig.shape[0] < MIN_REQUIRED_ROWS or df_orig.shape[1] < MIN_REQUIRED_COLS:
        raise ValueError(
            f"Estructura de '{orig_sheet}' inválida: se esperaban al menos "
            f"{MIN_REQUIRED_ROWS} filas y {MIN_REQUIRED_COLS} columnas, "
            f"pero tiene {df_orig.shape[0]}x{df_orig.shape[1]}."
        )

    managers: list[str] = []
    for idx in range(NUM_MANAGERS):
        raw_name = df_orig.iloc[ROW_PTS_START + idx, COL_NAME]
        name = (
            str(raw_name).strip()
            if pd.notna(raw_name) and str(raw_name).strip()
            else f"Mánager {idx + 1}"
        )
        managers.append(name)

    all_round_cols = [f"J{i + 1}" for i in range(TOTAL_ROUNDS)]

    df_pts_all = _extract_block_df(df_orig, ROW_PTS_START, managers, all_round_cols, as_int=True)
    df_cap_base_all = _extract_block_df(df_orig, ROW_CAP_BASE_START, managers, all_round_cols, as_int=False)
    df_mult_all = _extract_block_df(df_orig, ROW_MULT_START, managers, all_round_cols, as_int=False)
    df_goals_all = _extract_block_df(df_orig, ROW_GOALS_START, managers, all_round_cols, as_int=True)
    df_assists_all = _extract_block_df(df_orig, ROW_ASSISTS_START, managers, all_round_cols, as_int=True)
    df_pac_all = _extract_block_df(df_orig, ROW_PAC_START, managers, all_round_cols, as_int=True)

    # Filtrar solo jornadas donde haya al menos un dato registrado
    played_mask = df_pts_all.notna().any(axis=0)
    played_cols = [c for c in all_round_cols if played_mask[c]]

    df_pts = df_pts_all[played_cols]
    df_cap_base = df_cap_base_all[played_cols].fillna(0.0)
    df_mult = df_mult_all[played_cols].fillna(0.0)
    df_cap_pts = (df_cap_base * df_mult).round(0)
    df_goals = df_goals_all[played_cols].fillna(0.0)
    df_assists = df_assists_all[played_cols].fillna(0.0)
    df_ga = df_goals + df_assists
    df_pac = df_pac_all[played_cols].fillna(0.0)

    # Hoja Capitanes (opcional)
    cap_extra: dict[str, dict] = {}
    cap_sheet_name = find_sheet_name(sheet_names, "Capitanes")
    if cap_sheet_name:
        df_cap_sheet = xls.parse(sheet_name=cap_sheet_name, header=None)
        n_cols = df_cap_sheet.shape[1]
        for r_i in range(CAP_ROW_START, min(len(df_cap_sheet), CAP_ROW_END)):
            m_raw = df_cap_sheet.iloc[r_i, COL_NAME]
            if pd.notna(m_raw) and str(m_raw).strip():
                m_key = str(m_raw).strip()
                cap_extra[m_key] = {
                    "Veces x3": safe_int(df_cap_sheet.iloc[r_i, CAP_COLS["Veces x3"]]) if n_cols > CAP_COLS["Veces x3"] else 0,
                    "Veces x2": safe_int(df_cap_sheet.iloc[r_i, CAP_COLS["Veces x2"]]) if n_cols > CAP_COLS["Veces x2"] else 0,
                    "Veces x1,5": safe_int(df_cap_sheet.iloc[r_i, CAP_COLS["Veces x1,5"]]) if n_cols > CAP_COLS["Veces x1,5"] else 0,
                    "Mult. Promedio": round(safe_float(df_cap_sheet.iloc[r_i, CAP_COLS["Mult. Promedio"]]), 2) if n_cols > CAP_COLS["Mult. Promedio"] else 0.0,
                    "Cap. Top (≥5)": safe_int(df_cap_sheet.iloc[r_i, CAP_COLS["Cap. Top (≥5)"]]) if n_cols > CAP_COLS["Cap. Top (≥5)"] else 0,
                    "% Aporte Cap. Excel": round(safe_float(df_cap_sheet.iloc[r_i, CAP_COLS["% Aporte Cap."]]) * 100, 1) if n_cols > CAP_COLS["% Aporte Cap."] else 0.0,
                }

    # Hoja Mánagers (resúmenes editoriales)
    max_j_played = int(df_pts.notna().sum(axis=1).max()) if not df_pts.empty else 1
    summaries_by_round: dict[str, dict[str, str]] = {}
    mgr_sheet_name = find_sheet_name(sheet_names, "Mánagers")
    if mgr_sheet_name:
        df_mgr_sheet = xls.parse(sheet_name=mgr_sheet_name, header=None)
        # FIX: pasamos lista de nombres (list[str]), no dicts
        summaries_by_round = extract_managers_summaries(
            df_mgr_sheet,
            max_j_played,
            known_managers=managers,
        )

    return {
        "managers": managers,
        "played_cols": played_cols,
        "df_pts": df_pts,
        "df_cap_base": df_cap_base,
        "df_mult": df_mult,
        "df_cap_pts": df_cap_pts,
        "df_goals": df_goals,
        "df_assists": df_assists,
        "df_ga": df_ga,
        "df_pac": df_pac,
        "cap_extra": cap_extra,
        "summaries_by_round": summaries_by_round,
    }
