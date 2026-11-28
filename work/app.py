from io import BytesIO
from pathlib import Path
import re
import unicodedata
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(
    page_title="Donqui League A — Analytics",
    page_icon="⚽",
    layout="wide",
    initial_sidebar_state="expanded",
)

EXCEL_FILENAME = "262701_DQDLC_A_v003.xlsx"
TOTAL_ROUNDS = 17
NUM_MANAGERS = 10

# Índices de filas en la hoja "Datos Originales" (0-based)
ROW_PTS_START = 5
ROW_CAP_BASE_START = 19
ROW_GOALS_START = 32
ROW_ASSISTS_START = 44
ROW_PAC_START = 56
ROW_MULT_START = 69
COL_ROUNDS_START = 8
COL_ROUNDS_END = 25  # slice [8:25] -> 17 jornadas

APP_DIR = Path(__file__).resolve().parent


def _strip_accents(text: str) -> str:
    return "".join(
        ch for ch in unicodedata.normalize("NFKD", str(text))
        if not unicodedata.combining(ch)
    )


def _norm_key(text: str) -> str:
    s = _strip_accents(str(text))
    return " ".join(s.strip().lower().split())


def _find_sheet_name(sheet_names: list[str], target: str) -> str | None:
    target_norm = _norm_key(target)
    for name in sheet_names:
        if name.strip().casefold() == target.strip().casefold():
            return name
    for name in sheet_names:
        if _norm_key(name) == target_norm:
            return name
    return None


def find_excel_file(filename: str) -> Path | None:
    """Busca el archivo Excel priorizando el nombre exacto o archivos de la liga."""
    search_dirs = [
        APP_DIR,
        APP_DIR.parent,
        APP_DIR.parent.parent,
        Path.cwd(),
    ]

    valid_dirs = [d for d in search_dirs if d.exists() and d.is_dir()]

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


def _is_valid_number(val) -> bool:
    if pd.isna(val):
        return False
    s = str(val).strip()
    return s != "" and s.lower() not in {"nan", "none", "null", "n/a", "-", "—"}


def _safe_float(val, default: float = 0.0) -> float:
    if not _is_valid_number(val):
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


def _safe_int(val, default: int = 0) -> int:
    f = _safe_float(val, default=float(default))
    return int(round(f))


def _round_num_str(rk: str) -> str:
    digits = re.sub(r"\D", "", str(rk))
    return digits or str(rk)


def _parse_round_header(text: str) -> tuple[str, str] | None:
    """Detecta encabezados tipo 'JORNADA 13', 'Jornada 14' o 'J13'."""
    s = _strip_accents(str(text)).strip()
    m = re.fullmatch(r"(?:jornada|j)\s*(\d{1,2})", s, flags=re.IGNORECASE)
    if m:
        num = int(m.group(1))
        return f"J{num}", f"JORNADA {num}"
    return None


def _extract_managers_summaries(
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
    known_map = {_norm_key(m): m for m in (known_managers or [])}

    for r_i in range(len(df_mgr_sheet)):
        row_vals = [
            str(df_mgr_sheet.iloc[r_i, c_idx]).strip()
            for c_idx in range(df_mgr_sheet.shape[1])
            if pd.notna(df_mgr_sheet.iloc[r_i, c_idx]) and str(df_mgr_sheet.iloc[r_i, c_idx]).strip()
        ]
        if not row_vals:
            continue

        first_cell = row_vals[0]
        hdr = _parse_round_header(first_cell)
        if hdr:
            current_rk = hdr[0]
            summaries_by_round.setdefault(current_rk, {})
            continue

        if len(row_vals) >= 2:
            mgr_name = row_vals[0]
            summary_text = " ".join(row_vals[1:]).strip()
            canon = known_map.get(_norm_key(mgr_name), mgr_name)
            if canon and summary_text:
                summaries_by_round.setdefault(current_rk, {})[canon] = summary_text
            continue

        m_colon = re.match(r"^([^:]+):\s*(.+)$", first_cell, flags=re.DOTALL)
        if m_colon:
            cand_raw = m_colon.group(1).strip()
            canon = known_map.get(_norm_key(cand_raw), cand_raw)
            body = m_colon.group(2).strip()
            if canon and body:
                summaries_by_round.setdefault(current_rk, {})[canon] = body
                continue

        for k_norm, canon in known_map.items():
            if _norm_key(first_cell).startswith(k_norm + " "):
                summaries_by_round.setdefault(current_rk, {})[canon] = first_cell
                break

    return summaries_by_round


def _split_summary_and_fact(text: str) -> tuple[str, str]:
    """Separa el resumen principal del 'Dato curioso:' si existe."""
    parts = re.split(r"\bDato curioso:\s*", text.strip(), maxsplit=1, flags=re.IGNORECASE)
    main_txt = parts[0].strip().rstrip(" .") + "." if parts[0].strip() else ""
    fact_txt = parts[1].strip() if len(parts) > 1 else ""
    return main_txt, fact_txt


def _clean_lead_name(manager_name: str, text: str) -> str:
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


def _apply_clean_layout(
    fig: go.Figure,
    height: int = 400,
    title: str | None = None,
    top_margin: int = 55,
    right_margin: int = 25,
) -> go.Figure:
    """Aplica márgenes, tamaños de fuente y alineación homogéneos a los gráficos."""
    if title is not None:
        fig.update_layout(title=dict(text=title, font=dict(size=15)))
    fig.update_layout(
        height=height,
        margin=dict(l=20, r=right_margin, t=top_margin, b=20),
        font=dict(size=12),
    )
    return fig


@st.cache_data(show_spinner="Cargando datos de Donqui League...")
def load_league_data(source_key: str, file_bytes: bytes):
    bio = BytesIO(file_bytes)
    xls = pd.ExcelFile(bio)
    sheet_names = xls.sheet_names

    orig_sheet = _find_sheet_name(sheet_names, "Datos Originales") or "Datos Originales"
    df_orig = pd.read_excel(bio, sheet_name=orig_sheet, header=None)

    cap_extra: dict[str, dict] = {}
    cap_sheet_name = _find_sheet_name(sheet_names, "Capitanes")
    if cap_sheet_name:
        bio.seek(0)
        df_cap_sheet = pd.read_excel(bio, sheet_name=cap_sheet_name, header=None)
        for r_i in range(16, min(len(df_cap_sheet), 35)):
            m_raw = df_cap_sheet.iloc[r_i, 0]
            if pd.notna(m_raw) and str(m_raw).strip():
                m_key = str(m_raw).strip()
                n_cols = df_cap_sheet.shape[1]
                cap_extra[m_key] = {
                    "Veces x3": _safe_int(df_cap_sheet.iloc[r_i, 2]) if n_cols > 2 else 0,
                    "Veces x2": _safe_int(df_cap_sheet.iloc[r_i, 3]) if n_cols > 3 else 0,
                    "Veces x1,5": _safe_int(df_cap_sheet.iloc[r_i, 4]) if n_cols > 4 else 0,
                    "Mult. Promedio": round(_safe_float(df_cap_sheet.iloc[r_i, 5]), 2) if n_cols > 5 else 0.0,
                    "Cap. Top (≥5)": _safe_int(df_cap_sheet.iloc[r_i, 7]) if n_cols > 7 else 0,
                    "% Aporte Cap.": round(_safe_float(df_cap_sheet.iloc[r_i, 8]) * 100, 1) if n_cols > 8 else 0.0,
                }

    managers = []
    pts_matrix = {}
    cap_matrix = {}
    ga_matrix = {}
    pac_matrix = {}

    for idx in range(NUM_MANAGERS):
        raw_name = df_orig.iloc[ROW_PTS_START + idx, 0]
        name = (
            str(raw_name).strip()
            if pd.notna(raw_name) and str(raw_name).strip()
            else f"Mánager {idx + 1}"
        )
        raw_pts = df_orig.iloc[ROW_PTS_START + idx, COL_ROUNDS_START:COL_ROUNDS_END].tolist()

        played_pts: list[float] = []
        round_labels: list[str] = []
        for j_i, val in enumerate(raw_pts):
            if _is_valid_number(val):
                played_pts.append(_safe_float(val))
                round_labels.append(f"J{j_i + 1}")

        n_j = len(played_pts)
        rem_j = max(0, TOTAL_ROUNDS - n_j)

        cap_b = [
            _safe_float(df_orig.iloc[ROW_CAP_BASE_START + idx, COL_ROUNDS_START + j])
            for j in range(n_j)
        ]
        mult = [
            _safe_float(df_orig.iloc[ROW_MULT_START + idx, COL_ROUNDS_START + j])
            for j in range(n_j)
        ]
        goals = [
            _safe_float(df_orig.iloc[ROW_GOALS_START + idx, COL_ROUNDS_START + j])
            for j in range(n_j)
        ]
        assists = [
            _safe_float(df_orig.iloc[ROW_ASSISTS_START + idx, COL_ROUNDS_START + j])
            for j in range(n_j)
        ]
        pac = [
            _safe_float(df_orig.iloc[ROW_PAC_START + idx, COL_ROUNDS_START + j])
            for j in range(n_j)
        ]

        tot = sum(played_pts)
        cap_round_pts = [int(round(b * m)) for b, m in zip(cap_b, mult)]
        cap_pts = sum(cap_round_pts)
        net_bonus = int(round(cap_pts - sum(cap_b)))
        no_cap = int(round(tot - net_bonus))

        avg = float(np.mean(played_pts)) if n_j else 0.0
        med = float(np.median(played_pts)) if n_j else 0.0
        sd = float(np.std(played_pts, ddof=1)) if n_j > 1 else 0.0
        cv = (sd / avg) if avg else 0.0
        last3 = float(np.mean(played_pts[-3:])) if n_j >= 3 else avg
        last_j = int(round(played_pts[-1])) if n_j else 0

        ranks: list[int] = []
        for j in range(n_j):
            col_vals = [
                _safe_float(df_orig.iloc[ROW_PTS_START + k, COL_ROUNDS_START + j])
                for k in range(NUM_MANAGERS)
                if _is_valid_number(df_orig.iloc[ROW_PTS_START + k, COL_ROUNDS_START + j])
            ]
            if not col_vals:
                continue
            s_ranks = pd.Series(col_vals, dtype="float64").rank(ascending=False, method="min")
            match_idx = next(
                (i_v for i_v, v_c in enumerate(col_vals) if np.isclose(v_c, played_pts[j])),
                None,
            )
            if match_idx is not None:
                ranks.append(int(s_ranks.iloc[match_idx]))

        ga_tot = int(round(sum(goals) + sum(assists)))
        extra = cap_extra.get(
            name,
            {
                "Veces x3": 0,
                "Veces x2": 0,
                "Veces x1,5": 0,
                "Mult. Promedio": 0.0,
                "Cap. Top (≥5)": 0,
                "% Aporte Cap.": round((cap_pts / tot * 100), 1) if tot else 0.0,
            },
        )

        managers.append(
            {
                "Mánager": name,
                "Jornadas": n_j,
                "Puntos Totales": int(round(tot)),
                "Últ. J": last_j,
                "Media / J": round(avg, 1),
                "Mediana / J": round(med, 1),
                "DE": round(sd, 1),
                "CV": round(cv, 2),
                "Últ. 3J": round(last3, 1),
                "Bonus Cap.": net_bonus,
                "Sin Cap.": no_cap,
                "Goles": int(round(sum(goals))),
                "Asistencias": int(round(sum(assists))),
                "G+A": ga_tot,
                "Pts / (G+A)": round(tot / ga_tot, 1) if ga_tot else 0.0,
                "PaC": int(round(sum(pac))),
                "Pos. Media": round(float(np.mean(ranks)), 1) if ranks else 0.0,
                "Victorias J": sum(1 for r in ranks if r == 1),
                "Top 3 J": sum(1 for r in ranks if r <= 3),
                "Proy. Media J17": int(round(tot + avg * rem_j)),
                "Proy. Forma J17": int(round(tot + last3 * rem_j)),
                "Mínima": int(round(min(played_pts))) if played_pts else 0,
                "Máxima": int(round(max(played_pts))) if played_pts else 0,
                **extra,
            }
        )

        pts_matrix[name] = dict(zip(round_labels, [int(round(x)) for x in played_pts]))
        cap_matrix[name] = dict(zip(round_labels, cap_round_pts))
        ga_matrix[name] = dict(zip(round_labels, [int(round(g + a)) for g, a in zip(goals, assists)]))
        pac_matrix[name] = dict(zip(round_labels, [int(round(p)) for p in pac]))

    df_cls = (
        pd.DataFrame(managers)
        .sort_values(["Puntos Totales", "Media / J"], ascending=False)
        .reset_index(drop=True)
    )
    df_cls.insert(0, "Pos", df_cls.index + 1)
    leader_pts = int(df_cls["Puntos Totales"].iloc[0]) if not df_cls.empty else 0
    df_cls.insert(3, "Dif. Líder", df_cls["Puntos Totales"] - leader_pts)

    all_round_cols = [f"J{i + 1}" for i in range(TOTAL_ROUNDS)]
    present_cols = [c for c in all_round_cols if any(c in pts_matrix.get(m, {}) for m in df_cls["Mánager"])]

    df_pts_j = pd.DataFrame.from_dict(pts_matrix, orient="index").reindex(index=df_cls["Mánager"], columns=present_cols)
    df_cap_j = pd.DataFrame.from_dict(cap_matrix, orient="index").reindex(index=df_cls["Mánager"], columns=present_cols)
    df_ga_j = pd.DataFrame.from_dict(ga_matrix, orient="index").reindex(index=df_cls["Mánager"], columns=present_cols)
    df_pac_j = pd.DataFrame.from_dict(pac_matrix, orient="index").reindex(index=df_cls["Mánager"], columns=present_cols)

    max_j_played = int(df_cls["Jornadas"].max()) if not df_cls.empty else 1
    summaries_by_round: dict[str, dict[str, str]] = {}

    mgr_sheet_name = _find_sheet_name(sheet_names, "Mánagers")
    if mgr_sheet_name:
        bio.seek(0)
        df_mgr_sheet = pd.read_excel(bio, sheet_name=mgr_sheet_name, header=None)
        summaries_by_round = _extract_managers_summaries(
            df_mgr_sheet,
            max_j_played,
            known_managers=managers,
        )

    return df_cls, df_pts_j, df_cap_j, df_ga_j, df_pac_j, summaries_by_round


st.title("⚽ Donqui League — Centro Analítico")
st.caption("Actualizado desde la última versión del Excel de la liga")

col_top_1, col_top_2 = st.columns([6, 1])
with col_top_2:
    if st.button("🔄 Actualizar", width="stretch"):
        st.cache_data.clear()
        st.rerun()

excel_path = find_excel_file(EXCEL_FILENAME)
if excel_path is not None:
    stat = excel_path.stat()
    source_key = f"file::{excel_path.resolve()}::{stat.st_mtime_ns}::{stat.st_size}"
    file_bytes = excel_path.read_bytes()
    st.caption(f"Fuente activa: `{excel_path.name}`")
else:
    st.warning(
        f"No se encontró `{EXCEL_FILENAME}` ni otro archivo `.xlsx` en el repositorio. "
        "Sube el Excel aquí para cargarlo al instante:"
    )
    uploaded_file = st.file_uploader("Sube el Excel (.xlsx)", type=["xlsx"])
    if uploaded_file is None:
        st.stop()
    file_bytes = uploaded_file.getvalue()
    source_key = f"upload::{uploaded_file.name}::{len(file_bytes)}"

df_cls, df_pts_j, df_cap_j, df_ga_j, df_pac_j, summaries_by_round = load_league_data(
    source_key, file_bytes
)

if df_cls.empty or df_pts_j.empty:
    st.info("El archivo se cargó, pero no hay datos de jornadas registrados todavía.")
    st.stop()

current_round = int(df_cls["Jornadas"].max())
rem_rounds = max(0, TOTAL_ROUNDS - current_round)
all_managers = df_cls["Mánager"].tolist()
all_rounds = df_pts_j.columns.tolist()

# --- FILTROS EN SIDEBAR ---
st.sidebar.header("🎛️ Filtros de análisis")

if "mgr_filter" not in st.session_state or set(st.session_state["mgr_filter"]) - set(all_managers):
    st.session_state["mgr_filter"] = all_managers.copy()

b_all, b_none = st.sidebar.columns(2)
if b_all.button("Todos", width="stretch"):
    st.session_state["mgr_filter"] = all_managers.copy()
if b_none.button("Limpiar", width="stretch"):
    st.session_state["mgr_filter"] = []

selected_managers = st.sidebar.multiselect(
    "Mánagers",
    options=all_managers,
    key="mgr_filter",
)
if not selected_managers:
    st.sidebar.caption("Mostrando todos los mánagers por defecto.")
    active_managers = all_managers
else:
    active_managers = selected_managers

st.sidebar.divider()

preset = st.sidebar.radio(
    "Tramo rápido de jornadas",
    options=["Todas", "Últimas 3J", "Últimas 5J", "Rango / Manual"],
    horizontal=True,
    index=0,
)

if preset == "Todas":
    active_rounds = all_rounds
elif preset == "Últimas 3J":
    active_rounds = all_rounds[-3:]
elif preset == "Últimas 5J":
    active_rounds = all_rounds[-5:]
else:
    round_nums = [int(r[1:]) for r in all_rounds]
    if len(round_nums) >= 2:
        r_start, r_end = st.sidebar.slider(
            "Rango de jornadas",
            min_value=min(round_nums),
            max_value=max(round_nums),
            value=(min(round_nums), max(round_nums)),
        )
        range_rounds = [f"J{n}" for n in range(r_start, r_end + 1) if f"J{n}" in all_rounds]
    else:
        range_rounds = all_rounds

    active_rounds = st.sidebar.multiselect(
        "Ajuste fino de jornadas",
        options=all_rounds,
        default=range_rounds,
    )
    if not active_rounds:
        active_rounds = range_rounds

st.sidebar.caption(f"Jornadas activas: **{', '.join(active_rounds)}**")

# --- RESUMEN SUPERIOR (KPIs) ---
leader = df_cls.iloc[0]
top_ga = df_cls.sort_values("G+A", ascending=False).iloc[0]
most_consistent = df_cls.sort_values("CV", ascending=True).iloc[0]
hottest = df_cls.sort_values("Últ. 3J", ascending=False).iloc[0]

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Jornada actual", f"J{current_round} / J{TOTAL_ROUNDS}", f"{rem_rounds} restantes")
k2.metric("Líder", leader["Mánager"], f"{leader['Puntos Totales']} pts")
k3.metric("Mejor racha (3J)", hottest["Mánager"], f"{hottest['Últ. 3J']} media")
k4.metric("Más regular (CV)", most_consistent["Mánager"], f"CV {most_consistent['CV']}")
k5.metric("Top G+A", top_ga["Mánager"], f"{top_ga['G+A']} contribuciones")

# --- PESTAÑAS ---
tab_vis, tab_evo, tab_reg, tab_proj, tab_mgr, tab_res = st.tabs(
    [
        "🏆 Visión General",
        "📈 Evolución por Jornada",
        "📊 Regularidad y Capitán",
        "🎯 Proyección J17",
        "🕵️ Ficha por Mánager",
        "📝 Resúmenes por Jornada",
    ]
)

# 1. VISIÓN GENERAL
with tab_vis:
    overview_cols = [
        "Pos",
        "Mánager",
        "Dif. Líder",
        "Puntos Totales",
        "Goles",
        "Asistencias",
        "PaC",
        "Pos. Media",
        "Máxima",
        "Mínima",
    ]
    view_vis = df_cls[df_cls["Mánager"].isin(active_managers)][overview_cols]

    st.dataframe(
        view_vis,
        width="stretch",
        hide_index=True,
        column_config={
            "Puntos Totales": st.column_config.ProgressColumn(
                "Puntos Totales",
                min_value=0,
                max_value=max(int(df_cls["Puntos Totales"].max()), 1),
                format="%d",
            ),
            "PaC": st.column_config.NumberColumn("PaC", help="Porterías a cero"),
        },
    )

    c_left, c_right = st.columns(2)

    with c_left:
        bar_df = view_vis.sort_values("Puntos Totales", ascending=True)
        fig_bar = px.bar(
            bar_df,
            x="Puntos Totales",
            y="Mánager",
            orientation="h",
            text="Puntos Totales",
            color="Puntos Totales",
            color_continuous_scale="Tealgrn",
            title=f"Puntos Totales (J1–J{current_round})",
        )
        fig_bar.update_traces(
            textposition="outside",
            textfont=dict(size=12),
            cliponaxis=False,
        )
        _apply_clean_layout(
            fig_bar,
            height=410,
            title=f"Puntos Totales (J1–J{current_round})",
            top_margin=55,
            right_margin=45,
        )
        fig_bar.update_layout(
            coloraxis_showscale=False,
            xaxis=dict(title="Puntos Totales", tickfont=dict(size=12), rangemode="tozero"),
            yaxis=dict(title="", tickfont=dict(size=12)),
        )
        st.plotly_chart(fig_bar, width="stretch")

    with c_right:
        range_df = view_vis.sort_values("Máxima", ascending=True)
        fig_mm = go.Figure()
        fig_mm.add_trace(
            go.Bar(
                name="Máxima",
                y=range_df["Mánager"],
                x=range_df["Máxima"],
                orientation="h",
                text=range_df["Máxima"],
                textposition="outside",
                textfont=dict(size=12),
                cliponaxis=False,
            )
        )
        fig_mm.add_trace(
            go.Bar(
                name="Mínima",
                y=range_df["Mánager"],
                x=range_df["Mínima"],
                orientation="h",
                text=range_df["Mínima"],
                textposition="outside",
                textfont=dict(size=12),
                cliponaxis=False,
            )
        )
        _apply_clean_layout(
            fig_mm,
            height=410,
            title="Puntuación Máxima vs. Mínima por Mánager",
            top_margin=70,
            right_margin=45,
        )
        fig_mm.update_layout(
            barmode="group",
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="left",
                x=0,
                font=dict(size=12),
            ),
            xaxis=dict(title="Puntos", tickfont=dict(size=12), rangemode="tozero"),
            yaxis=dict(title="", tickfont=dict(size=12)),
        )
        st.plotly_chart(fig_mm, width="stretch")

# 2. EVOLUCIÓN POR JORNADA (SEGUIMIENTO)
with tab_evo:
    st.subheader("Seguimiento por jornada")
    st.caption("Comparativa alineada de puntuación y rendimiento jornada a jornada.")

    sub_pts = df_pts_j.loc[active_managers, active_rounds].copy()
    sub_pts = sub_pts.dropna(how="all")

    if sub_pts.empty:
        st.info("No hay datos de seguimiento para los filtros seleccionados.")
    else:
        df_melt = (
            sub_pts.reset_index()
            .melt(id_vars="Mánager", var_name="Jornada", value_name="Puntos")
            .dropna(subset=["Puntos"])
        )
        df_melt["Puntos"] = df_melt["Puntos"].round(0).astype(int)

        fig_line = px.line(
            df_melt,
            x="Jornada",
            y="Puntos",
            color="Mánager",
            category_orders={"Jornada": active_rounds},
            markers=True,
            title="Evolución de puntos por jornada",
        )
        fig_line.update_traces(
            marker=dict(size=7, line=dict(width=1, color="#111827")),
            line=dict(width=2.2),
        )
        _apply_clean_layout(
            fig_line,
            height=430,
            title="Evolución de puntos por jornada",
            top_margin=70,
            right_margin=25,
        )
        fig_line.update_layout(
            hovermode="x unified",
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="left",
                x=0,
                font=dict(size=12),
            ),
            xaxis=dict(title="Jornada", tickfont=dict(size=12)),
            yaxis=dict(title="Puntos", tickfont=dict(size=12), rangemode="tozero"),
        )
        st.plotly_chart(fig_line, width="stretch")

        pts_min = int(sub_pts.min().min()) if not sub_pts.isna().all().all() else 0
        pts_max = int(sub_pts.max().max()) if not sub_pts.isna().all().all() else 100

        fig_heat = px.imshow(
            sub_pts,
            text_auto=".0f",
            aspect="auto",
            color_continuous_scale="Tealgrn",
            zmin=pts_min,
            zmax=pts_max,
            title="Matriz de puntos por jornada",
        )
        fig_heat.update_traces(textfont=dict(size=12))
        _apply_clean_layout(
            fig_heat,
            height=380,
            title="Matriz de puntos por jornada",
            top_margin=55,
            right_margin=25,
        )
        fig_heat.update_layout(
            coloraxis_colorbar=dict(title="Pts", thickness=14, len=0.85),
            xaxis=dict(title="Jornada", tickfont=dict(size=12)),
            yaxis=dict(title="", tickfont=dict(size=12)),
        )
        st.plotly_chart(fig_heat, width="stretch")

        with st.expander("Ver tabla numérica de seguimiento"):
            num_cols = {col: st.column_config.NumberColumn(col, format="%d") for col in active_rounds}
            st.dataframe(
                sub_pts.reset_index(),
                width="stretch",
                hide_index=True,
                column_config=num_cols,
            )

# 3. REGULARIDAD Y CAPITÁN
with tab_reg:
    deep_cols = [
        "Pos",
        "Mánager",
        "Media / J",
        "Mediana / J",
        "DE",
        "CV",
        "Últ. 3J",
        "Bonus Cap.",
        "Sin Cap.",
        "Mult. Promedio",
        "Veces x3",
        "Cap. Top (≥5)",
        "Pts / (G+A)",
        "Victorias J",
        "Top 3 J",
    ]
    view_deep = df_cls[df_cls["Mánager"].isin(active_managers)][deep_cols]
    st.dataframe(view_deep, width="stretch", hide_index=True)

    r_left, r_right = st.columns(2)
    with r_left:
        fig_scatter = px.scatter(
            view_deep,
            x="Media / J",
            y="CV",
            size="Top 3 J",
            color="Mánager",
            text="Mánager",
            title="Regularidad (CV) vs. Media por Jornada",
            labels={"CV": "Coeficiente de Variación (menor = más regular)"},
        )
        fig_scatter.update_traces(
            textposition="top center",
            textfont=dict(size=11),
        )
        _apply_clean_layout(
            fig_scatter,
            height=410,
            title="Regularidad (CV) vs. Media por Jornada",
            top_margin=55,
            right_margin=25,
        )
        fig_scatter.update_layout(
            showlegend=False,
            xaxis=dict(title="Media / J", tickfont=dict(size=12)),
            yaxis=dict(title="CV (menor = más regular)", tickfont=dict(size=12), rangemode="tozero"),
        )
        st.plotly_chart(fig_scatter, width="stretch")

    with r_right:
        cap_df = view_deep.sort_values("Bonus Cap.", ascending=True)
        fig_cap = px.bar(
            cap_df,
            x="Bonus Cap.",
            y="Mánager",
            orientation="h",
            text="Bonus Cap.",
            title="Bonus Neto de Capitán por Mánager",
        )
        fig_cap.update_traces(
            textposition="outside",
            textfont=dict(size=12),
            cliponaxis=False,
        )
        _apply_clean_layout(
            fig_cap,
            height=410,
            title="Bonus Neto de Capitán por Mánager",
            top_margin=55,
            right_margin=45,
        )
        fig_cap.update_layout(
            xaxis=dict(title="Bonus Cap.", tickfont=dict(size=12)),
            yaxis=dict(title="", tickfont=dict(size=12)),
        )
        st.plotly_chart(fig_cap, width="stretch")

# 4. PROYECCIÓN J17
with tab_proj:
    if rem_rounds == 0:
        st.info(f"Se han completado las {TOTAL_ROUNDS} jornadas; las proyecciones coinciden con los puntos finales.")

    view_proj = df_cls[df_cls["Mánager"].isin(active_managers)].copy()
    proj_long = view_proj.melt(
        id_vars=["Mánager"],
        value_vars=["Puntos Totales", "Proy. Media J17", "Proy. Forma J17"],
        var_name="Escenario",
        value_name="Puntos",
    )
    fig_proj = px.bar(
        proj_long,
        x="Mánager",
        y="Puntos",
        color="Escenario",
        barmode="group",
        text_auto=".0f",
        title="Puntos Actuales vs. Proyección a J17 (Media vs. Últimas 3 Jornadas)",
    )
    fig_proj.update_traces(
        textposition="outside",
        textfont=dict(size=11),
        cliponaxis=False,
    )
    _apply_clean_layout(
        fig_proj,
        height=450,
        title="Puntos Actuales vs. Proyección a J17 (Media vs. Últimas 3 Jornadas)",
        top_margin=70,
        right_margin=25,
    )
    fig_proj.update_layout(
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="left",
            x=0,
            font=dict(size=12),
        ),
        xaxis=dict(title="", tickfont=dict(size=12), tickangle=-20),
        yaxis=dict(title="Puntos", tickfont=dict(size=12), rangemode="tozero"),
    )
    st.plotly_chart(fig_proj, width="stretch")

    st.dataframe(
        view_proj[
            ["Pos", "Mánager", "Puntos Totales", "Media / J", "Últ. 3J", "Proy. Media J17", "Proy. Forma J17"]
        ],
        width="stretch",
        hide_index=True,
    )

# 5. FICHA POR MÁNAGER
with tab_mgr:
    chosen = st.selectbox("Elige un mánager", options=all_managers)
    row = df_cls[df_cls["Mánager"] == chosen].iloc[0]
    chosen_norm = _norm_key(chosen)

    ordered_summary_rounds = sorted(
        [
            rk
            for rk, m_map in summaries_by_round.items()
            if any(_norm_key(k) == chosen_norm and v.strip() for k, v in m_map.items())
        ],
        key=lambda x: int(_round_num_str(x)) if _round_num_str(x).isdigit() else 0,
    )

    if ordered_summary_rounds:
        last_rk = ordered_summary_rounds[-1]
        last_map = {_norm_key(k): v for k, v in summaries_by_round[last_rk].items()}
        clean_last = _clean_lead_name(chosen, last_map[chosen_norm])
        main_last, fact_last = _split_summary_and_fact(clean_last)

        st.markdown(f"### 🗞️ Resumen — {chosen} (JORNADA {_round_num_str(last_rk)})")
        with st.container(border=True):
            st.markdown(main_last)
            if fact_last:
                st.caption(f"💡 *Dato curioso:* {fact_last}")

        if len(ordered_summary_rounds) > 1:
            with st.expander(f"Evolución en la hoja 'Mánagers' para {chosen} ({len(ordered_summary_rounds)} jornadas)"):
                for rk in reversed(ordered_summary_rounds):
                    rk_map = {_norm_key(k): v for k, v in summaries_by_round[rk].items()}
                    st.markdown(f"**JORNADA {_round_num_str(rk)}:** {rk_map[chosen_norm]}")
    else:
        st.caption(f"No hay resumen en la hoja 'Mánagers' del Excel para **{chosen}**.")

    m1, m2, m3, m4, m5, m6 = st.columns(6)
    m1.metric("Puesto", f"#{int(row['Pos'])}")
    m2.metric("Puntos Totales", f"{int(row['Puntos Totales'])}")
    m3.metric("Media / Mediana", f"{row['Media / J']} / {row['Mediana / J']}")
    m4.metric("Últ. 3J (Forma)", f"{row['Últ. 3J']}")
    m5.metric("G+A / PaC", f"{int(row['G+A'])} / {int(row['PaC'])}")
    m6.metric("Pos. Media / Top 3", f"{row['Pos. Media']} / {int(row['Top 3 J'])}")

    r_cols = st.columns(3)
    with r_cols[0]:
        s_pts = df_pts_j.loc[chosen, active_rounds].rename("Puntos")
        fig_m = px.bar(s_pts, x=s_pts.index, y="Puntos", text_auto=".0f", title=f"Puntos por jornada — {chosen}")
        fig_m.update_traces(textposition="outside", textfont=dict(size=11), cliponaxis=False)
        _apply_clean_layout(fig_m, height=320, title=f"Puntos por jornada — {chosen}", top_margin=55)
        fig_m.update_layout(
            showlegend=False,
            xaxis=dict(title="Jornada", tickfont=dict(size=11)),
            yaxis=dict(title="Puntos", tickfont=dict(size=11), rangemode="tozero"),
        )
        st.plotly_chart(fig_m, width="stretch")

    with r_cols[1]:
        s_cap = df_cap_j.loc[chosen, active_rounds].rename("Capitán")
        fig_c = px.bar(
            s_cap,
            x=s_cap.index,
            y="Capitán",
            text_auto=".0f",
            title=f"Puntos capitán — {chosen}",
            color_discrete_sequence=["#0F766E"],
        )
        fig_c.update_traces(textposition="outside", textfont=dict(size=11), cliponaxis=False)
        _apply_clean_layout(fig_c, height=320, title=f"Puntos capitán — {chosen}", top_margin=55)
        fig_c.update_layout(
            showlegend=False,
            xaxis=dict(title="Jornada", tickfont=dict(size=11)),
            yaxis=dict(title="Capitán", tickfont=dict(size=11), rangemode="tozero"),
        )
        st.plotly_chart(fig_c, width="stretch")

    with r_cols[2]:
        df_of = pd.DataFrame(
            {
                "G+A": df_ga_j.loc[chosen, active_rounds],
                "PaC": df_pac_j.loc[chosen, active_rounds],
            }
        )
        fig_o = px.bar(
            df_of,
            x=df_of.index,
            y=["G+A", "PaC"],
            barmode="group",
            text_auto=".0f",
            title=f"Ofensiva (G+A) y PaC — {chosen}",
        )
        fig_o.update_traces(textposition="outside", textfont=dict(size=11), cliponaxis=False)
        _apply_clean_layout(fig_o, height=320, title=f"Ofensiva (G+A) y PaC — {chosen}", top_margin=65)
        fig_o.update_layout(
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="left",
                x=0,
                font=dict(size=11),
            ),
            xaxis=dict(title="Jornada", tickfont=dict(size=11)),
            yaxis=dict(title="Total", tickfont=dict(size=11), rangemode="tozero"),
        )
        st.plotly_chart(fig_o, width="stretch")

# 6. RESÚMENES POR JORNADA
with tab_res:
    st.subheader("Resúmenes por jornada")
    st.caption("Lectura limpia desde la hoja **Mánagers** del Excel.")

    ordered_rounds_all = sorted(
        summaries_by_round.keys(),
        key=lambda x: int(_round_num_str(x)) if _round_num_str(x).isdigit() else 0,
    )

    if not ordered_rounds_all:
        st.info("No se encontraron resúmenes en la hoja 'Mánagers' del Excel.")
    else:
        f_col1, f_col2, f_col3 = st.columns([2, 2, 1.5])
        with f_col1:
            selected_sum_round = st.selectbox(
                "Jornada",
                options=ordered_rounds_all,
                index=len(ordered_rounds_all) - 1,
                format_func=lambda rk: f"JORNADA {_round_num_str(rk)}",
            )
        with f_col2:
            search_q = st.text_input(
                "Filtrar por mánager o texto",
                placeholder="Ej. Donqui, capitán, J11...",
            ).strip().casefold()
        with f_col3:
            st.write("")
            st.write("")
            show_facts = st.toggle("Mostrar dato curioso", value=True)

        mgr_map = summaries_by_round.get(selected_sum_round, {})
        mgr_map_norm = {_norm_key(k): v for k, v in mgr_map.items()}

        items_to_show: list[tuple[str, str]] = []
        for m_name in all_managers:
            if m_name not in active_managers:
                continue
            raw_txt = mgr_map.get(m_name) or mgr_map_norm.get(_norm_key(m_name))
            if not raw_txt:
                continue
            clean_txt = _clean_lead_name(m_name, raw_txt)
            if search_q and search_q not in m_name.casefold() and search_q not in clean_txt.casefold():
                continue
            items_to_show.append((m_name, clean_txt))

        if not items_to_show:
            st.info(
                f"No hay resúmenes que coincidan con el filtro en **JORNADA {_round_num_str(selected_sum_round)}**."
            )
        else:
            st.caption(f"Mostrando {len(items_to_show)} mánagers en **JORNADA {_round_num_str(selected_sum_round)}**.")
            grid_cols = st.columns(2)
            for idx, (m_name, clean_txt) in enumerate(items_to_show):
                main_txt, fact_txt = _split_summary_and_fact(clean_txt)
                with grid_cols[idx % 2]:
                    with st.container(border=True):
                        st.markdown(f"#### {m_name}")
                        st.markdown(main_txt)
                        if show_facts and fact_txt:
                            st.caption(f"💡 *Dato curioso:* {fact_txt}")
