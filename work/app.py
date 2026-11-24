from io import BytesIO
from pathlib import Path
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(
    page_title="Donqui League — Analytics",
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


def find_excel_file(filename: str) -> Path | None:
    """Busca el archivo Excel priorizando el nombre exacto o archivos de la liga."""
    app_dir = Path(__file__).resolve().parent
    search_dirs = [
        app_dir,
        app_dir.parent,
        app_dir.parent.parent,
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
    return s != "" and s.lower() != "nan"


@st.cache_data(show_spinner="Cargando datos de Donqui League...")
def load_league_data(source_key: str, file_bytes: bytes):
    df_orig = pd.read_excel(BytesIO(file_bytes), sheet_name="Datos Originales", header=None)

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
                played_pts.append(float(val))
                round_labels.append(f"J{j_i + 1}")

        n_j = len(played_pts)
        rem_j = max(0, TOTAL_ROUNDS - n_j)

        cap_b = [
            float(df_orig.iloc[ROW_CAP_BASE_START + idx, COL_ROUNDS_START + j])
            if _is_valid_number(df_orig.iloc[ROW_CAP_BASE_START + idx, COL_ROUNDS_START + j])
            else 0.0
            for j in range(n_j)
        ]
        mult = [
            float(df_orig.iloc[ROW_MULT_START + idx, COL_ROUNDS_START + j])
            if _is_valid_number(df_orig.iloc[ROW_MULT_START + idx, COL_ROUNDS_START + j])
            else 0.0
            for j in range(n_j)
        ]
        goals = [
            float(df_orig.iloc[ROW_GOALS_START + idx, COL_ROUNDS_START + j])
            if _is_valid_number(df_orig.iloc[ROW_GOALS_START + idx, COL_ROUNDS_START + j])
            else 0.0
            for j in range(n_j)
        ]
        assists = [
            float(df_orig.iloc[ROW_ASSISTS_START + idx, COL_ROUNDS_START + j])
            if _is_valid_number(df_orig.iloc[ROW_ASSISTS_START + idx, COL_ROUNDS_START + j])
            else 0.0
            for j in range(n_j)
        ]
        pac = [
            float(df_orig.iloc[ROW_PAC_START + idx, COL_ROUNDS_START + j])
            if _is_valid_number(df_orig.iloc[ROW_PAC_START + idx, COL_ROUNDS_START + j])
            else 0.0
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
                float(df_orig.iloc[ROW_PTS_START + k, COL_ROUNDS_START + j])
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

    return df_cls, df_pts_j, df_cap_j, df_ga_j, df_pac_j


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

df_cls, df_pts_j, df_cap_j, df_ga_j, df_pac_j = load_league_data(source_key, file_bytes)

if df_cls.empty or df_pts_j.empty:
    st.info("El archivo se cargó, pero no hay datos de jornadas registrados todavía.")
    st.stop()

current_round = int(df_cls["Jornadas"].max())
rem_rounds = max(0, TOTAL_ROUNDS - current_round)
all_managers = df_cls["Mánager"].tolist()
all_rounds = df_pts_j.columns.tolist()

# --- FILTROS MEJORADOS EN SIDEBAR ---
st.sidebar.header("🎛️ Filtros de análisis")

# 1. Filtro de mánagers con botones rápidos
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

# 2. Filtro de jornadas con presets + rango + detalle
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
tab_vis, tab_evo, tab_reg, tab_proj, tab_mgr = st.tabs(
    [
        "🏆 Visión General",
        "📈 Evolución por Jornada",
        "📊 Regularidad y Capitán",
        "🎯 Proyección J17",
        "🕵️ Ficha por Mánager",
    ]
)

# 1. VISIÓN GENERAL (Limpia y visual: solo métricas solicitadas)
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
        fig_bar.update_traces(textposition="outside", cliponaxis=False)
        fig_bar.update_layout(
            height=400,
            coloraxis_showscale=False,
            margin=dict(l=10, r=45, t=45, b=10),
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
            )
        )
        fig_mm.update_layout(
            barmode="group",
            height=400,
            title="Puntuación Máxima vs. Mínima por Mánager",
            margin=dict(l=10, r=45, t=45, b=10),
        )
        st.plotly_chart(fig_mm, width="stretch")

# 2. EVOLUCIÓN POR JORNADA
with tab_evo:
    sub_pts = df_pts_j.loc[active_managers, active_rounds]
    df_melt = sub_pts.reset_index().melt(
        id_vars="Mánager",
        var_name="Jornada",
        value_name="Puntos",
    )

    fig_line = px.line(
        df_melt,
        x="Jornada",
        y="Puntos",
        color="Mánager",
        markers=True,
        title="Puntuación por Jornada",
    )
    fig_line.update_layout(height=420, hovermode="x unified")
    st.plotly_chart(fig_line, width="stretch")

    fig_heat = px.imshow(
        sub_pts,
        text_auto=True,
        aspect="auto",
        color_continuous_scale="Blues",
        title="Mapa de Calor de Puntos por Jornada",
    )
    fig_heat.update_layout(height=380)
    st.plotly_chart(fig_heat, width="stretch")

# 3. REGULARIDAD Y CAPITÁN (Métricas avanzadas)
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
        fig_scatter.update_traces(textposition="top center")
        fig_scatter.update_layout(height=400, showlegend=False, margin=dict(l=10, r=20, t=45, b=10))
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
        fig_cap.update_traces(textposition="outside", cliponaxis=False)
        fig_cap.update_layout(height=400, margin=dict(l=10, r=45, t=45, b=10))
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
        text_auto=True,
        title="Puntos Actuales vs. Proyección a J17 (Media vs. Últimas 3 Jornadas)",
    )
    fig_proj.update_layout(height=450, xaxis_tickangle=-25)
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
        fig_m = px.bar(s_pts, x=s_pts.index, y="Puntos", text_auto=True, title=f"Puntos por jornada — {chosen}")
        fig_m.update_layout(height=300, showlegend=False)
        st.plotly_chart(fig_m, width="stretch")

    with r_cols[1]:
        s_cap = df_cap_j.loc[chosen, active_rounds].rename("Capitán")
        fig_c = px.bar(
            s_cap,
            x=s_cap.index,
            y="Capitán",
            text_auto=True,
            title=f"Puntos capitán — {chosen}",
            color_discrete_sequence=["#0F766E"],
        )
        fig_c.update_layout(height=300, showlegend=False)
        st.plotly_chart(fig_c, width="stretch")

    with r_cols[2]:
        df_of = pd.DataFrame(
            {
                "G+A": df_ga_j.loc[chosen, active_rounds],
                "PaC": df_pac_j.loc[chosen, active_rounds],
            }
        )
        fig_o = px.bar(df_of, x=df_of.index, y=["G+A", "PaC"], barmode="group", title=f"Ofensiva (G+A) y PaC — {chosen}")
        fig_o.update_layout(height=300)
        st.plotly_chart(fig_o, width="stretch")
