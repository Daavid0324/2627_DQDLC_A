from pathlib import Path
import numpy as np
import pandas as pd
import plotly.express as px
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

    # 1. Coincidencia exacta
    for folder in valid_dirs:
        exact = folder / filename
        if exact.is_file():
            return exact

    # 2. Archivos .xlsx que parezcan de la liga
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
    from io import BytesIO

    df_orig = pd.read_excel(BytesIO(file_bytes), sheet_name="Datos Originales", header=None)

    managers = []
    pts_matrix = {}
    cap_matrix = {}
    ga_matrix = {}
    pac_matrix = {}

    for idx in range(NUM_MANAGERS):
        raw_name = df_orig.iloc[ROW_PTS_START + idx, 0]
        name = str(raw_name).strip() if pd.notna(raw_name) and str(raw_name).strip() else f"Mánager {idx + 1}"
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

        # Cálculo robusto de posición en cada jornada (maneja empates sin ValueError)
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
            # Busca el índice correspondiente a played_pts[j]
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

st.sidebar.header("Filtros de análisis")
selected_managers = st.sidebar.multiselect(
    "Seleccionar mánagers",
    options=df_cls["Mánager"].tolist(),
    default=df_cls["Mánager"].tolist(),
)
if not selected_managers:
    selected_managers = df_cls["Mánager"].tolist()

all_rounds = df_pts_j.columns.tolist()
selected_rounds = st.sidebar.multiselect(
    "Jornadas a mostrar",
    options=all_rounds,
    default=all_rounds,
)
if not selected_rounds:
    selected_rounds = all_rounds

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

tab1, tab2, tab3, tab4 = st.tabs(
    ["🏆 Clasificación y Métricas", "📈 Evolución por Jornada", "🎯 Proyección J17", "🕵️ Ficha por Mánager"]
)

with tab1:
    view_cls = df_cls[df_cls["Mánager"].isin(selected_managers)]
    st.dataframe(view_cls, width="stretch", hide_index=True)

    c_left, c_right = st.columns(2)
    with c_left:
        fig_bar = px.bar(
            view_cls.sort_values("Puntos Totales", ascending=True),
            x="Puntos Totales",
            y="Mánager",
            orientation="h",
            text="Puntos Totales",
            color="Media / J",
            color_continuous_scale="Tealgrn",
            title=f"Puntos Totales Acumulados (J1–J{current_round})",
        )
        fig_bar.update_traces(textposition="outside", cliponaxis=False)
        fig_bar.update_layout(height=420, margin=dict(l=10, r=45, t=45, b=10))
        st.plotly_chart(fig_bar, width="stretch")

    with c_right:
        fig_scatter = px.scatter(
            view_cls,
            x="Media / J",
            y="CV",
            size="Top 3 J",
            color="Mánager",
            text="Mánager",
            title="Regularidad vs. Rendimiento Medio (menor CV = más constante)",
            labels={"CV": "Coeficiente de Variación (DE / Media)"},
        )
        fig_scatter.update_traces(textposition="top center")
        fig_scatter.update_layout(height=420, showlegend=False, margin=dict(l=10, r=20, t=45, b=10))
        st.plotly_chart(fig_scatter, width="stretch")

with tab2:
    sub_pts = df_pts_j.loc[selected_managers, selected_rounds]
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

with tab3:
    if rem_rounds == 0:
        st.info(f"Se han completado las {TOTAL_ROUNDS} jornadas; las proyecciones coinciden con los puntos finales.")

    view_proj = df_cls[df_cls["Mánager"].isin(selected_managers)].copy()
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
        view_proj[["Pos", "Mánager", "Puntos Totales", "Media / J", "Últ. 3J", "Proy. Media J17", "Proy. Forma J17"]],
        width="stretch",
        hide_index=True,
    )

with tab4:
    manager_options = df_cls["Mánager"].tolist()
    chosen = st.selectbox("Elige un mánager", options=manager_options)
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
        s_pts = df_pts_j.loc[chosen, selected_rounds].rename("Puntos")
        fig_m = px.bar(s_pts, x=s_pts.index, y="Puntos", text_auto=True, title=f"Puntos por jornada — {chosen}")
        fig_m.update_layout(height=300, showlegend=False)
        st.plotly_chart(fig_m, width="stretch")

    with r_cols[1]:
        s_cap = df_cap_j.loc[chosen, selected_rounds].rename("Capitán")
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
        df_of = pd.DataFrame({
            "G+A": df_ga_j.loc[chosen, selected_rounds],
            "PaC": df_pac_j.loc[chosen, selected_rounds],
        })
        fig_o = px.bar(df_of, x=df_of.index, y=["G+A", "PaC"], barmode="group", title=f"Ofensiva (G+A) y PaC — {chosen}")
        fig_o.update_layout(height=300)
        st.plotly_chart(fig_o, width="stretch")
