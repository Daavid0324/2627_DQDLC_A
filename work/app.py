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


def find_excel_file(filename: str) -> Path | None:
    app_dir = Path(__file__).resolve().parent
    search_dirs = [
        app_dir,               # /mount/src/2627_dqdlc_a/work
        app_dir.parent,        # /mount/src/2627_dqdlc_a (raíz del repo)
        app_dir.parent.parent, # /mount/src
        Path.cwd(),
    ]

    # 1. Buscar coincidencia exacta
    for directory in search_dirs:
        if directory.exists() and directory.is_dir():
            target = directory / filename
            if target.is_file():
                return target

    # 2. Buscar cualquier .xlsx válido (útil si subes v004, v005, etc.)
    for directory in search_dirs:
        if directory.exists() and directory.is_dir():
            candidates = sorted(
                [p for p in directory.rglob("*.xlsx") if not p.name.startswith("~$")],
                key=lambda p: p.stat().st_mtime,
                reverse=True,
            )
            if candidates:
                return candidates[0]

    return None


@st.cache_data(ttl=60)
def load_league_data_from_source(source):
    df_orig = pd.read_excel(source, sheet_name="Datos Originales", header=None)
    # ... mantén exactamente aquí todo el procesamiento de df_orig ...

    managers = []
    pts_matrix = {}
    cap_matrix = {}
    ga_matrix = {}
    pac_matrix = {}

    for idx in range(10):
        name = str(df_orig.iloc[5 + idx, 0]).strip()
        raw_pts = df_orig.iloc[5 + idx, 8:25].tolist()

        # Considera jornada jugada si hay valor numérico no nulo
        played_pts = []
        round_labels = []
        for j_i, val in enumerate(raw_pts):
            if pd.notna(val) and str(val).strip() != "":
                played_pts.append(float(val))
                round_labels.append(f"J{j_i + 1}")

        n_j = len(played_pts)
        rem_j = max(0, 17 - n_j)

        cap_b = [
            float(df_orig.iloc[19 + idx, 8 + j]) if pd.notna(df_orig.iloc[19 + idx, 8 + j]) else 0.0
            for j in range(n_j)
        ]
        mult = [
            float(df_orig.iloc[69 + idx, 8 + j]) if pd.notna(df_orig.iloc[69 + idx, 8 + j]) else 0.0
            for j in range(n_j)
        ]
        goals = [
            float(df_orig.iloc[32 + idx, 8 + j]) if pd.notna(df_orig.iloc[32 + idx, 8 + j]) else 0.0
            for j in range(n_j)
        ]
        assists = [
            float(df_orig.iloc[44 + idx, 8 + j]) if pd.notna(df_orig.iloc[44 + idx, 8 + j]) else 0.0
            for j in range(n_j)
        ]
        pac = [
            float(df_orig.iloc[56 + idx, 8 + j]) if pd.notna(df_orig.iloc[56 + idx, 8 + j]) else 0.0
            for j in range(n_j)
        ]

        tot = sum(played_pts)
        cap_round_pts = [round(b * m) for b, m in zip(cap_b, mult)]
        cap_pts = sum(cap_round_pts)
        net_bonus = cap_pts - sum(cap_b)
        no_cap = tot - net_bonus

        avg = np.mean(played_pts) if n_j else 0.0
        med = np.median(played_pts) if n_j else 0.0
        sd = np.std(played_pts, ddof=1) if n_j > 1 else 0.0
        cv = (sd / avg) if avg else 0.0
        last3 = np.mean(played_pts[-3:]) if n_j >= 3 else avg
        last_j = int(played_pts[-1]) if n_j else 0

        ranks = []
        for j in range(n_j):
            col_vals = [
                float(df_orig.iloc[5 + k, 8 + j])
                for k in range(10)
                if pd.notna(df_orig.iloc[5 + k, 8 + j]) and str(df_orig.iloc[5 + k, 8 + j]).strip() != ""
            ]
            ordered = sorted(col_vals, reverse=True)
            ranks.append(ordered.index(played_pts[j]) + 1)

        ga_tot = int(sum(goals) + sum(assists))
        managers.append(
            {
                "Mánager": name,
                "Jornadas": n_j,
                "Puntos Totales": int(tot),
                f"J{n_j}": last_j,
                "Media / J": round(avg, 1),
                "Mediana / J": round(med, 1),
                "DE": round(sd, 1),
                "CV": round(cv, 2),
                "Últ. 3J": round(last3, 1),
                "Bonus Cap.": int(net_bonus),
                "Sin Cap.": int(no_cap),
                "Goles": int(sum(goals)),
                "Asistencias": int(sum(assists)),
                "G+A": ga_tot,
                "Pts / (G+A)": round(tot / ga_tot, 1) if ga_tot else 0.0,
                "PaC": int(sum(pac)),
                "Pos. Media": round(np.mean(ranks), 1) if ranks else 0.0,
                "Victorias J": sum(1 for r in ranks if r == 1),
                "Top 3 J": sum(1 for r in ranks if r <= 3),
                "Proy. Media J17": int(round(tot + avg * rem_j)),
                "Proy. Forma J17": int(round(tot + last3 * rem_j)),
                "Mínima": int(min(played_pts)) if played_pts else 0,
                "Máxima": int(max(played_pts)) if played_pts else 0,
            }
        )

        pts_matrix[name] = dict(zip(round_labels, [int(x) for x in played_pts]))
        cap_matrix[name] = dict(zip(round_labels, cap_round_pts))
        ga_matrix[name] = dict(zip(round_labels, [int(g + a) for g, a in zip(goals, assists)]))
        pac_matrix[name] = dict(zip(round_labels, [int(p) for p in pac]))

    df_cls = (
        pd.DataFrame(managers)
        .sort_values(["Puntos Totales", "Media / J"], ascending=False)
        .reset_index(drop=True)
    )
    df_cls.insert(0, "Pos", df_cls.index + 1)
    leader_pts = df_cls["Puntos Totales"].iloc[0]
    df_cls.insert(3, "Dif. Líder", df_cls["Puntos Totales"] - leader_pts)

    df_pts_j = pd.DataFrame.from_dict(pts_matrix, orient="index").reindex(df_cls["Mánager"])
    df_cap_j = pd.DataFrame.from_dict(cap_matrix, orient="index").reindex(df_cls["Mánager"])
    df_ga_j = pd.DataFrame.from_dict(ga_matrix, orient="index").reindex(df_cls["Mánager"])
    df_pac_j = pd.DataFrame.from_dict(pac_matrix, orient="index").reindex(df_cls["Mánager"])
    return df_cls, df_pts_j, df_cap_j, df_ga_j, df_pac_j


st.title("⚽ Donqui League — Centro Analítico")
st.caption("Actualizado automáticamente desde la última versión del Excel de la liga")

col_top_1, col_top_2 = st.columns([6, 1])
with col_top_2:
    if st.button("🔄 Actualizar", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

excel_path = find_excel_file(EXCEL_FILENAME)

if excel_path is not None:
    data_source = str(excel_path)
else:
    st.warning(
        f"No se encontró `{EXCEL_FILENAME}` ni otro `.xlsx` en el repositorio (`work/` o raíz). "
        "Sube el archivo Excel aquí para visualizarlo al instante:"
    )
    uploaded_file = st.file_uploader("Sube el Excel de la liga (.xlsx)", type=["xlsx"])
    if uploaded_file is None:
        st.stop()
    data_source = uploaded_file

df_cls, df_pts_j, df_cap_j, df_ga_j, df_pac_j = load_league_data_from_source(data_source)

df_cls, df_pts_j, df_cap_j, df_ga_j, df_pac_j = load_league_data(EXCEL_FILENAME)
current_round = int(df_cls["Jornadas"].max())

# Sidebar filters
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

# Top KPIs
leader = df_cls.iloc[0]
top_ga = df_cls.sort_values("G+A", ascending=False).iloc[0]
most_consistent = df_cls.sort_values("CV", ascending=True).iloc[0]
hottest = df_cls.sort_values("Últ. 3J", ascending=False).iloc[0]

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Jornada actual", f"J{current_round} / J17", f"{17 - current_round} restantes")
k2.metric("Líder", leader["Mánager"], f"{leader['Puntos Totales']} pts")
k3.metric("Mejor racha (3J)", hottest["Mánager"], f"{hottest['Últ. 3J']} media")
k4.metric("Más regular (CV)", most_consistent["Mánager"], f"CV {most_consistent['CV']}")
k5.metric("Top G+A", top_ga["Mánager"], f"{top_ga['G+A']} contribuciones")

tab1, tab2, tab3, tab4 = st.tabs(
    ["🏆 Clasificación y Métricas", "📈 Evolución por Jornada", "🎯 Proyección J17", "🕵️ Ficha por Mánager"]
)

with tab1:
    view_cls = df_cls[df_cls["Mánager"].isin(selected_managers)]
    st.dataframe(view_cls, use_container_width=True, hide_index=True)

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
        fig_bar.update_traces(textposition="outside")
        fig_bar.update_layout(height=420, margin=dict(l=10, r=20, t=45, b=10))
        st.plotly_chart(fig_bar, use_container_width=True)

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
        st.plotly_chart(fig_scatter, use_container_width=True)

with tab2:
    sub_pts = df_pts_j.loc[selected_managers, selected_rounds]
    df_melt = sub_pts.reset_index().melt(id_vars="index", var_name="Jornada", value_name="Puntos")
    df_melt.rename(columns={"index": "Mánager"}, inplace=True)

    fig_line = px.line(
        df_melt,
        x="Jornada",
        y="Puntos",
        color="Mánager",
        markers=True,
        title="Puntuación por Jornada",
    )
    fig_line.update_layout(height=420, hovermode="x unified")
    st.plotly_chart(fig_line, use_container_width=True)

    fig_heat = px.imshow(
        sub_pts,
        text_auto=True,
        aspect="auto",
        color_continuous_scale="Blues",
        title="Mapa de Calor de Puntos por Jornada",
    )
    fig_heat.update_layout(height=380)
    st.plotly_chart(fig_heat, use_container_width=True)

with tab3:
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
    st.plotly_chart(fig_proj, use_container_width=True)

    st.dataframe(
        view_proj[["Pos", "Mánager", "Puntos Totales", "Media / J", "Últ. 3J", "Proy. Media J17", "Proy. Forma J17"]],
        use_container_width=True,
        hide_index=True,
    )

with tab4:
    chosen = st.selectbox("Elige un mánager", options=df_cls["Mánager"].tolist())
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
        s_pts = pd.Series(df_pts_j.loc[chosen, selected_rounds], name="Puntos")
        fig_m = px.bar(s_pts, x=s_pts.index, y="Puntos", text_auto=True, title=f"Puntos por jornada — {chosen}")
        fig_m.update_layout(height=300, showlegend=False)
        st.plotly_chart(fig_m, use_container_width=True)

    with r_cols[1]:
        s_cap = pd.Series(df_cap_j.loc[chosen, selected_rounds], name="Capitán")
        fig_c = px.bar(s_cap, x=s_cap.index, y="Capitán", text_auto=True, title=f"Puntos capitán — {chosen}", color_discrete_sequence=["#0F766E"])
        fig_c.update_layout(height=300, showlegend=False)
        st.plotly_chart(fig_c, use_container_width=True)

    with r_cols[2]:
        df_of = pd.DataFrame({
            "G+A": df_ga_j.loc[chosen, selected_rounds],
            "PaC": df_pac_j.loc[chosen, selected_rounds],
        })
        fig_o = px.bar(df_of, x=df_of.index, y=["G+A", "PaC"], barmode="group", title=f"Ofensiva (G+A) y PaC — {chosen}")
        fig_o.update_layout(height=300)
        st.plotly_chart(fig_o, use_container_width=True)
