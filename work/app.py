from __future__ import annotations

from pathlib import Path
import streamlit as st

from donqui_league.charts import (
    build_manager_color_map,
    make_captain_bonus_bar,
    make_dumbbell_range_chart,
    make_evolution_line_chart,
    make_manager_mini_bars,
    make_points_heatmap,
    make_projection_chart,
    make_regularity_quadrant_scatter,
    make_total_points_bar,
)
from donqui_league.data_loader import (
    EXCEL_FILENAME,
    TOTAL_ROUNDS,
    clean_lead_name,
    find_excel_file,
    load_raw_excel_data,
    norm_key,
    round_num_str,
    split_summary_and_fact,
)
from donqui_league.metrics import build_manager_diagnostic, compute_league_metrics

st.set_page_config(
    page_title="Donqui League A — Analytics",
    page_icon="⚽",
    layout="wide",
    initial_sidebar_state="expanded",
)

APP_DIR = Path(__file__).resolve().parent


@st.cache_data(show_spinner="Leyendo archivo Excel...")
def cached_read_file_bytes(path_str: str, mtime_ns: int, size: int) -> bytes:
    """Cachea la lectura de disco vinculada a mtime y tamaño para evitar I/O en cada rerun."""
    return Path(path_str).read_bytes()


@st.cache_data(show_spinner="Procesando datos de Donqui League...")
def cached_load_raw(source_key: str, file_bytes: bytes) -> dict:
    return load_raw_excel_data(file_bytes)


@st.cache_data(show_spinner=False)
def cached_compute_metrics(source_key: str, file_bytes: bytes, rounds_tuple: tuple[str, ...]):
    raw_data = cached_load_raw(source_key, file_bytes)
    return compute_league_metrics(raw_data, list(rounds_tuple))


st.title("⚽ Donqui League — Centro Analítico")
st.caption("Actualizado desde la última versión del Excel de la liga")

col_top_1, col_top_2 = st.columns([6, 1])
with col_top_2:
    if st.button("🔄 Actualizar", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

excel_path = find_excel_file(EXCEL_FILENAME, base_dir=APP_DIR)
if excel_path is not None:
    stat = excel_path.stat()
    resolved_str = str(excel_path.resolve())
    source_key = f"file::{resolved_str}::{stat.st_mtime_ns}::{stat.st_size}"
    file_bytes = cached_read_file_bytes(resolved_str, stat.st_mtime_ns, stat.st_size)
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

try:
    raw_data = cached_load_raw(source_key, file_bytes)
except ValueError as err:
    st.error(f"Error de estructura en el archivo Excel: {err}")
    st.stop()

df_pts_j = raw_data["df_pts"]
df_cap_j = raw_data["df_cap_pts"]
df_ga_j = raw_data["df_ga"]
df_pac_j = raw_data["df_pac"]
summaries_by_round = raw_data["summaries_by_round"]
all_rounds = raw_data["played_cols"]

if not all_rounds or df_pts_j.empty:
    st.info("El archivo se cargó, pero no hay datos de jornadas registrados todavía.")
    st.stop()

# Clasificación completa de la temporada (para orden base, paleta y proyecciones)
df_season = cached_compute_metrics(source_key, file_bytes, tuple(all_rounds))
current_round = len(all_rounds)
rem_rounds = max(0, TOTAL_ROUNDS - current_round)
all_managers = df_season["Mánager"].tolist()
color_map = build_manager_color_map(all_managers)

# --- FILTROS EN SIDEBAR ---
st.sidebar.header("🎛️ Filtros de análisis")

# Reiniciar filtro si cambió el archivo cargado a otro con distinta lista de mánagers
if (
    st.session_state.get("mgr_filter_source") != source_key
    or "mgr_filter" not in st.session_state
    or set(st.session_state["mgr_filter"]) - set(all_managers)
):
    st.session_state["mgr_filter"] = all_managers.copy()
    st.session_state["mgr_filter_source"] = source_key

b_all, b_none = st.sidebar.columns(2)
if b_all.button("Todos", use_container_width=True):
    st.session_state["mgr_filter"] = all_managers.copy()
if b_none.button("Limpiar", use_container_width=True):
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

is_full_season = active_rounds == all_rounds
round_range_label = (
    f"J1–J{current_round}"
    if is_full_season
    else f"{active_rounds[0]}–{active_rounds[-1]} ({len(active_rounds)}J)"
)
st.sidebar.caption(f"Jornadas activas: **{', '.join(active_rounds)}**")

# Métricas recalculadas dinámicamente sobre el tramo de jornadas seleccionado
df_active = (
    df_season
    if is_full_season
    else cached_compute_metrics(source_key, file_bytes, tuple(active_rounds))
)

# --- RESUMEN SUPERIOR (KPIs DINÁMICOS AL TRAMO) ---
scope_tag = "temporada" if is_full_season else "tramo"
leader = df_active.iloc[0]
top_ga = df_active.sort_values("G+A", ascending=False).iloc[0]
most_consistent = df_active.sort_values("CV", ascending=True).iloc[0]
hottest = df_season.sort_values("Últ. 3J", ascending=False).iloc[0]

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Jornada actual", f"J{current_round} / J{TOTAL_ROUNDS}", f"{rem_rounds} restantes")
k2.metric(f"Líder ({scope_tag})", leader["Mánager"], f"{leader['Puntos Totales']} pts")
k3.metric("Mejor racha (3J)", hottest["Mánager"], f"{hottest['Últ. 3J']} media")
k4.metric(f"Más regular ({scope_tag})", most_consistent["Mánager"], f"CV {most_consistent['CV']}")
k5.metric(f"Top G+A ({scope_tag})", top_ga["Mánager"], f"{top_ga['G+A']} contribuciones")

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
    if not is_full_season:
        st.info(f"Mostrando clasificación y estadísticas acumuladas del tramo filtrado: **{round_range_label}**.")

    # Jerarquía lógica: Identidad -> Rendimiento global -> Aporte estadístico -> Extremos
    overview_cols = [
        "Pos",
        "Mánager",
        "Puntos Totales",
        "Dif. Líder",
        "Media / J",
        "Pos. Media",
        "G+A",
        "Goles",
        "Asistencias",
        "PaC",
        "Máxima",
        "Mínima",
    ]
    view_vis = df_active[df_active["Mánager"].isin(active_managers)][overview_cols]

    st.dataframe(
        view_vis,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Puntos Totales": st.column_config.ProgressColumn(
                "Puntos Totales",
                min_value=0,
                max_value=max(int(df_active["Puntos Totales"].max()), 1),
                format="%d",
            ),
            "Media / J": st.column_config.NumberColumn("Media / J", format="%.1f"),
            "Pos. Media": st.column_config.NumberColumn("Pos. Media", format="%.1f"),
            "PaC": st.column_config.NumberColumn("PaC", help="Porterías a cero"),
        },
    )

    c_left, c_right = st.columns(2)
    with c_left:
        fig_bar = make_total_points_bar(view_vis, round_range_label, color_map)
        st.plotly_chart(fig_bar, use_container_width=True)

    with c_right:
        fig_mm = make_dumbbell_range_chart(view_vis, round_range_label)
        st.plotly_chart(fig_mm, use_container_width=True)

# 2. EVOLUCIÓN POR JORNADA (SEGUIMIENTO)
with tab_evo:
    st.subheader("Seguimiento por jornada")
    st.caption(f"Comparativa alineada de puntuación y rendimiento en **{round_range_label}**.")

    sub_pts = df_pts_j.loc[active_managers, active_rounds].copy().dropna(how="all")

    if sub_pts.empty:
        st.info("No hay datos de seguimiento para los filtros seleccionados.")
    else:
        fig_line = make_evolution_line_chart(sub_pts, active_rounds, color_map)
        st.plotly_chart(fig_line, use_container_width=True)

        fig_heat = make_points_heatmap(sub_pts)
        st.plotly_chart(fig_heat, use_container_width=True)

        with st.expander("Ver tabla numérica de seguimiento"):
            num_cols = {col: st.column_config.NumberColumn(col, format="%d") for col in active_rounds}
            st.dataframe(
                sub_pts.reset_index().rename(columns={"index": "Mánager"}),
                use_container_width=True,
                hide_index=True,
                column_config=num_cols,
            )

# 3. REGULARIDAD Y CAPITÁN
with tab_reg:
    if not is_full_season:
        st.caption(f"Métricas de regularidad y aporte de capitán calculadas sobre **{round_range_label}**.")

    view_deep = df_active[df_active["Mánager"].isin(active_managers)].copy()

    # Subsección A: Consistencia y Forma
    st.markdown("#### 📐 Consistencia y Forma")
    consistency_cols = [
        "Pos",
        "Mánager",
        "Media / J",
        "Mediana / J",
        "DE",
        "CV",
        "Últ. 3J",
        "Pos. Media",
        "Victorias J",
        "Top 3 J",
        "Pts / (G+A)",
    ]
    st.dataframe(
        view_deep[consistency_cols],
        use_container_width=True,
        hide_index=True,
        column_config={
            "Media / J": st.column_config.NumberColumn("Media / J", format="%.1f"),
            "Mediana / J": st.column_config.NumberColumn("Mediana / J", format="%.1f"),
            "DE": st.column_config.NumberColumn("DE", format="%.1f"),
            "CV": st.column_config.NumberColumn("CV", format="%.2f"),
            "Últ. 3J": st.column_config.NumberColumn("Últ. 3J", format="%.1f"),
            "Pos. Media": st.column_config.NumberColumn("Pos. Media", format="%.1f"),
            "Pts / (G+A)": st.column_config.NumberColumn("Pts / (G+A)", format="%.1f"),
        },
    )

    # Subsección B: Gestión de Capitán
    st.markdown("#### ©️ Gestión de Capitán")
    captain_cols = [
        "Pos",
        "Mánager",
        "Pts Cap.",
        "Bonus Cap.",
        "Sin Cap.",
        "% Aporte Cap.",
        "Mult. Promedio",
        "Veces x3",
        "Veces x2",
        "Veces x1,5",
        "Cap. Top (≥5)",
    ]
    st.dataframe(
        view_deep[captain_cols],
        use_container_width=True,
        hide_index=True,
        column_config={
            "% Aporte Cap.": st.column_config.NumberColumn("% Aporte Cap.", format="%.1f%%"),
            "Mult. Promedio": st.column_config.NumberColumn("Mult. Promedio", format="%.2f"),
        },
    )

    r_left, r_right = st.columns(2)
    with r_left:
        fig_scatter = make_regularity_quadrant_scatter(view_deep, color_map)
        st.plotly_chart(fig_scatter, use_container_width=True)

    with r_right:
        fig_cap = make_captain_bonus_bar(view_deep, color_map)
        st.plotly_chart(fig_cap, use_container_width=True)

# 4. PROYECCIÓN J17
with tab_proj:
    if rem_rounds == 0:
        st.info(f"Se han completado las {TOTAL_ROUNDS} jornadas; las proyecciones coinciden con los puntos finales.")
    elif not is_full_season:
        st.caption(
            "Las proyecciones a J17 parten siempre de los **puntos acumulados reales de toda la temporada** "
            f"(J1–J{current_round})."
        )

    view_proj = df_season[df_season["Mánager"].isin(active_managers)].copy()
    leader_season_pts = int(df_season["Puntos Totales"].iloc[0]) if not df_season.empty else 0

    fig_proj = make_projection_chart(view_proj, leader_season_pts, TOTAL_ROUNDS)
    st.plotly_chart(fig_proj, use_container_width=True)

    st.dataframe(
        view_proj[
            [
                "Pos",
                "Mánager",
                "Puntos Totales",
                "Media / J",
                "Últ. 3J",
                "Proy. Media J17",
                "Proy. Forma J17",
                "Dif. Proy. (Forma - Media)",
            ]
        ],
        use_container_width=True,
        hide_index=True,
        column_config={
            "Media / J": st.column_config.NumberColumn("Media / J", format="%.1f"),
            "Últ. 3J": st.column_config.NumberColumn("Últ. 3J", format="%.1f"),
            "Dif. Proy. (Forma - Media)": st.column_config.NumberColumn(
                "Tendencia (Forma - Media)",
                format="%+d pts",
                help="Diferencia de puntos proyectados entre mantener el ritmo de las últimas 3J frente a la media global.",
            ),
        },
    )

# 5. FICHA POR MÁNAGER
with tab_mgr:
    chosen = st.selectbox("Elige un mánager", options=all_managers)
    row = df_season[df_season["Mánager"] == chosen].iloc[0]
    chosen_norm = norm_key(chosen)

    # Diagnóstico rápido de fortaleza y punto de atención
    diag = build_manager_diagnostic(row, df_season)
    d1, d2 = st.columns(2)
    with d1:
        st.success(f"**Fortaleza principal:** {diag['fortaleza']}")
    with d2:
        st.warning(f"**Punto de atención:** {diag['atencion']}")

    ordered_summary_rounds = sorted(
        [
            rk
            for rk, m_map in summaries_by_round.items()
            if any(norm_key(k) == chosen_norm and v.strip() for k, v in m_map.items())
        ],
        key=lambda x: int(round_num_str(x)) if round_num_str(x).isdigit() else 0,
    )

    if ordered_summary_rounds:
        last_rk = ordered_summary_rounds[-1]
        last_map = {norm_key(k): v for k, v in summaries_by_round[last_rk].items()}
        raw_last = last_map.get(chosen_norm, "")
        clean_last = clean_lead_name(chosen, raw_last)
        main_last, fact_last = split_summary_and_fact(clean_last)

        st.markdown(f"### 🗞️ Resumen — {chosen} (JORNADA {round_num_str(last_rk)})")
        with st.container(border=True):
            st.markdown(main_last)
            if fact_last:
                st.caption(f"💡 *Dato curioso:* {fact_last}")

        if len(ordered_summary_rounds) > 1:
            with st.expander(f"Evolución en la hoja 'Mánagers' para {chosen} ({len(ordered_summary_rounds)} jornadas)"):
                for rk in reversed(ordered_summary_rounds):
                    rk_map = {norm_key(k): v for k, v in summaries_by_round[rk].items()}
                    txt_rk = rk_map.get(chosen_norm, "")
                    if txt_rk:
                        st.markdown(f"**JORNADA {round_num_str(rk)}:** {txt_rk}")
    else:
        st.caption(f"No hay resumen en la hoja 'Mánagers' del Excel para **{chosen}**.")

    m1, m2, m3, m4, m5, m6 = st.columns(6)
    m1.metric("Puesto", f"#{int(row['Pos'])}")
    m2.metric("Puntos Totales", f"{int(row['Puntos Totales'])}")
    m3.metric("Media / Mediana", f"{row['Media / J']} / {row['Mediana / J']}")
    m4.metric("Últ. 3J (Forma)", f"{row['Últ. 3J']}")
    m5.metric("G+A / PaC", f"{int(row['G+A'])} / {int(row['PaC'])}")
    m6.metric("Pos. Media / Top 3", f"{row['Pos. Media']} / {int(row['Top 3 J'])}")

    fig_m, fig_c, fig_o = make_manager_mini_bars(
        chosen=chosen,
        active_rounds=active_rounds,
        df_pts=df_pts_j,
        df_cap_pts=df_cap_j,
        df_ga=df_ga_j,
        df_pac=df_pac_j,
        manager_color=color_map.get(chosen, "#2563EB"),
    )
    r_cols = st.columns(3)
    with r_cols[0]:
        st.plotly_chart(fig_m, use_container_width=True)
    with r_cols[1]:
        st.plotly_chart(fig_c, use_container_width=True)
    with r_cols[2]:
        st.plotly_chart(fig_o, use_container_width=True)

# 6. RESÚMENES POR JORNADA
with tab_res:
    st.subheader("Resúmenes por jornada")
    st.caption("Lectura limpia desde la hoja **Mánagers** del Excel.")

    ordered_rounds_all = sorted(
        summaries_by_round.keys(),
        key=lambda x: int(round_num_str(x)) if round_num_str(x).isdigit() else 0,
    )

    if not ordered_rounds_all:
        st.info("No se encontraron resúmenes en la hoja 'Mánagers' del Excel.")
    else:
        f_col1, f_col2, f_col3, f_col4 = st.columns([1.6, 2, 1.8, 1.3])
        with f_col1:
            selected_sum_round = st.selectbox(
                "Jornada",
                options=ordered_rounds_all,
                index=len(ordered_rounds_all) - 1,
                format_func=lambda rk: f"JORNADA {round_num_str(rk)}",
            )
        with f_col2:
            search_q = st.text_input(
                "Filtrar por mánager o texto",
                placeholder="Ej. Donqui, capitán, J11...",
            ).strip().casefold()
        with f_col3:
            sort_mode = st.selectbox(
                "Ordenar tarjetas por",
                options=["Clasificación general", "Puntos en la jornada"],
            )
        with f_col4:
            st.write("")
            st.write("")
            show_facts = st.toggle("Mostrar dato curioso", value=True)

        mgr_map = summaries_by_round.get(selected_sum_round, {})
        mgr_map_norm = {norm_key(k): v for k, v in mgr_map.items()}

        items_to_show: list[tuple[str, str, int | None]] = []
        for m_name in all_managers:
            if m_name not in active_managers:
                continue
            raw_txt = mgr_map.get(m_name) or mgr_map_norm.get(norm_key(m_name))
            if not raw_txt:
                continue
            clean_txt = clean_lead_name(m_name, raw_txt)
            if search_q and search_q not in m_name.casefold() and search_q not in clean_txt.casefold():
                continue

            round_pts: int | None = None
            if selected_sum_round in df_pts_j.columns and m_name in df_pts_j.index:
                val_p = df_pts_j.loc[m_name, selected_sum_round]
                if val_p == val_p:  # not NaN
                    round_pts = int(round(float(val_p)))

            items_to_show.append((m_name, clean_txt, round_pts))

        if sort_mode == "Puntos en la jornada":
            items_to_show.sort(key=lambda item: (item[2] is not None, item[2] or -1), reverse=True)

        if not items_to_show:
            st.info(
                f"No hay resúmenes que coincidan con el filtro en **JORNADA {round_num_str(selected_sum_round)}**."
            )
        else:
            st.caption(
                f"Mostrando {len(items_to_show)} mánagers en **JORNADA {round_num_str(selected_sum_round)}**."
            )
            grid_cols = st.columns(2)
            for idx, (m_name, clean_txt, round_pts) in enumerate(items_to_show):
                main_txt, fact_txt = split_summary_and_fact(clean_txt)
                pts_badge = f" · `{round_pts} pts`" if round_pts is not None else ""
                with grid_cols[idx % 2]:
                    with st.container(border=True):
                        st.markdown(f"#### {m_name}{pts_badge}")
                        st.markdown(main_txt)
                        if show_facts and fact_txt:
                            st.caption(f"💡 *Dato curioso:* {fact_txt}")
