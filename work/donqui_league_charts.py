from __future__ import annotations

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# Paleta accesible y consistente para asignar un color fijo a cada mánager
MANAGER_PALETTE = [
    "#0F766E",  # Teal oscuro
    "#2563EB",  # Azul real
    "#D97706",  # Ámbar
    "#DC2626",  # Rojo coral
    "#7C3AED",  # Violeta
    "#059669",  # Esmeralda
    "#DB2777",  # Magenta
    "#0284C7",  # Cian
    "#EA580C",  # Naranja
    "#4F46E5",  # Índigo
    "#65A30D",  # Lima oscuro
    "#475569",  # Pizarra
]


def build_manager_color_map(managers: list[str]) -> dict[str, str]:
    """Asigna un color determinista a cada mánager para mantener coherencia en toda la app."""
    return {
        m: MANAGER_PALETTE[idx % len(MANAGER_PALETTE)]
        for idx, m in enumerate(managers)
    }


def apply_clean_layout(
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


def make_total_points_bar(
    view_df: pd.DataFrame,
    round_label: str,
    color_map: dict[str, str],
) -> go.Figure:
    bar_df = view_df.sort_values("Puntos Totales", ascending=True)
    fig = px.bar(
        bar_df,
        x="Puntos Totales",
        y="Mánager",
        orientation="h",
        text="Puntos Totales",
        color="Mánager",
        color_discrete_map=color_map,
        title=f"Puntos Totales ({round_label})",
    )
    fig.update_traces(
        textposition="outside",
        textfont=dict(size=12),
        cliponaxis=False,
    )
    apply_clean_layout(
        fig,
        height=410,
        title=f"Puntos Totales ({round_label})",
        top_margin=55,
        right_margin=45,
    )
    fig.update_layout(
        showlegend=False,
        xaxis=dict(title="Puntos Totales", tickfont=dict(size=12), rangemode="tozero"),
        yaxis=dict(title="", tickfont=dict(size=12)),
    )
    return fig


def make_dumbbell_range_chart(view_df: pd.DataFrame, round_label: str) -> go.Figure:
    """
    Dumbbell plot: conecta Mínima y Máxima por mánager y marca la Media / J,
    permitiendo visualizar amplitud y centro de rendimiento de un vistazo.
    """
    range_df = view_df.sort_values("Media / J", ascending=True)
    fig = go.Figure()

    # Segmentos horizontales Mínima -> Máxima
    for _, r in range_df.iterrows():
        fig.add_trace(
            go.Scatter(
                x=[r["Mínima"], r["Máxima"]],
                y=[r["Mánager"], r["Mánager"]],
                mode="lines",
                line=dict(color="#94A3B8", width=3),
                showlegend=False,
                hoverinfo="skip",
            )
        )

    fig.add_trace(
        go.Scatter(
            name="Mínima",
            x=range_df["Mínima"],
            y=range_df["Mánager"],
            mode="markers+text",
            text=range_df["Mínima"],
            textposition="middle left",
            textfont=dict(size=11, color="#DC2626"),
            marker=dict(color="#DC2626", size=10, symbol="circle"),
            hovertemplate="<b>%{y}</b><br>Mínima: %{x} pts<extra></extra>",
        )
    )

    fig.add_trace(
        go.Scatter(
            name="Media / J",
            x=range_df["Media / J"],
            y=range_df["Mánager"],
            mode="markers",
            marker=dict(color="#0F766E", size=11, symbol="diamond"),
            hovertemplate="<b>%{y}</b><br>Media: %{x:.1f} pts<extra></extra>",
        )
    )

    fig.add_trace(
        go.Scatter(
            name="Máxima",
            x=range_df["Máxima"],
            y=range_df["Mánager"],
            mode="markers+text",
            text=range_df["Máxima"],
            textposition="middle right",
            textfont=dict(size=11, color="#2563EB"),
            marker=dict(color="#2563EB", size=10, symbol="circle"),
            hovertemplate="<b>%{y}</b><br>Máxima: %{x} pts<extra></extra>",
        )
    )

    x_min = max(0, float(range_df["Mínima"].min()) - 12)
    x_max = float(range_df["Máxima"].max()) + 14

    apply_clean_layout(
        fig,
        height=410,
        title=f"Amplitud de Puntuación: Mínima · Media · Máxima ({round_label})",
        top_margin=70,
        right_margin=35,
    )
    fig.update_layout(
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="left",
            x=0,
            font=dict(size=12),
        ),
        xaxis=dict(title="Puntos", tickfont=dict(size=12), range=[x_min, x_max]),
        yaxis=dict(title="", tickfont=dict(size=12)),
    )
    return fig


def make_evolution_line_chart(
    sub_pts: pd.DataFrame,
    active_rounds: list[str],
    color_map: dict[str, str],
) -> go.Figure:
    df_melt = (
        sub_pts.reset_index()
        .rename(columns={"index": "Mánager"})
        .melt(id_vars="Mánager", var_name="Jornada", value_name="Puntos")
        .dropna(subset=["Puntos"])
    )
    df_melt["Puntos"] = df_melt["Puntos"].round(0).astype(int)

    fig = px.line(
        df_melt,
        x="Jornada",
        y="Puntos",
        color="Mánager",
        color_discrete_map=color_map,
        category_orders={"Jornada": active_rounds},
        markers=True,
        title="Evolución de puntos por jornada",
    )
    fig.update_traces(
        marker=dict(size=7, line=dict(width=1, color="#111827")),
        line=dict(width=2.2),
    )
    apply_clean_layout(
        fig,
        height=430,
        title="Evolución de puntos por jornada",
        top_margin=70,
        right_margin=25,
    )
    fig.update_layout(
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
    return fig


def make_points_heatmap(sub_pts: pd.DataFrame) -> go.Figure:
    pts_min = int(sub_pts.min().min()) if not sub_pts.isna().all().all() else 0
    pts_max = int(sub_pts.max().max()) if not sub_pts.isna().all().all() else 100

    fig = px.imshow(
        sub_pts,
        text_auto=".0f",
        aspect="auto",
        color_continuous_scale="Tealgrn",
        zmin=pts_min,
        zmax=pts_max,
        title="Matriz de puntos por jornada",
    )
    fig.update_traces(textfont=dict(size=12))
    apply_clean_layout(
        fig,
        height=380,
        title="Matriz de puntos por jornada",
        top_margin=55,
        right_margin=25,
    )
    fig.update_layout(
        coloraxis_colorbar=dict(title="Pts", thickness=14, len=0.85),
        xaxis=dict(title="Jornada", tickfont=dict(size=12)),
        yaxis=dict(title="", tickfont=dict(size=12)),
    )
    return fig


def make_regularity_quadrant_scatter(
    view_deep: pd.DataFrame,
    color_map: dict[str, str],
) -> go.Figure:
    """
    Scatter Media/J vs CV con líneas de referencia (medianas del grupo)
    formando 4 cuadrantes interpretables y hover limpio sin solapes.
    """
    med_avg = float(view_deep["Media / J"].median()) if not view_deep.empty else 0.0
    med_cv = float(view_deep["CV"].median()) if not view_deep.empty else 0.0

    fig = px.scatter(
        view_deep,
        x="Media / J",
        y="CV",
        size="Top 3 J",
        color="Mánager",
        color_discrete_map=color_map,
        hover_name="Mánager",
        custom_data=["Mediana / J", "DE", "Top 3 J", "Victorias J"],
        title="Cuadrantes de Regularidad (CV) vs. Media por Jornada",
    )
    fig.update_traces(
        marker=dict(line=dict(width=1.2, color="#111827"), opacity=0.9),
        hovertemplate=(
            "<b>%{hovertext}</b><br>"
            "Media / J: %{x:.1f} pts<br>"
            "CV: %{y:.2f}<br>"
            "Mediana / J: %{customdata[0]:.1f} pts<br>"
            "DE: %{customdata[1]:.1f}<br>"
            "Podios (Top 3): %{customdata[2]} · Victorias: %{customdata[3]}"
            "<extra></extra>"
        ),
    )

    # Líneas de referencia de cuadrantes
    fig.add_vline(
        x=med_avg,
        line_dash="dash",
        line_color="#64748B",
        annotation_text=f"Mediana Media ({med_avg:.1f})",
        annotation_position="top left",
        annotation_font=dict(size=10, color="#475569"),
    )
    fig.add_hline(
        y=med_cv,
        line_dash="dash",
        line_color="#64748B",
        annotation_text=f"Mediana CV ({med_cv:.2f})",
        annotation_position="bottom right",
        annotation_font=dict(size=10, color="#475569"),
    )

    # Etiquetas de cuadrantes en las esquinas
    x_min, x_max = float(view_deep["Media / J"].min()), float(view_deep["Media / J"].max())
    y_max = float(view_deep["CV"].max())
    fig.add_annotation(
        x=x_max, y=0.01, text="★ Élite regular", showarrow=False,
        xanchor="right", yanchor="bottom", font=dict(size=10, color="#0F766E"),
    )
    fig.add_annotation(
        x=x_max, y=y_max, text="⚡ Explosivo / irregular", showarrow=False,
        xanchor="right", yanchor="top", font=dict(size=10, color="#D97706"),
    )
    fig.add_annotation(
        x=x_min, y=0.01, text="🛡️ Regular bajo", showarrow=False,
        xanchor="left", yanchor="bottom", font=dict(size=10, color="#475569"),
    )
    fig.add_annotation(
        x=x_min, y=y_max, text="⚠️ Irregular bajo", showarrow=False,
        xanchor="left", yanchor="top", font=dict(size=10, color="#DC2626"),
    )

    apply_clean_layout(
        fig,
        height=420,
        title="Cuadrantes de Regularidad (CV) vs. Media por Jornada",
        top_margin=60,
        right_margin=25,
    )
    fig.update_layout(
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="left",
            x=0,
            font=dict(size=11),
        ),
        xaxis=dict(title="Media / J (mayor = mejor)", tickfont=dict(size=12)),
        yaxis=dict(title="CV (menor = más regular)", tickfont=dict(size=12), rangemode="tozero"),
    )
    return fig


def make_captain_bonus_bar(
    view_deep: pd.DataFrame,
    color_map: dict[str, str],
) -> go.Figure:
    cap_df = view_deep.sort_values("Bonus Cap.", ascending=True)
    fig = px.bar(
        cap_df,
        x="Bonus Cap.",
        y="Mánager",
        orientation="h",
        text="Bonus Cap.",
        color="Mánager",
        color_discrete_map=color_map,
        title="Bonus Neto de Capitán por Mánager",
    )
    fig.update_traces(
        textposition="outside",
        textfont=dict(size=12),
        cliponaxis=False,
    )
    apply_clean_layout(
        fig,
        height=420,
        title="Bonus Neto de Capitán por Mánager",
        top_margin=60,
        right_margin=45,
    )
    fig.update_layout(
        showlegend=False,
        xaxis=dict(title="Bonus Neto Capitán (pts)", tickfont=dict(size=12)),
        yaxis=dict(title="", tickfont=dict(size=12)),
    )
    return fig


def make_projection_chart(
    view_proj: pd.DataFrame,
    leader_pts: int,
    total_rounds: int,
) -> go.Figure:
    proj_long = view_proj.melt(
        id_vars=["Mánager"],
        value_vars=["Puntos Totales", "Proy. Media J17", "Proy. Forma J17"],
        var_name="Escenario",
        value_name="Puntos",
    )
    fig = px.bar(
        proj_long,
        x="Mánager",
        y="Puntos",
        color="Escenario",
        barmode="group",
        text_auto=".0f",
        color_discrete_map={
            "Puntos Totales": "#64748B",
            "Proy. Media J17": "#0F766E",
            "Proy. Forma J17": "#2563EB",
        },
        title=f"Puntos Actuales vs. Proyección a J{total_rounds} (Media vs. Últimas 3J)",
    )
    fig.update_traces(
        textposition="outside",
        textfont=dict(size=11),
        cliponaxis=False,
    )
    if leader_pts > 0:
        fig.add_hline(
            y=leader_pts,
            line_dash="dot",
            line_color="#DC2626",
            annotation_text=f"Líder actual ({leader_pts} pts)",
            annotation_position="top right",
            annotation_font=dict(size=11, color="#DC2626"),
        )

    apply_clean_layout(
        fig,
        height=450,
        title=f"Puntos Actuales vs. Proyección a J{total_rounds} (Media vs. Últimas 3J)",
        top_margin=70,
        right_margin=25,
    )
    fig.update_layout(
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
    return fig


def make_manager_mini_bars(
    chosen: str,
    active_rounds: list[str],
    df_pts: pd.DataFrame,
    df_cap_pts: pd.DataFrame,
    df_ga: pd.DataFrame,
    df_pac: pd.DataFrame,
    manager_color: str,
) -> tuple[go.Figure, go.Figure, go.Figure]:
    s_pts = df_pts.loc[chosen, active_rounds].rename("Puntos")
    fig_m = px.bar(
        s_pts,
        x=s_pts.index,
        y="Puntos",
        text_auto=".0f",
        title=f"Puntos por jornada — {chosen}",
        color_discrete_sequence=[manager_color],
    )
    fig_m.update_traces(textposition="outside", textfont=dict(size=11), cliponaxis=False)
    apply_clean_layout(fig_m, height=320, title=f"Puntos por jornada — {chosen}", top_margin=55)
    fig_m.update_layout(
        showlegend=False,
        xaxis=dict(title="Jornada", tickfont=dict(size=11)),
        yaxis=dict(title="Puntos", tickfont=dict(size=11), rangemode="tozero"),
    )

    s_cap = df_cap_pts.loc[chosen, active_rounds].rename("Capitán")
    fig_c = px.bar(
        s_cap,
        x=s_cap.index,
        y="Capitán",
        text_auto=".0f",
        title=f"Puntos capitán — {chosen}",
        color_discrete_sequence=["#0F766E"],
    )
    fig_c.update_traces(textposition="outside", textfont=dict(size=11), cliponaxis=False)
    apply_clean_layout(fig_c, height=320, title=f"Puntos capitán — {chosen}", top_margin=55)
    fig_c.update_layout(
        showlegend=False,
        xaxis=dict(title="Jornada", tickfont=dict(size=11)),
        yaxis=dict(title="Capitán", tickfont=dict(size=11), rangemode="tozero"),
    )

    df_of = pd.DataFrame(
        {
            "G+A": df_ga.loc[chosen, active_rounds],
            "PaC": df_pac.loc[chosen, active_rounds],
        }
    )
    fig_o = px.bar(
        df_of,
        x=df_of.index,
        y=["G+A", "PaC"],
        barmode="group",
        text_auto=".0f",
        color_discrete_map={"G+A": "#2563EB", "PaC": "#D97706"},
        title=f"Ofensiva (G+A) y PaC — {chosen}",
    )
    fig_o.update_traces(textposition="outside", textfont=dict(size=11), cliponaxis=False)
    apply_clean_layout(fig_o, height=320, title=f"Ofensiva (G+A) y PaC — {chosen}", top_margin=65)
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
    return fig_m, fig_c, fig_o
