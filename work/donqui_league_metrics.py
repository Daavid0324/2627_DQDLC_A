from __future__ import annotations

import numpy as np
import pandas as pd
from .data_loader import TOTAL_ROUNDS, norm_key


def compute_league_metrics(
    raw_data: dict,
    rounds: list[str] | None = None,
) -> pd.DataFrame:
    """
    Calcula de forma vectorizada la clasificación y todas las métricas analíticas
    para el conjunto de jornadas indicado (por defecto, todas las jugadas).
    """
    played_cols: list[str] = raw_data["played_cols"]
    target_rounds = [r for r in (rounds or played_cols) if r in played_cols]

    managers: list[str] = raw_data["managers"]
    if not target_rounds:
        return pd.DataFrame({"Mánager": managers})

    df_pts: pd.DataFrame = raw_data["df_pts"][target_rounds]
    df_cap_base: pd.DataFrame = raw_data["df_cap_base"][target_rounds]
    df_cap_pts: pd.DataFrame = raw_data["df_cap_pts"][target_rounds]
    df_goals: pd.DataFrame = raw_data["df_goals"][target_rounds]
    df_assists: pd.DataFrame = raw_data["df_assists"][target_rounds]
    df_ga: pd.DataFrame = raw_data["df_ga"][target_rounds]
    df_pac: pd.DataFrame = raw_data["df_pac"][target_rounds]
    cap_extra: dict[str, dict] = raw_data["cap_extra"]

    # Rankings vectorizados por columna (jornada), conservando índice por mánager
    # FIX: evita colisiones cuando dos mánagers empatan en puntos en la misma jornada
    df_ranks = df_pts.rank(axis=0, ascending=False, method="min")

    # Jornadas restantes globales respecto al total de la liga
    total_played_season = len(played_cols)
    rem_j = max(0, TOTAL_ROUNDS - total_played_season)

    cap_extra_norm = {norm_key(k): v for k, v in cap_extra.items()}

    rows: list[dict] = []
    for m in managers:
        s_pts = df_pts.loc[m].dropna()
        n_j = int(len(s_pts))

        tot = float(s_pts.sum()) if n_j else 0.0
        avg = float(s_pts.mean()) if n_j else 0.0
        med = float(s_pts.median()) if n_j else 0.0
        sd = float(s_pts.std(ddof=1)) if n_j > 1 else 0.0
        cv = (sd / avg) if avg > 0 else 0.0

        last3_series = s_pts.iloc[-3:] if n_j >= 3 else s_pts
        last3 = float(last3_series.mean()) if n_j else 0.0
        last_j = int(round(float(s_pts.iloc[-1]))) if n_j else 0

        valid_idx = s_pts.index
        cap_tot = float(df_cap_pts.loc[m, valid_idx].sum()) if n_j else 0.0
        cap_base_tot = float(df_cap_base.loc[m, valid_idx].sum()) if n_j else 0.0
        net_bonus = int(round(cap_tot - cap_base_tot))
        no_cap = int(round(tot - net_bonus))
        pct_cap = round((cap_tot / tot * 100), 1) if tot > 0 else 0.0

        goals_tot = int(round(float(df_goals.loc[m, valid_idx].sum()))) if n_j else 0
        assists_tot = int(round(float(df_assists.loc[m, valid_idx].sum()))) if n_j else 0
        ga_tot = int(round(float(df_ga.loc[m, valid_idx].sum()))) if n_j else 0
        pac_tot = int(round(float(df_pac.loc[m, valid_idx].sum()))) if n_j else 0

        s_ranks = df_ranks.loc[m, valid_idx].dropna()
        pos_media = round(float(s_ranks.mean()), 1) if not s_ranks.empty else 0.0
        wins_j = int((s_ranks == 1).sum()) if not s_ranks.empty else 0
        top3_j = int((s_ranks <= 3).sum()) if not s_ranks.empty else 0

        extra = cap_extra.get(m) or cap_extra_norm.get(norm_key(m)) or {}

        rows.append(
            {
                "Mánager": m,
                "Jornadas": n_j,
                "Puntos Totales": int(round(tot)),
                "Últ. J": last_j,
                "Media / J": round(avg, 1),
                "Mediana / J": round(med, 1),
                "DE": round(sd, 1),
                "CV": round(cv, 2),
                "Últ. 3J": round(last3, 1),
                "Pts Cap.": int(round(cap_tot)),
                "Bonus Cap.": net_bonus,
                "Sin Cap.": no_cap,
                "% Aporte Cap.": extra.get("% Aporte Cap. Excel", pct_cap) if extra else pct_cap,
                "Veces x3": extra.get("Veces x3", 0),
                "Veces x2": extra.get("Veces x2", 0),
                "Veces x1,5": extra.get("Veces x1,5", 0),
                "Mult. Promedio": extra.get("Mult. Promedio", 0.0),
                "Cap. Top (≥5)": extra.get("Cap. Top (≥5)", 0),
                "Goles": goals_tot,
                "Asistencias": assists_tot,
                "G+A": ga_tot,
                "Pts / (G+A)": round(tot / ga_tot, 1) if ga_tot > 0 else 0.0,
                "PaC": pac_tot,
                "Pos. Media": pos_media,
                "Victorias J": wins_j,
                "Top 3 J": top3_j,
                "Proy. Media J17": int(round(tot + avg * rem_j)),
                "Proy. Forma J17": int(round(tot + last3 * rem_j)),
                "Dif. Proy. (Forma - Media)": int(round((last3 - avg) * rem_j)),
                "Mínima": int(round(float(s_pts.min()))) if n_j else 0,
                "Máxima": int(round(float(s_pts.max()))) if n_j else 0,
            }
        )

    df_cls = (
        pd.DataFrame(rows)
        .sort_values(["Puntos Totales", "Media / J", "Pos. Media"], ascending=[False, False, True])
        .reset_index(drop=True)
    )
    df_cls.insert(0, "Pos", df_cls.index + 1)
    leader_pts = int(df_cls["Puntos Totales"].iloc[0]) if not df_cls.empty else 0
    df_cls.insert(3, "Dif. Líder", df_cls["Puntos Totales"] - leader_pts)

    return df_cls


def build_manager_diagnostic(row: pd.Series, league_df: pd.DataFrame) -> dict[str, str]:
    """Genera un diagnóstico automático de fortaleza y punto de mejora para un mánager."""
    avg_league = float(league_df["Media / J"].mean())
    cv_league = float(league_df["CV"].median())
    cap_pct_league = float(league_df["% Aporte Cap."].median())

    strengths: list[str] = []
    weaknesses: list[str] = []

    if row["Media / J"] >= avg_league:
        strengths.append(f"Ritmo superior a la media de liga ({row['Media / J']} vs. {avg_league:.1f} pts/J)")
    else:
        weaknesses.append(f"Media por debajo del promedio de liga ({row['Media / J']} vs. {avg_league:.1f} pts/J)")

    if row["CV"] <= cv_league:
        strengths.append(f"Alta regularidad (CV {row['CV']} ≤ mediana {cv_league:.2f})")
    else:
        weaknesses.append(f"Alta volatilidad entre jornadas (CV {row['CV']} vs. mediana {cv_league:.2f})")

    diff_proj = int(row["Dif. Proy. (Forma - Media)"])
    if diff_proj > 0:
        strengths.append(f"Tendencia alcista (+{diff_proj} pts extra proyectados si mantiene forma 3J)")
    elif diff_proj < 0:
        weaknesses.append(f"Desaceleración reciente ({diff_proj} pts proyectados respecto a su media)")

    if row["% Aporte Cap."] > cap_pct_league + 5:
        weaknesses.append(f"Alta dependencia del capitán ({row['% Aporte Cap.']}% de sus puntos)")
    elif row["Bonus Cap."] > 0:
        strengths.append(f"Buen aprovechamiento del brazalete (+{int(row['Bonus Cap.'])} pts netos)")

    return {
        "fortaleza": " · ".join(strengths[:2]) if strengths else "Rendimiento equilibrado en todos los apartados.",
        "atencion": " · ".join(weaknesses[:2]) if weaknesses else "Sin debilidades destacadas respecto al grupo.",
    }
