from __future__ import annotations

import matplotlib.pyplot as plt
import pandas as pd

from p3_0_carbon_utils import (
    PROCESSED_DIR,
    build_caption_table,
    compute_summary_stats,
    load_carbon_inputs,
    log_step,
    write_workbook,
)


OUTPUT_FILE = "Z_Carbon_Comparison_3_4.xlsx"
FIGURE_FILE = "Z_Carbon_Comparison_3_4_Cumulative_Returns.png"


def load_annual_carbon_table(file_name: str, sheet_name: str):
    """Je charge une table annuelle carbone deja produite par une section precedente."""
    return pd.read_excel(PROCESSED_DIR / file_name, sheet_name=sheet_name)


def build_comparison_table(data: dict[str, object]):
    """Je rassemble les statistiques financieres et carbone des quatre portefeuilles."""
    risk_free_series = data["risk_free"].set_index("date")["rf_decimal"]
    section_31_annual = load_annual_carbon_table("P_Carbon_3_1_WACI_CF.xlsx", "Annual Metrics")

    performance_map = {
        "mv_oos": data["mv_performance"],
        "vw": data["vw_performance"],
        "mv_carbon": pd.read_excel(PROCESSED_DIR / "R_MinVar_Carbon_3_2_Monthly_Performance.xlsx", sheet_name="Monthly Performance"),
        "vw_carbon": pd.read_excel(PROCESSED_DIR / "U_TrackingError_Carbon_3_3_Monthly_Performance.xlsx", sheet_name="Monthly Performance"),
    }

    annual_carbon_map = {
        "mv_oos": section_31_annual.query("portfolio == 'mv_oos'"),
        "vw": section_31_annual.query("portfolio == 'vw'"),
        "mv_carbon": load_annual_carbon_table("S_MinVar_Carbon_3_2_Summary.xlsx", "Annual Carbon"),
        "vw_carbon": load_annual_carbon_table("V_TrackingError_Carbon_3_3_Summary.xlsx", "Annual Carbon"),
    }

    rows: list[dict[str, object]] = []
    for portfolio_name, performance_df in performance_map.items():
        performance_df = performance_df.copy()
        performance_df["date"] = pd.to_datetime(performance_df["date"])
        stats = compute_summary_stats(performance_df.set_index("date")["portfolio_return"], risk_free_series)
        annual_carbon = annual_carbon_map[portfolio_name].copy()

        rows.append(
            {
                "portfolio": portfolio_name,
                "annualized_return": stats["annualized_return"],
                "annualized_volatility": stats["annualized_volatility"],
                "sharpe_ratio": stats["sharpe_ratio"],
                "min_monthly_return": stats["min_monthly_return"],
                "max_monthly_return": stats["max_monthly_return"],
                "average_annual_waci": annual_carbon["waci"].mean(),
                "average_annual_cf": annual_carbon["cf"].mean(),
            }
        )

    return pd.DataFrame(rows)


def build_comment_table(comparison_table: pd.DataFrame):
    """Je formule les commentaires attendus par la section 3.4."""
    comparison_by_portfolio = comparison_table.set_index("portfolio")
    mv_base = comparison_by_portfolio.loc["mv_oos"]
    mv_carbon = comparison_by_portfolio.loc["mv_carbon"]
    vw_base = comparison_by_portfolio.loc["vw"]
    vw_carbon = comparison_by_portfolio.loc["vw_carbon"]

    return pd.DataFrame(
        [
            {
                "topic": "Financial cost for the active investor",
                "comment": (
                    "The active investor's carbon constraint changes annualized return by "
                    f"{mv_carbon['annualized_return'] - mv_base['annualized_return']:.4f} and changes Sharpe ratio by "
                    f"{mv_carbon['sharpe_ratio'] - mv_base['sharpe_ratio']:.4f} relative to P(mv)_oos."
                ),
            },
            {
                "topic": "Financial cost for the passive investor",
                "comment": (
                    "The passive investor's carbon constraint changes annualized return by "
                    f"{vw_carbon['annualized_return'] - vw_base['annualized_return']:.4f} and changes Sharpe ratio by "
                    f"{vw_carbon['sharpe_ratio'] - vw_base['sharpe_ratio']:.4f} relative to P(vw)."
                ),
            },
            {
                "topic": "Difference in carbon-reduction mechanism",
                "comment": (
                    "P(mv)_oos(0.5) reduces carbon through a variance-minimizing reallocation under a direct footprint cap, "
                    "whereas P(vw)_oos(0.5) stays close to the passive benchmark and reduces carbon mainly by small active tilts around market-cap weights."
                ),
            },
        ]
    )


def save_cumulative_figure(mv_performance, mv_carbon_performance, vw_performance, vw_carbon_performance):
    """Je sauvegarde la figure commune des quatre portefeuilles de la section 3.4."""
    figure_path = PROCESSED_DIR / FIGURE_FILE
    fig, ax = plt.subplots(figsize=(12, 6), dpi=150)
    ax.plot(mv_performance["date"], mv_performance["cumulative_growth"], color="darkorange", label="P(mv)_oos")
    ax.plot(mv_carbon_performance["date"], mv_carbon_performance["cumulative_growth"], color="firebrick", label="P(mv)_oos(0.5)")
    ax.plot(vw_performance["date"], vw_performance["cumulative_growth"], color="steelblue", label="P(vw)")
    ax.plot(vw_carbon_performance["date"], vw_carbon_performance["cumulative_growth"], color="seagreen", label="P(vw)_oos(0.5)")
    ax.set_title("Cumulative Returns of the Four Portfolios")
    ax.set_xlabel("Date")
    ax.set_ylabel("Cumulative Growth")
    ax.legend()
    ax.grid(True)
    fig.tight_layout()
    fig.savefig(figure_path, bbox_inches="tight")
    plt.close(fig)


def main():
    """Je compare les quatre portefeuilles de la section 3.4."""
    log_step("Section 3.4 - Etape 1/3 - Je charge les sorties des sections precedentes...")
    data = load_carbon_inputs()
    mv_carbon_performance = pd.read_excel(PROCESSED_DIR / "R_MinVar_Carbon_3_2_Monthly_Performance.xlsx", sheet_name="Monthly Performance", parse_dates=["date"])
    vw_carbon_performance = pd.read_excel(PROCESSED_DIR / "U_TrackingError_Carbon_3_3_Monthly_Performance.xlsx", sheet_name="Monthly Performance", parse_dates=["date"])

    log_step("Section 3.4 - Etape 2/3 - Je construis la table comparative et les commentaires...")
    comparison_table = build_comparison_table(data)
    comments_table = build_comment_table(comparison_table)
    captions = build_caption_table(
        [
            {
                "item": "Comparison Table",
                "caption": "This table compares the four portfolios on annualized financial performance and average annual carbon metrics over the common out-of-sample window from January 2014 to December 2025.",
            },
            {
                "item": "Comments",
                "caption": "These comments interpret the financial cost of the 50% carbon-footprint constraint for the active and passive investors, and explain how the two optimization problems reduce carbon differently.",
            },
            {
                "item": "Figure",
                "caption": "The cumulative return figure plots the four portfolios on the same scale, starting from 1 at the beginning of January 2014.",
            },
        ]
    )

    log_step("Section 3.4 - Etape 3/3 - J'enregistre les sorties...")
    workbook_path = write_workbook(
        OUTPUT_FILE,
        {
            "Comparison": comparison_table,
            "Comments": comments_table,
            "Captions": captions,
        },
    )
    save_cumulative_figure(
        data["mv_performance"],
        mv_carbon_performance,
        data["vw_performance"],
        vw_carbon_performance,
    )

    log_step(f"Fichier comparatif ecrit: {workbook_path}")
    print("Section 3.4 complete.", flush=True)


if __name__ == "__main__":
    main()
