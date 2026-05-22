from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd
from scipy.optimize import minimize


# Je centralise les chemins pour garder un script simple.
BASE_DIR = Path(__file__).resolve().parents[1]
PROCESSED_DIR = BASE_DIR / "data" / "processed"
RAW_DIR = BASE_DIR / "data" / "Raw"


# Je recharge seulement les fichiers utiles a la section 2.2.
MONTHLY_DATA_FILE = "B_EM_Monthly_Data.xlsx"
INVESTMENT_SET_FILE = "F_MinVar_2_1_Investment_Set.xlsx"
COVARIANCE_FILE = "H_MinVar_2_1_Covariance_Matrices.xlsx"
RISK_FREE_FILE = "Risk_Free_Rate_2025.xlsx"


# Je borne l'exercice exactement comme dans le document.
FIRST_FORMATION_YEAR = 2013
LAST_FORMATION_YEAR = 2024


# Je nomme les sorties en continuant la serie.
OUTPUT_FILES = {
    "weights": "J_MinVar_2_2_Weights.xlsx",
    "monthly_returns": "K_MinVar_2_2_Monthly_Performance.xlsx",
    "summary": "L_MinVar_2_2_Summary.xlsx",
}


def log_step(message: str):
    """Je m'affiche dans le terminal pour rendre l'execution lisible."""
    print(message, flush=True)


def write_excel_with_fallback(df: pd.DataFrame, file_name: str):
    """J'ecris un Excel, ou une version _new si le fichier est deja ouvert."""
    target_path = PROCESSED_DIR / file_name

    try:
        df.to_excel(target_path, index=False)
        return target_path
    except PermissionError:
        fallback_path = target_path.with_name(f"{target_path.stem}_new{target_path.suffix}")
        df.to_excel(fallback_path, index=False)
        return fallback_path


def load_inputs():
    """
    Je charge uniquement les sorties utiles de la section 2.1 et le taux sans risque brut.
    """
    monthly_data = pd.read_excel(
        PROCESSED_DIR / MONTHLY_DATA_FILE,
        parse_dates=["Date", "Delisting Date"],
    )
    investment_set = pd.read_excel(
        PROCESSED_DIR / INVESTMENT_SET_FILE,
        parse_dates=["delisting_date"],
    )

    monthly_data = monthly_data.rename(columns={
        "ISIN": "isin",
        "Company Name": "company_name",
        "Country": "country",
        "Region": "region",
        "Delisting Date": "delisting_date",
        "Date": "date",
        "Market Value MUSD": "market_value_musd",
        "Return Index": "return_index",
        "Monthly Return": "monthly_return",
        "Is Delisting Month": "is_delisting_month",
    })

    for df in [monthly_data, investment_set]:
        df["isin"] = df["isin"].astype(str).str.strip()

    covariance_workbook = pd.ExcelFile(PROCESSED_DIR / COVARIANCE_FILE)
    covariance_matrices: dict[int, pd.DataFrame] = {}
    for sheet_name in covariance_workbook.sheet_names:
        formation_year = int(sheet_name.replace("Y_", ""))
        covariance_matrix = pd.read_excel(
            PROCESSED_DIR / COVARIANCE_FILE,
            sheet_name=sheet_name,
            index_col=0,
        )
        covariance_matrix.index = covariance_matrix.index.astype(str).str.strip()
        covariance_matrix.columns = covariance_matrix.columns.astype(str).str.strip()
        covariance_matrices[formation_year] = covariance_matrix

    risk_free_rate = pd.read_excel(RAW_DIR / RISK_FREE_FILE)
    first_column = risk_free_rate.columns[0]
    risk_free_rate = risk_free_rate.rename(columns={first_column: "yyyymm", "RF": "rf_percent"})
    risk_free_rate["yyyymm"] = risk_free_rate["yyyymm"].astype(str).str.strip()
    risk_free_rate["date"] = pd.to_datetime(risk_free_rate["yyyymm"] + "01", format="%Y%m%d")
    risk_free_rate["date"] = risk_free_rate["date"] + pd.offsets.MonthEnd(0)
    risk_free_rate["rf_decimal"] = pd.to_numeric(risk_free_rate["rf_percent"], errors="coerce") / 100
    risk_free_rate = risk_free_rate[["date", "rf_decimal"]].copy()

    return monthly_data, investment_set, covariance_matrices, risk_free_rate


def solve_long_only_min_variance(covariance_matrix: pd.DataFrame):
    """
    Je resous le probleme minimum-variance long-only.

    Je cherche les poids alpha qui minimisent:
    alpha' Sigma alpha
    sous les contraintes:
    - somme des poids = 1
    - poids >= 0
    """
    asset_names = covariance_matrix.columns.tolist()
    sigma = covariance_matrix.to_numpy(dtype=float)
    asset_count = len(asset_names)
    ones_vector = np.ones(asset_count, dtype=float)

    initial_weights = np.repeat(1 / asset_count, asset_count)
    bounds = [(0.0, 1.0)] * asset_count
    constraints = [
        {
            "type": "eq",
            "fun": lambda weights: float(np.sum(weights) - 1.0),
            "jac": lambda weights: ones_vector,
        }
    ]

    def objective(weights: np.ndarray):
        return float(weights.T @ sigma @ weights)

    def objective_gradient(weights: np.ndarray):
        return 2.0 * sigma @ weights

    optimization = minimize(
        objective,
        jac=objective_gradient,
        x0=initial_weights,
        method="SLSQP",
        bounds=bounds,
        constraints=constraints,
        options={"maxiter": 1000, "ftol": 1e-12},
    )

    if not optimization.success:
        raise RuntimeError(f"Optimization failed: {optimization.message}")

    weights = pd.Series(optimization.x, index=asset_names, name="weight")
    weights = weights.clip(lower=0)
    weights = weights / weights.sum()
    return weights


def build_optimal_weights(
    investment_set: pd.DataFrame,
    covariance_matrices: dict[int, pd.DataFrame],
):
    """
    Je calcule les poids optimaux a la fin de chaque annee de formation.

    Je pars des firmes eligibles de 2.1. Ensuite, pour l'optimisation elle-meme,
    je retire les firmes qui ont encore une ligne/colonne incomplete dans la matrice
    de covariance de l'annee consideree.
    """
    all_weights: list[pd.DataFrame] = []
    optimizer_eligible_counts: list[dict[str, int]] = []

    for formation_year in range(FIRST_FORMATION_YEAR, LAST_FORMATION_YEAR + 1):
        log_step(f"Je traite l'annee de formation {formation_year}.")

        year_investment_set = investment_set.loc[
            (investment_set["formation_year"] == formation_year)
            & (investment_set["min_var_eligible"]),
        ].copy()

        if formation_year not in covariance_matrices or year_investment_set.empty:
            continue

        covariance_matrix = covariance_matrices[formation_year].copy()
        eligible_isins = year_investment_set["isin"].astype(str).tolist()
        covariance_matrix = covariance_matrix.loc[eligible_isins, eligible_isins]

        valid_assets = covariance_matrix.index[
            covariance_matrix.notna().all(axis=1) & covariance_matrix.notna().all(axis=0)
        ].tolist()
        covariance_matrix = covariance_matrix.loc[valid_assets, valid_assets]

        if covariance_matrix.empty:
            continue

        company_info = year_investment_set[
            ["isin", "company_name", "country", "region"]
        ].drop_duplicates()

        log_step(f"Annee {formation_year} : optimisation sur {len(covariance_matrix)} actions.")
        optimal_weights = solve_long_only_min_variance(covariance_matrix)

        weights_df = optimal_weights.reset_index()
        weights_df.columns = ["isin", "weight"]
        weights_df["formation_year"] = formation_year
        weights_df["investment_year"] = formation_year + 1

        weights_df = weights_df.merge(company_info, on="isin", how="left")

        weights_df = weights_df[
            [
                "isin",
                "company_name",
                "country",
                "region",
                "formation_year",
                "investment_year",
                "weight",
            ]
        ]
        weights_df = weights_df.sort_values("weight", ascending=False).reset_index(drop=True)
        all_weights.append(weights_df)

        optimizer_eligible_counts.append(
            {
                "formation_year": formation_year,
                "eligible_after_covariance_cleanup": len(covariance_matrix),
            }
        )

    weights_table = pd.concat(all_weights, ignore_index=True)
    optimizer_stats = pd.DataFrame(optimizer_eligible_counts)

    optimizer_stats["formation_years_with_weights"] = int(weights_table["formation_year"].nunique())
    optimizer_stats["total_weight_rows"] = len(weights_table)
    return weights_table, optimizer_stats


def build_monthly_return_matrix(monthly_data: pd.DataFrame):
    """
    Je pivote les rendements mensuels pour pouvoir calculer les rendements de portefeuille.
    """
    return_matrix = monthly_data.pivot(index="date", columns="isin", values="monthly_return")
    return_matrix.columns = return_matrix.columns.astype(str)
    return_matrix = return_matrix.sort_index()
    return return_matrix


def compute_ex_post_performance(
    return_matrix: pd.DataFrame,
    weights_table: pd.DataFrame,
):
    """
    Je calcule la performance ex post du portefeuille minimum-variance.

    Pour chaque annee Y:
    - j'utilise les poids optimaux determines fin Y,
    - je les applique de janvier Y+1 a decembre Y+1,
    - je laisse ensuite les poids deriver naturellement mois par mois.
    """
    portfolio_rows: list[dict[str, object]] = []
    date_lookup = (
        return_matrix.index.to_series(index=return_matrix.index)
        .groupby([return_matrix.index.year, return_matrix.index.month])
        .max()
    )

    for formation_year in range(FIRST_FORMATION_YEAR, LAST_FORMATION_YEAR + 1):
        year_weights = weights_table.loc[weights_table["formation_year"] == formation_year].copy()
        if year_weights.empty:
            continue

        current_weights = year_weights.set_index("isin")["weight"].astype(float).copy()

        investment_year = formation_year + 1
        investment_dates = pd.date_range(
            start=pd.Timestamp(investment_year, 1, 31),
            end=pd.Timestamp(investment_year, 12, 31),
            freq="ME",
        )

        for calendar_date in investment_dates:
            actual_date = date_lookup.get((calendar_date.year, calendar_date.month))
            if pd.isna(actual_date):
                continue

            month_returns = return_matrix.loc[actual_date, current_weights.index]

            # Les rendements manquants sont traites comme 0 a l'interieur du portefeuille deja investi.
            # Cela permet de suivre proprement les poids apres radiation ou disparition de la serie.
            month_returns = month_returns.fillna(0.0)

            portfolio_return = float((current_weights * month_returns).sum())

            portfolio_rows.append(
                {
                    "date": actual_date,
                    "formation_year": formation_year,
                    "investment_year": investment_year,
                    "portfolio_return": portfolio_return,
                }
            )

            gross_asset_returns = 1.0 + month_returns
            gross_portfolio_return = 1.0 + portfolio_return

            if gross_portfolio_return == 0:
                current_weights = current_weights * 0.0
            else:
                current_weights = current_weights * gross_asset_returns / gross_portfolio_return
                weight_sum = current_weights.sum()
                if weight_sum > 0:
                    current_weights = current_weights / weight_sum

    performance = pd.DataFrame(portfolio_rows)
    performance = performance.sort_values("date").reset_index(drop=True)
    performance["cumulative_growth"] = (1 + performance["portfolio_return"]).cumprod()
    return performance


def compute_summary_statistics(
    portfolio_returns: pd.DataFrame,
    risk_free_rate: pd.DataFrame,
    optimizer_stats: pd.DataFrame,
):
    """
    Je calcule les statistiques demandees pour la section 2.2.
    """
    merged = portfolio_returns.merge(risk_free_rate, on="date", how="left")
    merged["excess_return"] = merged["portfolio_return"] - merged["rf_decimal"]

    annualized_average_return = merged["portfolio_return"].mean() * 12
    annualized_volatility = merged["portfolio_return"].std(ddof=1) * np.sqrt(12)

    if annualized_volatility == 0 or pd.isna(annualized_volatility):
        sharpe_ratio = np.nan
    else:
        sharpe_ratio = (merged["excess_return"].mean() * 12) / annualized_volatility

    summary_rows = [
        {"metric": "monthly_observations", "value": len(merged)},
        {"metric": "annualized_average_return", "value": annualized_average_return},
        {"metric": "annualized_volatility", "value": annualized_volatility},
        {"metric": "sharpe_ratio", "value": sharpe_ratio},
        {"metric": "minimum_monthly_return", "value": merged["portfolio_return"].min()},
        {"metric": "maximum_monthly_return", "value": merged["portfolio_return"].max()},
        {
            "metric": "final_cumulative_growth",
            "value": merged["cumulative_growth"].iloc[-1] if not merged.empty else np.nan,
        },
    ]

    if not optimizer_stats.empty:
        per_year_stats = optimizer_stats.loc[
            :, ["formation_year", "eligible_after_covariance_cleanup"]
        ].copy()
        per_year_stats["metric"] = per_year_stats["formation_year"].apply(
            lambda year: f"eligible_after_covariance_cleanup_{year}"
        )
        per_year_stats = per_year_stats.rename(columns={"eligible_after_covariance_cleanup": "value"})
        summary_rows.extend(per_year_stats[["metric", "value"]].to_dict(orient="records"))

    return pd.DataFrame(summary_rows)


def save_outputs(
    weights_table: pd.DataFrame,
    portfolio_returns: pd.DataFrame,
    summary_table: pd.DataFrame,
):
    """J'enregistre les sorties finales de la section 2.2."""
    PROCESSED_DIR.mkdir(parents=True, exist_ok=True)

    written_files = {
        "J": str(write_excel_with_fallback(weights_table, OUTPUT_FILES["weights"])),
        "K": str(write_excel_with_fallback(portfolio_returns, OUTPUT_FILES["monthly_returns"])),
        "L": str(write_excel_with_fallback(summary_table, OUTPUT_FILES["summary"])),
    }
    return written_files


def main():
    # Je charge d'abord les sorties propres de 2.1 et le taux sans risque.
    log_step("Etape 1/4 - Je charge les donnees utiles de la section 2.1...")
    monthly_data, investment_set, covariance_matrices, risk_free_rate = load_inputs()

    # Je calcule ensuite les poids optimaux minimum-variance long-only.
    log_step("Etape 2/4 - Je calcule les poids minimum-variance long-only...")
    weights_table, optimizer_stats = build_optimal_weights(
        investment_set=investment_set,
        covariance_matrices=covariance_matrices,
    )

    # Je calcule maintenant la performance ex post mensuelle du portefeuille.
    log_step("Etape 3/4 - Je calcule les rendements ex post mensuels du portefeuille...")
    return_matrix = build_monthly_return_matrix(monthly_data)
    portfolio_returns = compute_ex_post_performance(
        return_matrix=return_matrix,
        weights_table=weights_table,
    )

    # Je termine par les statistiques et les sorties.
    log_step("Etape 4/4 - J'enregistre les fichiers finaux...")
    summary_table = compute_summary_statistics(
        portfolio_returns=portfolio_returns,
        risk_free_rate=risk_free_rate,
        optimizer_stats=optimizer_stats,
    )
    written_files = save_outputs(
        weights_table=weights_table,
        portfolio_returns=portfolio_returns,
        summary_table=summary_table,
    )

    log_step("Partie 2.2 terminee.")
    log_step(f"Nombre total de lignes de poids: {len(weights_table)}")
    log_step(f"Nombre total de rendements mensuels de portefeuille: {len(portfolio_returns)}")
    log_step("Fichiers ecrits:")
    for label, path in written_files.items():
        log_step(f"{label} : {path}")


if __name__ == "__main__":
    main()
