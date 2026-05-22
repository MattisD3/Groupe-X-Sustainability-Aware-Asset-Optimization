from __future__ import annotations

from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parents[1]
PROCESSED_DIR = BASE_DIR / "data" / "processed"

MONTHLY_DATA_FILE = "B_EM_Monthly_Data.xlsx"
ANNUAL_DATA_FILE = "C_EM_Annual_Data.xlsx"
BASE_INVESTMENT_SET_FILE = "D_EM_Base_Investment_Set.xlsx"

# Je reprends ici les regles de la section 2.1.
MIN_MONTHLY_OBSERVATIONS = 36
ESTIMATION_WINDOW_YEARS = 10
FIRST_FORMATION_YEAR = 2013
LAST_FORMATION_YEAR = 2024

# Je continue les noms de sortie dans un ordre simple.
OUTPUT_FILES = {
    "investment_set": "F_MinVar_2_1_Investment_Set.xlsx",
    "expected_returns": "G_MinVar_2_1_Expected_Returns.xlsx",
    "covariances": "H_MinVar_2_1_Covariance_Matrices.xlsx",
    "summary": "I_MinVar_2_1_Summary.xlsx",
}


def log_step(message: str):
    """Je m'affiche dans le terminal pour que l'execution reste lisible."""
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


def write_covariance_workbook_with_fallback(
    covariance_matrices: dict[int, pd.DataFrame],
    file_name: str,
):
    """
    J'ecris les matrices de covariance dans un seul fichier Excel avec une feuille par annee de formation.
    """
    target_path = PROCESSED_DIR / file_name

    def write_workbook(path: Path):
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            for formation_year, covariance_matrix in covariance_matrices.items():
                covariance_matrix.to_excel(writer, sheet_name=f"Y_{formation_year}")

    try:
        write_workbook(target_path)
        return target_path
    except PermissionError:
        fallback_path = target_path.with_name(f"{target_path.stem}_new{target_path.suffix}")
        write_workbook(fallback_path)
        return fallback_path


def load_part1_outputs():
    """
    Je charge seulement les trois fichiers necessaires a la section 2.1.
    Pourquoi ces trois fichiers suffisent:
    - B contient les rendements mensuels nettoyes,
    - C contient la disponibilite des donnees carbone a la fin de chaque annee,
    - D contient deja le filtre de prix et le filtre stale prices.
    """
    monthly_data = pd.read_excel(
        PROCESSED_DIR / MONTHLY_DATA_FILE,
        parse_dates=["Date", "Delisting Date"],
    )
    annual_data = pd.read_excel(
        PROCESSED_DIR / ANNUAL_DATA_FILE,
        parse_dates=["Delisting Date"],
    )
    base_investment_set = pd.read_excel(
        PROCESSED_DIR / BASE_INVESTMENT_SET_FILE,
        parse_dates=["Delisting Date"],
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

    annual_data = annual_data.rename(columns={
        "ISIN": "isin",
        "Company Name": "company_name",
        "Country": "country",
        "Region": "region",
        "Delisting Date": "delisting_date",
        "Year": "year",
        "Scope 1 CO2": "scope1_co2",
        "Revenue Thousand USD": "revenue_thousand_usd",
        "Year End Market Value MUSD": "year_end_market_value_musd",
        "Year End Return Index": "year_end_return_index",
        "Price Available End Of Year": "price_available_eoy",
    })

    base_investment_set = base_investment_set.rename(columns={
        "ISIN": "isin",
        "Company Name": "company_name",
        "Country": "country",
        "Region": "region",
        "Delisting Date": "delisting_date",
        "Formation Year": "formation_year",
        "Investment Year": "investment_year",
        "Year End Market Value MUSD": "year_end_market_value_musd",
        "Year End Return Index": "year_end_return_index",
        "Price Available End Of Year": "price_available_eoy",
        "Valid Return Count 10Y": "valid_return_count_10y",
        "Zero Return Count 10Y": "zero_return_count_10y",
        "Zero Return Ratio 10Y": "zero_return_ratio_10y",
        "Stale Price Flag": "stale_price_flag",
        "Base Investable Next Year": "base_investable_next_year",
    })

    for df in [monthly_data, annual_data, base_investment_set]:
        df["isin"] = df["isin"].astype(str).str.strip()

    return monthly_data, annual_data, base_investment_set


def build_min_var_investment_set(
    annual_data: pd.DataFrame,
    base_investment_set: pd.DataFrame,
):
    """
    Je construis l'investment set de la section 2.1.

    A la fin de l'annee Y, je garde une entreprise si:
    - elle appartient deja a la bonne region,
    - elle est encore investissable apres les filtres de la partie 1,
    - elle a au moins 36 rendements mensuels disponibles sur les 10 dernieres annees,
    - elle a bien une donnee Scope 1 disponible a la fin de l'annee Y.
    """
    # Je garde seulement les variables annuelles dont j'ai besoin pour tester l'eligibilite carbone.
    carbon_data = annual_data[
        ["isin", "year", "scope1_co2", "revenue_thousand_usd"]
    ].copy()

    # Pour la suite du projet, j'ai besoin des emissions et du revenu.
    # Si l'un des deux manque, je considere que l'entreprise n'a pas une base carbone exploitable.
    carbon_data["has_carbon_data"] = (
        carbon_data["scope1_co2"].notna()
        & carbon_data["revenue_thousand_usd"].notna()
    )

    # Je rattache les donnees carbone de l'annee Y a l'univers de base forme a la fin de l'annee Y.
    investment_set = base_investment_set.merge(
        carbon_data,
        left_on=["isin", "formation_year"],
        right_on=["isin", "year"],
        how="left",
    )

    # Je demande au moins 36 rendements mensuels disponibles sur la fenetre de 10 ans.
    investment_set["enough_return_observations"] = (
        investment_set["valid_return_count_10y"] >= MIN_MONTHLY_OBSERVATIONS
    )

    # L'entreprise est eligible pour le minimum variance si:
    # - elle a deja passe les filtres de base de la partie 1,
    # - elle a assez de rendements mensuels,
    # - elle a des donnees carbone utilisables a la fin de l'annee Y.
    investment_set["min_var_eligible"] = (
        investment_set["base_investable_next_year"]
        & investment_set["enough_return_observations"]
        & investment_set["has_carbon_data"]
    )

    # Je garde seulement la periode demandee par le projet.
    investment_set = investment_set.loc[
        investment_set["formation_year"].between(FIRST_FORMATION_YEAR, LAST_FORMATION_YEAR)
    ].copy()

    investment_set = investment_set.sort_values(
        ["formation_year", "isin"]
    ).reset_index(drop=True)

    # Je garde uniquement les colonnes utiles pour la suite.
    selected_columns = [
        "isin",
        "company_name",
        "country",
        "region",
        "delisting_date",
        "formation_year",
        "investment_year",
        "year_end_market_value_musd",
        "year_end_return_index",
        "price_available_eoy",
        "valid_return_count_10y",
        "zero_return_count_10y",
        "zero_return_ratio_10y",
        "stale_price_flag",
        "base_investable_next_year",
        "scope1_co2",
        "revenue_thousand_usd",
        "has_carbon_data",
        "enough_return_observations",
        "min_var_eligible",
    ]

    return investment_set[selected_columns]


def compute_expected_returns(
    monthly_data: pd.DataFrame,
    investment_set: pd.DataFrame,
):
    """
    Je calcule le vecteur de rendements attendus pour chaque annee Y.

    Concretement:
    - je prends Jan(Y-9) a Dec(Y),
    - je garde seulement les entreprises eligibles,
    - je calcule la moyenne des rendements mensuels disponibles.
    """
    expected_returns_tables: list[pd.DataFrame] = []

    for formation_year in range(FIRST_FORMATION_YEAR, LAST_FORMATION_YEAR + 1):
        log_step(f"Je traite l'annee de formation {formation_year}.")
        eligible_isins = investment_set.loc[
            (investment_set["formation_year"] == formation_year)
            & (investment_set["min_var_eligible"]),
            "isin",
        ].tolist()

        if not eligible_isins:
            continue

        window_start = pd.Timestamp(formation_year - ESTIMATION_WINDOW_YEARS + 1, 1, 1)
        window_end = pd.Timestamp(formation_year, 12, 31)

        window_data = monthly_data.loc[
            (monthly_data["date"] >= window_start)
            & (monthly_data["date"] <= window_end)
            & (monthly_data["isin"].isin(eligible_isins)),
            ["isin", "monthly_return"],
        ].copy()

        mean_returns = (
            window_data.groupby("isin")["monthly_return"]
            .agg(
                used_monthly_observations=lambda values: int(values.notna().sum()),
                mean_monthly_return="mean",
            )
            .reset_index()
        )
        mean_returns["formation_year"] = formation_year
        mean_returns["investment_year"] = formation_year + 1
        expected_returns_tables.append(mean_returns)

    expected_returns = pd.concat(expected_returns_tables, ignore_index=True)
    expected_returns = expected_returns.merge(
        investment_set[
            ["isin", "company_name", "country", "formation_year", "investment_year"]
        ].drop_duplicates(),
        on=["isin", "formation_year", "investment_year"],
        how="left",
    )

    expected_returns = expected_returns[
        [
            "isin",
            "company_name",
            "country",
            "formation_year",
            "investment_year",
            "used_monthly_observations",
            "mean_monthly_return",
        ]
    ]
    expected_returns = expected_returns.sort_values(["formation_year", "isin"]).reset_index(drop=True)
    return expected_returns


def compute_covariance_matrices(
    monthly_data: pd.DataFrame,
    investment_set: pd.DataFrame,
):
    """
    Je calcule la matrice de covariance des rendements mensuels pour chaque annee Y.

    Je reprends la meme fenetre de 10 ans que pour les rendements moyens.
    """
    covariance_matrices: dict[int, pd.DataFrame] = {}

    for formation_year in range(FIRST_FORMATION_YEAR, LAST_FORMATION_YEAR + 1):
        log_step(f"Je calcule la covariance pour l'annee de formation {formation_year}.")
        eligible_isins = investment_set.loc[
            (investment_set["formation_year"] == formation_year)
            & (investment_set["min_var_eligible"]),
            "isin",
        ].tolist()

        if not eligible_isins:
            continue

        window_start = pd.Timestamp(formation_year - ESTIMATION_WINDOW_YEARS + 1, 1, 1)
        window_end = pd.Timestamp(formation_year, 12, 31)

        return_matrix = monthly_data.loc[
            (monthly_data["date"] >= window_start)
            & (monthly_data["date"] <= window_end)
            & (monthly_data["isin"].isin(eligible_isins)),
            ["date", "isin", "monthly_return"],
        ].pivot(index="date", columns="isin", values="monthly_return")

        return_matrix = return_matrix.reindex(columns=eligible_isins)
        covariance_matrices[formation_year] = return_matrix.cov(min_periods=MIN_MONTHLY_OBSERVATIONS)

    return covariance_matrices


def build_summary_table(
    investment_set: pd.DataFrame,
    expected_returns: pd.DataFrame,
    covariance_matrices: dict[int, pd.DataFrame],
):
    """
    Je prepare un petit resume clair pour verifier les sorties de la section 2.1.
    """
    summary_rows: list[dict[str, object]] = []

    for formation_year in range(FIRST_FORMATION_YEAR, LAST_FORMATION_YEAR + 1):
        eligible_firms = int(
            investment_set.loc[
                (investment_set["formation_year"] == formation_year)
                & (investment_set["min_var_eligible"]),
                "isin",
            ].nunique()
        )
        expected_return_vectors = int(
            expected_returns.loc[expected_returns["formation_year"] == formation_year, "isin"].nunique()
        )
        covariance_matrix_size = 0
        if formation_year in covariance_matrices:
            covariance_matrix_size = covariance_matrices[formation_year].shape[0]

        summary_rows.append(
            {
                "formation_year": formation_year,
                "investment_year": formation_year + 1,
                "eligible_firms": eligible_firms,
                "expected_return_vectors": expected_return_vectors,
                "covariance_matrix_size": covariance_matrix_size,
            }
        )

    return pd.DataFrame(summary_rows)


def save_outputs(
    investment_set: pd.DataFrame,
    expected_returns: pd.DataFrame,
    covariance_matrices: dict[int, pd.DataFrame],
    summary_table: pd.DataFrame,
):
    """J'enregistre les sorties finales de la section 2.1."""
    PROCESSED_DIR.mkdir(parents=True, exist_ok=True)

    written_files = {
        "F": str(write_excel_with_fallback(investment_set, OUTPUT_FILES["investment_set"])),
        "G": str(write_excel_with_fallback(expected_returns, OUTPUT_FILES["expected_returns"])),
        "H": str(write_covariance_workbook_with_fallback(covariance_matrices, OUTPUT_FILES["covariances"])),
        "I": str(write_excel_with_fallback(summary_table, OUTPUT_FILES["summary"])),
    }
    return written_files


def main():
    # Je charge d'abord uniquement ce dont j'ai besoin venant de la partie 1.
    log_step("Etape 1/4 - Je charge les sorties utiles de la partie 1...")
    monthly_data, annual_data, base_investment_set = load_part1_outputs()

    # Je construis ensuite l'investment set strict de la section 2.1.
    log_step("Etape 2/4 - Je construis l'investment set minimum-variance...")
    investment_set = build_min_var_investment_set(
        annual_data=annual_data,
        base_investment_set=base_investment_set,
    )

    # Je calcule ensuite les moments de rendement utiles a la suite.
    log_step("Etape 3/4 - Je calcule les rendements moyens et les matrices de covariance...")
    expected_returns = compute_expected_returns(
        monthly_data=monthly_data,
        investment_set=investment_set,
    )
    covariance_matrices = compute_covariance_matrices(
        monthly_data=monthly_data,
        investment_set=investment_set,
    )

    # Je termine par les sorties finales.
    log_step("Etape 4/4 - J'enregistre les sorties Excel...")
    summary_table = build_summary_table(
        investment_set=investment_set,
        expected_returns=expected_returns,
        covariance_matrices=covariance_matrices,
    )
    written_files = save_outputs(
        investment_set=investment_set,
        expected_returns=expected_returns,
        covariance_matrices=covariance_matrices,
        summary_table=summary_table,
    )

    log_step("Partie 2.1 terminee.")
    log_step(f"Lignes dans l'investment set final: {len(investment_set)}")
    log_step(f"Lignes de rendements moyens: {len(expected_returns)}")
    log_step("Fichiers ecrits:")
    for label, path in written_files.items():
        log_step(f"{label} : {path}")


if __name__ == "__main__":
    main()
