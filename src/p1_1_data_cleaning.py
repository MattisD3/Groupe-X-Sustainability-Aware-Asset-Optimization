from __future__ import annotations
from datetime import datetime
from pathlib import Path
import pandas as pd

BASE_DIR = Path(__file__).resolve().parents[1]
RAW_DIR = BASE_DIR / "data" / "Raw"
PROCESSED_DIR = BASE_DIR / "data" / "processed"

# Je garde seulement les fichiers utiles pour la partie 1
STATIC_FILE = "Static_2025.xlsx"
SCOPE1_FILE = "DS_CO2_SCOPE_1_Y_2025.xlsx"
REVENUE_FILE = "DS_REV_Y_2025.xlsx"
MARKET_VALUE_MONTHLY_FILE = "DS_MV_T_USD_M_2025.xlsx"
RETURN_INDEX_MONTHLY_FILE = "DS_RI_T_USD_M_2025.xlsx"

# les regles importantes du document
LOW_PRICE_THRESHOLD = 0.5
STALE_PRICE_THRESHOLD = 0.5
ESTIMATION_WINDOW_YEARS = 10
FIRST_FORMATION_YEAR = 2013
LAST_FORMATION_YEAR = 2024

# Je nomme les sorties dans un ordre logique
OUTPUT_FILES = {
    "companies": "A_EM_Companies.xlsx",
    "monthly_data": "B_EM_Monthly_Data.xlsx",
    "annual_data": "C_EM_Annual_Data.xlsx",
    "base_investment_set": "D_EM_Base_Investment_Set.xlsx"}

# je les rends plus lisibles au moment de l'export
EXPORT_COLUMN_NAMES = {
    "isin": "ISIN",
    "company_name": "Company Name",
    "country": "Country",
    "region": "Region",
    "delisting_date": "Delisting Date",
    "date": "Date",
    "market_value_musd": "Market Value MUSD",
    "return_index": "Return Index",
    "monthly_return": "Monthly Return",
    "is_delisting_month": "Is Delisting Month",
    "year": "Year",
    "scope1_co2": "Scope 1 CO2",
    "revenue_thousand_usd": "Revenue Thousand USD",
    "year_end_market_value_musd": "Year End Market Value MUSD",
    "year_end_return_index": "Year End Return Index",
    "price_available_eoy": "Price Available End Of Year",
    "formation_year": "Formation Year",
    "investment_year": "Investment Year",
    "valid_return_count_10y": "Valid Return Count 10Y",
    "zero_return_count_10y": "Zero Return Count 10Y",
    "zero_return_ratio_10y": "Zero Return Ratio 10Y",
    "stale_price_flag": "Stale Price Flag",
    "base_investable_next_year": "Base Investable Next Year"}

# IA input 
def log_step(message: str):
    """Je m'affiche dans le terminal pour montrer clairement l'avancement."""
    print(message, flush=True)

def write_excel(df: pd.DataFrame, file_name: str): # J'enregistre un DataFrame dans le dossier processed avec un nom de fichier donne
    target_path = PROCESSED_DIR / file_name
    df.to_excel(target_path, index=False)
    return target_path


def rename_columns_for_export(df: pd.DataFrame):
    """
    Je rends les noms de colonnes plus clairs dans les fichiers Excel.
    """
    renamed_df = df.copy()  # Je cree une copie pour ne pas changer les noms utilises dans le reste du script.
    renamed_df = renamed_df.rename(columns=EXPORT_COLUMN_NAMES)  # Je remplace les noms techniques par des titres plus lisibles.
    return renamed_df  # Je renvoie la version prete pour l'export Excel.


def extract_delisting_date(company_name: str):
    """
    Je cherche une date de delisting dans le nom Datastream.
    Datastream ajoute souvent quelque chose comme:
    DEAD - DELIST.23/09/25
    """
    if not isinstance(company_name, str):  # Je verifie que j'ai bien un texte a lire.
        return pd.NaT  # Si ce n'est pas du texte, je considere qu'il n'y a pas de date.

    marker = "DEAD - DELIST."  # Je definis le texte qui indique un delisting dans Datastream.

    if marker not in company_name:  # Je teste si ce mot-cle est present dans le nom.
        return pd.NaT  # Si je ne le trouve pas, je laisse la date vide.

    date_text = company_name.split(marker, 1)[1][:8]  # Je recupere les 8 caracteres qui suivent, donc la date.
    return pd.to_datetime(date_text, format="%d/%m/%y", errors="coerce")  # Je transforme ce texte en vraie date.


def load_datastream_file(file_name: str):
    """
    Je charge un export Datastream brut.

    Ce que je fais tout de suite:
    - je renomme les colonnes utiles,
    - je retire la ligne parasite du haut sans ISIN
    - je force les colonnes de valeurs en numerique pour eviter les mauvaises surprises plus tard
    """
    df = pd.read_excel(RAW_DIR / file_name)
    df = df.rename(columns={"NAME": "company_name_raw", "ISIN": "isin"})
    df = df.dropna(subset=["isin"]).copy()
    df["isin"] = df["isin"].astype(str).str.strip()

    value_columns = [col for col in df.columns if col not in ["company_name_raw", "isin"]]
    for column in value_columns:
        df[column] = pd.to_numeric(df[column], errors="coerce")

    return df



def load_em_companies():
    """
    Je pars du fichier statique et je garde seulement les entreprises EM
    """
    static_df = pd.read_excel(RAW_DIR / STATIC_FILE)
    static_df = static_df.rename(
        columns={
            "ISIN": "isin",
            "NAME": "company_name",
            "Country": "country",
            "Region": "region"})

    static_df = static_df.dropna(subset=["isin"]).drop_duplicates(subset="isin")
    static_df["isin"] = static_df["isin"].astype(str).str.strip() # IA input
    static_df["region"] = static_df["region"].astype(str).str.strip() # IA input
    static_df["delisting_date"] = static_df["company_name"].apply(extract_delisting_date)

    em_companies = static_df.loc[static_df["region"] == "EM"].copy()
    em_companies = em_companies.sort_values("isin").reset_index(drop=True)
    return em_companies


def keep_only_common_isins(em_companies: pd.DataFrame):
    """
    Je retire les entreprises qui manquent completement dans au moins une table utile : 
        - si la ligne complete est absente, je supprime l'ISIN de toutes les tables.
    """
    common_isins = set(em_companies["isin"])

    for file_name in [SCOPE1_FILE, REVENUE_FILE, MARKET_VALUE_MONTHLY_FILE, RETURN_INDEX_MONTHLY_FILE]:
        file_df = load_datastream_file(file_name)
        common_isins &= set(file_df["isin"])

    filtered_companies = em_companies.loc[em_companies["isin"].isin(common_isins)].copy()
    filtered_companies = filtered_companies.sort_values("isin").reset_index(drop=True)

    return filtered_companies


def fill_annual_missing_with_previous_value(
    wide_df: pd.DataFrame,
    year_columns: list[int]):
    """
    Je remplis les valeurs annuelles manquantes avec la valeur precedente a partir du
    premier point observe: 
    - ll faut remplir les trous au milieu,
    - et aussi les trous en fin d'echantillon pour le carbone et le revenu,
    - mais je laisse les trous au debut, car la firme n'est peut-etre pas encore investissable.
    """
    cleaned_df = wide_df.copy()  # Je travaille sur une copie pour garder le fichier d'origine intact.
    filled_cells = 0  # Je compte combien de cases j'ai remplies.

    for row_index in cleaned_df.index:  # Je traite une entreprise a la fois.
        previous_value = pd.NA  # Je memorise ici la derniere valeur connue.
        seen_first_value = False  # Je distingue ce qui se passe avant et apres la premiere vraie valeur.

        for year in year_columns:  # Je parcours les annees dans l'ordre.
            current_value = cleaned_df.at[row_index, year]  # Je lis la valeur de l'annee courante.

            if pd.notna(current_value):  # Si j'ai une vraie valeur,
                previous_value = current_value  # je la garde en memoire,
                seen_first_value = True  # et je sais que la serie a commence.
                continue  # Je passe alors a l'annee suivante.

            if seen_first_value and pd.notna(previous_value):  # Si la serie a deja commence et qu'il manque une valeur,
                cleaned_df.at[row_index, year] = previous_value  # je remplace le trou par la valeur precedente.
                filled_cells += 1  # Je compte cette case remplie.

    return cleaned_df, filled_cells  # Je renvoie la table nettoyee et le nombre de trous remplis.


def find_matching_month_column(
    date_columns: list[pd.Timestamp | datetime],
    delisting_date: pd.Timestamp):
    """
    Je retrouve le dernier jour de bourse disponible du mois de delisting.
    Car le Datastream travaille souvent avec la fin de mois de bourse, pas forcement la fin de mois calendaire.
    """
    matching_columns = [
        pd.Timestamp(column)
        for column in date_columns
        if pd.Timestamp(column).year == delisting_date.year
        and pd.Timestamp(column).month == delisting_date.month]

    if not matching_columns:
        return None

    return max(matching_columns)


def build_monthly_data(em_companies: pd.DataFrame):
    """
    Je nettoie les prix mensuels et je calcule les rendements mensuels: 
    - je garde seulement les entreprises EM retenues,
    - je traite les RI < 0.5 comme des valeurs manquantes,
    - je mets le prix a 0 au mois du delisting,
    - je mets les mois suivants a manquant,
    - je calcule les simple returns mensuels a partir du RI.
    """
    market_value_wide = load_datastream_file(MARKET_VALUE_MONTHLY_FILE) # Je charge le fichier Datastream des return indexes mensuels
    return_index_wide = load_datastream_file(RETURN_INDEX_MONTHLY_FILE)

    market_value_wide = em_companies.merge(market_value_wide, on="isin", how="left") # Je rattache les market values mensuelles a mon univers EM
    return_index_wide = em_companies.merge(return_index_wide, on="isin", how="left")

    market_value_columns = sorted( # Je repere et je trie les colonnes mensuelles de market value
        [column for column in market_value_wide.columns if isinstance(column, (datetime, pd.Timestamp))]) # Je garde seulement les colonnes qui sont des dates
    return_index_columns = sorted(
        [column for column in return_index_wide.columns if isinstance(column, (datetime, pd.Timestamp))])

    low_price_mask = return_index_wide[return_index_columns].lt(LOW_PRICE_THRESHOLD)  # Je repere les return indexes strictement inferieurs au seuil 0.5.
    low_price_mask = low_price_mask & return_index_wide[return_index_columns].notna() # Je verifie que je ne marque pas des valeurs deja manquantes.
    return_index_wide.loc[:, return_index_columns] = return_index_wide[return_index_columns].mask(low_price_mask)  # Je remplace ces prix trop faibles par des valeurs manquantes.
    
    for row_index in em_companies.index: # cas des entreprises delistées: je mets le prix a 0 au mois du delisting et les mois suivants a manquant
        delisting_date = em_companies.at[row_index, "delisting_date"]
        if pd.isna(delisting_date):
            continue

        market_value_month = find_matching_month_column(market_value_columns, pd.Timestamp(delisting_date))
        return_index_month = find_matching_month_column(return_index_columns, pd.Timestamp(delisting_date))

        if market_value_month is None or return_index_month is None:
            continue

        later_market_value_columns = [column for column in market_value_columns if pd.Timestamp(column) > market_value_month]
        later_return_index_columns = [column for column in return_index_columns if pd.Timestamp(column) > return_index_month]

        market_value_wide.loc[row_index, later_market_value_columns] = pd.NA
        return_index_wide.loc[row_index, later_return_index_columns] = pd.NA

        market_value_wide.at[row_index, market_value_month] = 0.0
        return_index_wide.at[row_index, return_index_month] = 0.0

    id_columns = ["isin", "company_name", "country", "region", "delisting_date"]


    # IA input: on transforme les tableaux pour qu’ils deviennent exploitables analytiquement
    market_value_long = market_value_wide.melt(
        id_vars=id_columns, value_vars=market_value_columns, var_name="date", value_name="market_value_musd")
    
    return_index_long = return_index_wide.melt(
        id_vars=id_columns, value_vars=return_index_columns, var_name="date", value_name="return_index")

    monthly_data = market_value_long.merge(
        return_index_long[["isin", "date", "return_index"]], on=["isin", "date"], how="outer")


    # Mettre les dates au bon format, calculer les rendements mensuels, et trier les donnees pour la suite
    monthly_data["date"] = pd.to_datetime(monthly_data["date"])
    monthly_data = monthly_data.sort_values(["isin", "date"]).reset_index(drop=True)

    monthly_data["return_index_lag"] = monthly_data.groupby("isin")["return_index"].shift(1)
    monthly_data["monthly_return"] = (monthly_data["return_index"] / monthly_data["return_index_lag"] - 1)

    invalid_return_mask = monthly_data["return_index"].isna() | monthly_data["return_index_lag"].isna()
    monthly_data.loc[invalid_return_mask, "monthly_return"] = pd.NA

    monthly_data["is_delisting_month"] = (
        monthly_data["delisting_date"].notna()
        & (monthly_data["date"].dt.to_period("M") == monthly_data["delisting_date"].dt.to_period("M"))
        & monthly_data["return_index"].eq(0))

    monthly_data = monthly_data[
        [
            "isin",
            "company_name",
            "country",
            "region",
            "delisting_date",
            "date",
        "market_value_musd",
            "return_index",
            "monthly_return",
            "is_delisting_month"]]

    return monthly_data


def build_annual_data(em_companies: pd.DataFrame, monthly_data: pd.DataFrame):
    """
    Je garde uniquement:
    - Scope 1,
    - Revenue,
    - prix de fin d'annee nettoyes a partir du mensuel.
    """
    scope1_wide = load_datastream_file(SCOPE1_FILE)
    revenue_wide = load_datastream_file(REVENUE_FILE)

    scope1_wide = scope1_wide.loc[scope1_wide["isin"].isin(em_companies["isin"])].copy()
    revenue_wide = revenue_wide.loc[revenue_wide["isin"].isin(em_companies["isin"])].copy()

    scope1_year_columns = sorted([column for column in scope1_wide.columns if isinstance(column, int)])
    revenue_year_columns = sorted([column for column in revenue_wide.columns if isinstance(column, int)])

    scope1_wide, _ = fill_annual_missing_with_previous_value(scope1_wide, scope1_year_columns)
    revenue_wide, _ = fill_annual_missing_with_previous_value(revenue_wide, revenue_year_columns)

    scope1_long = scope1_wide.melt(
        id_vars=["isin"], value_vars=scope1_year_columns, var_name="year", value_name="scope1_co2")
    
    revenue_long = revenue_wide.melt(
        id_vars=["isin"], value_vars=revenue_year_columns, var_name="year", value_name="revenue_thousand_usd")
    

    scope1_long["year"] = scope1_long["year"].astype(int)
    revenue_long["year"] = revenue_long["year"].astype(int)

    annual_data = em_companies.merge(scope1_long, on="isin", how="left")
    annual_data = annual_data.merge(revenue_long, on=["isin", "year"], how="left")

    year_end_prices = monthly_data.loc[monthly_data["date"].dt.month == 12].copy()
    year_end_prices["year"] = year_end_prices["date"].dt.year
    year_end_prices = year_end_prices.rename(
        columns={
        "market_value_musd": "year_end_market_value_musd",
            "return_index": "year_end_return_index"})
    
    year_end_prices["price_available_eoy"] = year_end_prices["year_end_return_index"].notna()

    # IA input
    annual_data = annual_data.merge(
        year_end_prices[
            [
                "isin",
                "year",
        "year_end_market_value_musd",
                "year_end_return_index",
                "price_available_eoy",
            ]
        ],
        on=["isin", "year"],
        how="left")

    annual_data = annual_data.sort_values(["isin", "year"]).reset_index(drop=True)

    return annual_data


def build_base_investment_set(monthly_data: pd.DataFrame):
    """
    Je construis l'univers investissable de base pour la partie 2.1.

    A la fin de l'annee Y, je garde l'entreprise si:
    - le prix de fin d'annee Y est disponible,
    - la part de rendements mensuels egaux a 0 sur les 10 dernieres annees
      ne depasse pas 50%.

    Je me limite aux annees de formation 2013 a 2024
    """
    # Je garde seulement les observations de decembre, car ce sont elles qui representent la fin d'annee.
    year_end_rows = monthly_data.loc[monthly_data["date"].dt.month == 12].copy()

    # J'associe chaque fin d'annee Y a son annee de formation et a l'annee d'investissement Y+1.
    year_end_rows["formation_year"] = year_end_rows["date"].dt.year
    year_end_rows["investment_year"] = year_end_rows["formation_year"] + 1

    # Je verifie si le prix de fin d'annee est disponible, car sans lui on ne peut pas investir l'annee suivante.
    year_end_rows["price_available_eoy"] = year_end_rows["return_index"].notna()

    # Je prepare une liste qui va contenir une table par annee de formation.
    yearly_results = []

    for formation_year in range(FIRST_FORMATION_YEAR, LAST_FORMATION_YEAR + 1):
        # Je definis la fenetre glissante de 10 ans qui sert a mesurer les stale prices.
        window_start = pd.Timestamp(formation_year - ESTIMATION_WINDOW_YEARS + 1, 1, 1)
        window_end = pd.Timestamp(formation_year, 12, 31)

        # Je garde seulement les rendements mensuels de cette fenetre pour calculer les statistiques par entreprise.
        window_data = monthly_data.loc[
            (monthly_data["date"] >= window_start) & (monthly_data["date"] <= window_end),
            ["isin", "monthly_return"]].copy()

        # Je compte combien de rendements mensuels sont disponibles pour chaque entreprise. (IA Input)
        valid_return_count = (
            window_data.groupby("isin")["monthly_return"]
            .apply(lambda values: int(values.notna().sum()))
            .rename("valid_return_count_10y")
            .reset_index())

        # Je compte combien de rendements egaux a 0 apparaissent dans la meme fenetre. (IA Input)
        zero_return_count = (
            window_data.groupby("isin")["monthly_return"]
            .apply(lambda values: int((values.eq(0) & values.notna()).sum()))
            .rename("zero_return_count_10y")
            .reset_index())

        # Je regroupe les deux comptes dans une seule petite table annuelle.
        stale_stats = valid_return_count.merge(zero_return_count, on="isin", how="left")

        # Je calcule la part de rendements nuls sur les rendements observes.
        stale_stats["zero_return_ratio_10y"] = (stale_stats["zero_return_count_10y"] / stale_stats["valid_return_count_10y"])

        # Si je n'ai aucun rendement valide, je laisse le ratio vide pour eviter une division trompeuse.
        stale_stats.loc[stale_stats["valid_return_count_10y"] == 0, "zero_return_ratio_10y"] = pd.NA

        # Je marque comme stale toute entreprise dont plus de 50% des rendements observes sont egaux a 0.
        stale_stats["stale_price_flag"] = stale_stats["zero_return_ratio_10y"] > STALE_PRICE_THRESHOLD
        stale_stats["stale_price_flag"] = stale_stats["stale_price_flag"].fillna(False)

        # Je prends maintenant uniquement les entreprises observees a fin decembre de l'annee de formation.
        year_slice = year_end_rows.loc[year_end_rows["formation_year"] == formation_year].copy()

        # Je rattache a chaque entreprise ses statistiques de stale prices calculees sur la fenetre 10 ans.
        year_slice = year_slice.merge(stale_stats, on="isin", how="left")

        # Si une entreprise n'a pas de statistiques, je complete a 0 ou False pour garder une table propre.
        year_slice["valid_return_count_10y"] = year_slice["valid_return_count_10y"].fillna(0).astype(int)
        year_slice["zero_return_count_10y"] = year_slice["zero_return_count_10y"].fillna(0).astype(int)
        year_slice["stale_price_flag"] = year_slice["stale_price_flag"].fillna(False)

        # Une entreprise est investissable de base si son prix de fin d'annee existe encore et si elle n'est pas stale.
        year_slice["base_investable_next_year"] = (year_slice["price_available_eoy"] & year_slice["stale_price_flag"].eq(False))

        # Je garde seulement les colonnes utiles pour la suite du projet.
        year_slice = year_slice[
            [
                "isin",
                "company_name",
                "country",
                "region",
                "delisting_date",
                "formation_year",
                "investment_year",
                "market_value_musd",
                "return_index",
                "price_available_eoy",
                "valid_return_count_10y",
                "zero_return_count_10y",
                "zero_return_ratio_10y",
                "stale_price_flag",
                "base_investable_next_year"]].copy()

        # Je renomme les prix de decembre pour montrer clairement qu'ils correspondent a la fin d'annee.
        year_slice = year_slice.rename(
            columns={
                "market_value_musd": "year_end_market_value_musd",
                "return_index": "year_end_return_index"})

        # J'ajoute la table de cette annee a la liste finale.
        yearly_results.append(year_slice)

    # Je rassemble toutes les annees dans une seule table puis je les trie proprement.
    base_investment_set = pd.concat(yearly_results, ignore_index=True)
    base_investment_set = base_investment_set.sort_values(["formation_year", "isin"]).reset_index(drop=True)

    return base_investment_set


def save_outputs(
    em_companies: pd.DataFrame,
    monthly_data: pd.DataFrame,
    annual_data: pd.DataFrame,
    base_investment_set: pd.DataFrame):
    
    PROCESSED_DIR.mkdir(parents=True, exist_ok=True)

    written_files = {
        "A": str(write_excel(rename_columns_for_export(em_companies), OUTPUT_FILES["companies"])),
        "B": str(write_excel(rename_columns_for_export(monthly_data), OUTPUT_FILES["monthly_data"])),
        "C": str(write_excel(rename_columns_for_export(annual_data), OUTPUT_FILES["annual_data"])),
        "D": str(write_excel(rename_columns_for_export(base_investment_set), OUTPUT_FILES["base_investment_set"]))}
    
    return written_files

def main():
    # Je commence par definir l'univers EM du projet.
    log_step("Etape 1/5 - Je charge les entreprises Emerging Markets...")
    em_companies = load_em_companies()

    # Je retire ensuite les entreprises qui manquent completement dans une table utile.
    log_step("Etape 2/5 - Je retire les ISIN dont la ligne complete manque dans une table utile...")
    em_companies = keep_only_common_isins(em_companies)

    # Je nettoie les prix mensuels avant toute chose, car ils servent aux regles d'investissement.
    log_step("Etape 3/5 - Je nettoie les prix mensuels et je calcule les rendements mensuels...")
    monthly_data = build_monthly_data(em_companies)

    # Je nettoie ensuite les donnees annuelles utiles a la partie 2.1.
    log_step("Etape 4/5 - Je nettoie Scope 1, Revenue et les prix de fin d'annee...")
    annual_data = build_annual_data(em_companies, monthly_data)

    # Je termine par l'univers investissable de base et les exports.
    log_step("Etape 5/5 - Je construis l'univers investissable de base et j'enregistre les sorties...")
    base_investment_set = build_base_investment_set(monthly_data)
    written_files = save_outputs(
        em_companies=em_companies,
        monthly_data=monthly_data,
        annual_data=annual_data,
        base_investment_set=base_investment_set)

    log_step("Partie 1 terminee.")
    log_step(f"Nombre d'entreprises EM retenues: {len(em_companies)}")
    log_step(f"Lignes du fichier mensuel: {len(monthly_data)}")
    log_step(f"Lignes du fichier annuel: {len(annual_data)}")
    log_step(f"Lignes du base investment set: {len(base_investment_set)}")
    log_step("Fichiers ecrits:")
    for label, path in written_files.items():
        log_step(f"{label} : {path}")

    print("Data cleaning termine.", flush=True)

if __name__ == "__main__":
    main()
