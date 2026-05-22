from __future__ import annotations

from datetime import datetime
from pathlib import Path
import subprocess
import sys


def log_step(message: str):
    """Je montre clairement dans le terminal ou j'en suis."""
    print(message, flush=True)


BASE_DIR = Path(__file__).resolve().parents[1]
SRC_DIR = BASE_DIR / "src"


def run_step(step_label: str, script_name: str):
    """
    Je lance une etape en sous-processus et j'affiche son temps d'execution.

    Si une etape casse, le script s'arrete tout de suite.
    """
    start_time = datetime.now()  # Je note l'heure de debut de l'etape.
    log_step(f"{step_label} - Debut")  # J'annonce le debut de l'etape.

    script_path = SRC_DIR / script_name  # Je pointe vers le script exact a executer.
    subprocess.run(
        [sys.executable, str(script_path)],
        cwd=str(BASE_DIR),
        check=True,
    )  # Je lance le script dans un processus Python separe pour eviter les blocages de memoire ou d'etat.

    end_time = datetime.now()  # Je note l'heure de fin.
    duration = end_time - start_time  # Je calcule la duree totale.
    log_step(f"{step_label} - Termine en {duration}")  # J'affiche une fin claire avec la duree.


def main():
    """
    Je lance la pipeline actuelle du projet dans le bon ordre.

    J'enchaine d'abord la Part I, puis les Parties 3 et 4 bloc par bloc.
    """
    log_step("Lancement de la pipeline du projet...")  # J'annonce le lancement global.

    run_step("Etape 1/11 - Data cleaning", "p1_1_data_cleaning.py")  # Je commence par nettoyer les donnees.
    run_step("Etape 2/11 - Minimum Variance 2.1", "p2_1_investment_set.py")  # Je construis ensuite l'investment set.
    run_step("Etape 3/11 - Minimum Variance 2.2", "p2_2_minimum_variance_portfolio.py")  # Je calcule ensuite le portefeuille minimum variance.
    run_step("Etape 4/11 - Value-Weighted 2.3", "p2_3_value_weighted_portfolio.py")  # Je termine la Part I avec le benchmark value-weighted.
    run_step("Etape 5/11 - MinVar vs Value-Weighted Comparison 2.4", "p2_4_minvar_vs_valueweighted_comparison.py") # Je trace le graphique comparatif minimum variance vs value-weighted.
    run_step("Etape 6/11 - Carbon Footprint 3.1", "p3_1_carbon_footprint.py")  # Je mesure ensuite le profil carbone des deux portefeuilles de reference.
    run_step("Etape 7/11 - Minimum Variance Carbon 3.2", "p3_2_minimum_variance_carbon.py")  # Je construis ensuite la version active avec contrainte carbone a 50%.
    run_step("Etape 8/11 - Tracking Error Carbon 3.3", "p3_3_tracking_error_carbon.py")  # Je construis ensuite la version passive avec contrainte carbone a 50%.
    run_step("Etape 9/11 - Carbon Comparison 3.4", "p3_4_carbon_comparison.py")  # Je termine la comparaison de la Partie 3.
    run_step("Etape 10/11 - Net Zero 4.1", "p4_1_net_zero.py")  # Je construis ensuite la strategie net-zero passive.
    run_step("Etape 11/11 - Passive Comparison 4.2", "p4_2_passive_comparison.py")  # Je termine par la comparaison finale des strategies passives.

    log_step("Pipeline terminee.")  # Je confirme que la pipeline actuelle est finie.


if __name__ == "__main__":
    main()
