from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd
import time


BASE_DIR = Path(__file__).resolve().parents[1]
PROCESSED_DIR = BASE_DIR / "data" / "processed"

MV_PERFORMANCE_FILE = "K_MinVar_2_2_Monthly_Performance.xlsx"
VW_PERFORMANCE_FILE = "N_ValueWeighted_2_3_Monthly_Performance.xlsx"
CUMULATIVE_FIGURE = "P_MinVar_vs_ValueWeighted_Cumulative_Returns.png"


def save_cumulative_figure(mv_performance: pd.DataFrame, vw_performance: pd.DataFrame):
    """Je sauvegarde la comparaison des rendements cumules entre P(mv)_oos et P(vw)."""
    figure_path = PROCESSED_DIR / CUMULATIVE_FIGURE

    fig, ax = plt.subplots(figsize=(12, 6), dpi=150)

    ax.plot(
        mv_performance["date"],
        mv_performance["cumulative_growth"],
        label="P(mv)_oos",
    )
    ax.plot(
        vw_performance["date"],
        vw_performance["cumulative_growth"],
        label="P(vw)",
    )

    ax.set_title("Cumulative Returns: P(mv)_oos vs P(vw)")
    ax.set_xlabel("Date")
    ax.set_ylabel("Cumulative Growth")
    ax.legend()
    ax.grid(True)

    fig.tight_layout()
    fig.savefig(figure_path, bbox_inches="tight")
    plt.close(fig)

    print(f"Figure cumulative ecrite: {figure_path}", flush=True)


def main():
    start_time = time.perf_counter()
    print("Section 2.4 - Je compare P(mv)_oos et P(vw)", flush=True)

    mv_performance = pd.read_excel(PROCESSED_DIR / MV_PERFORMANCE_FILE, parse_dates=["date"])
    vw_performance = pd.read_excel(PROCESSED_DIR / VW_PERFORMANCE_FILE, parse_dates=["date"])

    mv_performance["cumulative_growth"] = (1.0 + mv_performance["portfolio_return"]).cumprod()
    vw_performance["cumulative_growth"] = (1.0 + vw_performance["portfolio_return"]).cumprod()

    save_cumulative_figure(mv_performance, vw_performance)

    elapsed = time.perf_counter() - start_time
    print("Section 2.4 complete.", flush=True)


if __name__ == "__main__":
    main()