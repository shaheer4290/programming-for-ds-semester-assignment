"""
COM7024 Programming for Data Analytics
Student number: 25106724

Exploratory analysis of Manchester_house_Dataset[3678].csv for Estate
Management Plc. Every figure and number in the submitted report is produced
by this script.

Programming paradigms used on purpose:
- procedural: this module-level workflow
- object-oriented: HousingStudy holds the data and the analysis steps
- functional: column derivations use vectorised expressions rather than loops
"""

from __future__ import annotations

import json
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns
from scipy import stats
from sklearn.linear_model import LinearRegression
from sklearn.metrics import mean_absolute_error, r2_score
from sklearn.model_selection import train_test_split

STUDENT_ID = "25106724"
ROOT = Path(__file__).resolve().parent
DATA = ROOT / "Manchester_house_Dataset[3678].csv"
FIG = ROOT / "figures"
FIG.mkdir(exist_ok=True)

NAVY = "#1F4E79"
TEAL = "#1A7A6D"
AMBER = "#C47B17"
SLATE = "#5C6B73"

NUMERIC = [
    "price",
    "bedrooms",
    "bathrooms",
    "sqft_living",
    "sqft_total",
    "floors",
    "waterfront",
    "view",
    "condition",
    "grade",
    "built",
    "renovated",
    "living_area",
]


def configure_plots() -> None:
    sns.set_theme(style="whitegrid", context="notebook")
    plt.rcParams.update(
        {
            "font.size": 11,
            "axes.titlesize": 12,
            "axes.labelsize": 11,
            "figure.dpi": 140,
            "savefig.dpi": 170,
            "axes.spines.top": False,
            "axes.spines.right": False,
            "axes.titleweight": "semibold",
        }
    )


def _summary(frame: pd.DataFrame) -> dict:
    """Basic statistical profile used before and after preprocessing."""
    desc = frame[NUMERIC].describe().round(2)
    skew = frame[NUMERIC].skew(numeric_only=True).round(3)
    missing = frame[NUMERIC].isna().sum().astype(int)
    profile = {}
    for column in NUMERIC:
        profile[column] = {
            "count": int(desc.loc["count", column]),
            "mean": float(desc.loc["mean", column]),
            "std": float(desc.loc["std", column]),
            "min": float(desc.loc["min", column]),
            "q1": float(desc.loc["25%", column]),
            "median": float(desc.loc["50%", column]),
            "q3": float(desc.loc["75%", column]),
            "max": float(desc.loc["max", column]),
            "skew": float(skew[column]),
            "missing": int(missing[column]),
        }
    return profile


def _corr_pair(x: pd.Series, y: pd.Series) -> dict:
    pearson_r, pearson_p = stats.pearsonr(x, y)
    spearman_r, spearman_p = stats.spearmanr(x, y)
    return {
        "n": int(len(x)),
        "pearson_r": round(float(pearson_r), 4),
        "pearson_p": float(pearson_p),
        "spearman_r": round(float(spearman_r), 4),
        "spearman_p": float(spearman_p),
    }


class HousingStudy:
    """One object owns the raw frame, the cleaned frame, and the tests."""

    def __init__(self, path: Path) -> None:
        self.path = path
        self.raw: pd.DataFrame | None = None
        self.typed: pd.DataFrame | None = None
        self.clean: pd.DataFrame | None = None
        self.audit: dict = {}

    def load(self) -> pd.DataFrame:
        try:
            frame = pd.read_csv(self.path)
        except FileNotFoundError as exc:
            raise FileNotFoundError(
                f"Dataset not found at {self.path}. Place the module CSV beside this script."
            ) from exc
        # The export starts with an unnamed row-index column.
        unnamed = [col for col in frame.columns if str(col).startswith("Unnamed") or str(col).strip() == ""]
        frame = frame.drop(columns=unnamed)
        self.raw = frame
        self.audit["rows_imported"] = int(len(frame))
        self.audit["columns_imported"] = list(frame.columns)
        self.audit["dtypes_on_import"] = {col: str(dtype) for col, dtype in frame.dtypes.items()}
        return frame

    def correct_types(self) -> pd.DataFrame:
        frame = self.raw.copy()
        frame["id"] = frame["id"].astype("string")
        for column in NUMERIC:
            frame[column] = pd.to_numeric(frame[column], errors="coerce")
        frame["waterfront"] = frame["waterfront"].astype("Int64")
        frame["renovated_flag"] = (frame["renovated"].fillna(0) > 0).astype(int)
        self.typed = frame
        self.audit["dtypes_corrected"] = {col: str(dtype) for col, dtype in frame.dtypes.items()}
        return frame

    def preprocess(self) -> pd.DataFrame:
        """Remove records that cannot be properties, and keep the market tail.

        Luxury and waterfront prices sit in the right tail. Winsorising price
        would hide the very contrast the estate manager asked about, so only
        structurally impossible rows are removed.
        """
        frame = self.typed.copy()
        before = len(frame)
        exact_dupes = int(frame.duplicated().sum())
        id_dupes = int(frame["id"].duplicated().sum())
        frame = frame.drop_duplicates()
        frame = frame.drop_duplicates(subset="id", keep="first")

        invalid = (
            frame["price"].isna()
            | (frame["price"] <= 0)
            | frame["sqft_living"].isna()
            | (frame["sqft_living"] <= 0)
            | frame["sqft_total"].isna()
            | (frame["sqft_total"] <= 0)
            | frame["bedrooms"].isna()
            | (frame["bedrooms"] <= 0)
            | (frame["bedrooms"] > 15)
            | frame["bathrooms"].isna()
            | (frame["bathrooms"] <= 0)
            | frame["built"].isna()
            | (frame["built"] < 1800)
            | (frame["built"] > 2025)
            | ~frame["waterfront"].isin([0, 1])
        )
        removed = frame.loc[invalid].copy()
        frame = frame.loc[~invalid].copy()
        frame["price_per_sqft"] = frame["price"] / frame["sqft_living"]
        frame["decade"] = (frame["built"] // 10 * 10).astype(int)
        frame["age_proxy"] = int(frame["built"].max()) - frame["built"]

        self.clean = frame
        self.audit.update(
            {
                "rows_before_cleaning": before,
                "exact_duplicate_rows": exact_dupes,
                "duplicate_ids": id_dupes,
                "invalid_rows_removed": int(len(removed)),
                "invalid_examples": removed.head(8)[
                    ["id", "price", "bedrooms", "bathrooms", "sqft_living", "built", "waterfront"]
                ].astype(object).to_dict(orient="records"),
                "rows_after_cleaning": int(len(frame)),
                "rows_removed": int(before - len(frame)),
            }
        )
        return frame

    def initial_profile(self) -> dict:
        return _summary(self.typed)

    def cleaned_profile(self) -> dict:
        return _summary(self.clean)


def waterfront_tests(frame: pd.DataFrame) -> dict:
    water = frame.loc[frame["waterfront"] == 1, "price"]
    inland = frame.loc[frame["waterfront"] == 0, "price"]
    water_psf = frame.loc[frame["waterfront"] == 1, "price_per_sqft"]
    inland_psf = frame.loc[frame["waterfront"] == 0, "price_per_sqft"]
    u_stat, u_p = stats.mannwhitneyu(water, inland, alternative="two-sided")
    u_psf, p_psf = stats.mannwhitneyu(water_psf, inland_psf, alternative="two-sided")
    n1, n2 = len(water), len(inland)
    # Positive when the first sample (waterfront) tends to take the higher ranks.
    rank_biserial = (2 * u_stat) / (n1 * n2) - 1
    return {
        "n_waterfront": int(n1),
        "n_inland": int(n2),
        "share_waterfront": round(n1 / len(frame), 4),
        "median_price_waterfront": float(water.median()),
        "median_price_inland": float(inland.median()),
        "mean_price_waterfront": float(water.mean()),
        "mean_price_inland": float(inland.mean()),
        "median_ratio": round(float(water.median() / inland.median()), 3),
        "median_psf_waterfront": float(water_psf.median()),
        "median_psf_inland": float(inland_psf.median()),
        "median_living_waterfront": float(frame.loc[frame["waterfront"] == 1, "sqft_living"].median()),
        "median_living_inland": float(frame.loc[frame["waterfront"] == 0, "sqft_living"].median()),
        "mannwhitney_price_U": float(u_stat),
        "mannwhitney_price_p": float(u_p),
        "rank_biserial": round(float(rank_biserial), 4),
        "mannwhitney_psf_p": float(p_psf),
        "mannwhitney_psf_U": float(u_psf),
    }


def floor_space_tests(frame: pd.DataFrame) -> dict:
    targets = ["sqft_living", "sqft_total", "floors", "living_area", "bedrooms", "bathrooms", "grade"]
    correlations = {column: _corr_pair(frame[column], frame["price"]) for column in targets}
    by_group = {}
    for label, flag in (("waterfront", 1), ("inland", 0)):
        subset = frame.loc[frame["waterfront"] == flag]
        by_group[label] = _corr_pair(subset["sqft_living"], subset["price"])
    return {"with_price": correlations, "living_by_waterfront": by_group}


def built_price_tests(frame: pd.DataFrame) -> dict:
    overall = _corr_pair(frame["built"], frame["price"])
    by_group = {}
    for label, flag in (("waterfront", 1), ("inland", 0)):
        subset = frame.loc[frame["waterfront"] == flag]
        by_group[label] = _corr_pair(subset["built"], subset["price"])
    decade = (
        frame.groupby("decade")
        .agg(
            n=("price", "size"),
            median_price=("price", "median"),
            median_living=("sqft_living", "median"),
            waterfront_share=("waterfront", "mean"),
        )
        .reset_index()
    )
    decade["median_price"] = decade["median_price"].round(0)
    decade["median_living"] = decade["median_living"].round(0)
    decade["waterfront_share"] = decade["waterfront_share"].round(4)
    return {
        "overall": overall,
        "by_waterfront": by_group,
        "by_decade": decade.to_dict(orient="records"),
    }


def taught_patterns(frame: pd.DataFrame) -> dict:
    """Conditionals and iteration, applied to the decade medians rather than left as a demo."""
    bands = {"under_400k": 0, "400k_to_550k": 0, "550k_or_more": 0}
    for _decade, group in frame.groupby("decade"):
        median_price = float(group["price"].median())
        if median_price < 400_000:
            bands["under_400k"] += 1
        elif median_price < 550_000:
            bands["400k_to_550k"] += 1
        else:
            bands["550k_or_more"] += 1
    type_notes = []
    for column, expected in (("id", "string"), ("price", "int"), ("waterfront", "Int")):
        actual = str(frame[column].dtype)
        if expected not in actual:
            type_notes.append(f"{column} dtype is {actual}, expected to contain '{expected}'")
    return {"decade_median_bands": bands, "type_notes": type_notes}


def comparison_tests(frame: pd.DataFrame) -> dict:
    """Welch t-test is recorded so it can be compared with the rank test. It is not the lead result."""
    water = np.log(frame.loc[frame["waterfront"] == 1, "price"].astype(float))
    inland = np.log(frame.loc[frame["waterfront"] == 0, "price"].astype(float))
    t_stat, t_p = stats.ttest_ind(water, inland, equal_var=False)
    decade_groups = [group["price"].to_numpy() for _, group in frame.groupby("decade") if len(group) > 1]
    h_stat, h_p = stats.kruskal(*decade_groups)
    return {
        "welch_log_price_t": round(float(t_stat), 2),
        "welch_log_price_p": float(t_p),
        "kruskal_decade_H": round(float(h_stat), 2),
        "kruskal_decade_p": float(h_p),
        "n_decade_groups": len(decade_groups),
    }


def distribution_tests(before: pd.DataFrame, after: pd.DataFrame) -> dict:
    """Normality of price, before and after the rows that failed validation."""
    out = {}
    for label, series in (("before", before["price"]), ("after", after["price"])):
        sample = series.sample(n=min(5000, len(series)), random_state=24)
        stat, p_value = stats.normaltest(sample)
        out[label] = {
            "n_tested": int(len(sample)),
            "dagostino_k2": round(float(stat), 2),
            "p_value": float(p_value),
            "skew": round(float(series.skew()), 3),
        }
    renovated = after["renovated_flag"] == 1
    out["renovated_share"] = round(float(renovated.mean()), 4)
    out["median_price_renovated"] = float(after.loc[renovated, "price"].median())
    out["median_price_never_renovated"] = float(after.loc[~renovated, "price"].median())
    return out


def baseline_model(frame: pd.DataFrame) -> dict:
    """Transparent linear baseline. This is a decision aid, not a valuation model."""
    features = ["sqft_living", "sqft_total", "floors", "waterfront", "built", "grade", "condition", "view"]
    work = frame[features + ["price"]].dropna().copy()
    work["waterfront"] = work["waterfront"].astype(int)
    x_train, x_test, y_train, y_test = train_test_split(
        work[features], np.log(work["price"]), test_size=0.25, random_state=25106724
    )
    model = LinearRegression()
    model.fit(x_train, y_train)
    pred = model.predict(x_test)
    mae = mean_absolute_error(np.exp(y_test), np.exp(pred))
    coefficients = {name: round(float(coef), 6) for name, coef in zip(features, model.coef_)}
    percent = {name: round((float(np.exp(coef)) - 1) * 100, 2) for name, coef in coefficients.items()}
    return {
        "features": features,
        "n_train": int(len(x_train)),
        "n_test": int(len(x_test)),
        "r2_log_price": round(float(r2_score(y_test, pred)), 4),
        "mae_currency": round(float(mae), 0),
        "intercept": round(float(model.intercept_), 4),
        "coefficients_log": coefficients,
        "approx_percent_effect": percent,
    }


def _save(fig: plt.Figure, name: str) -> None:
    fig.tight_layout()
    fig.savefig(FIG / name, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def make_figures(frame: pd.DataFrame, before: pd.DataFrame) -> None:
    fig, axes = plt.subplots(1, 2, figsize=(10.2, 4.2))
    axes[0].hist(before["price"] / 1000, bins=40, color=SLATE, edgecolor="white")
    axes[0].set_title("Price before cleaning")
    axes[0].set_xlabel("Price (thousands, as recorded)")
    axes[1].hist(frame["price"] / 1000, bins=40, color=NAVY, edgecolor="white")
    axes[1].set_title("Price after removing invalid rows")
    axes[1].set_xlabel("Price (thousands, as recorded)")
    for axis in axes:
        axis.set_ylabel("Properties")
    _save(fig, "fig1_price_before_after.png")

    fig, ax = plt.subplots(figsize=(7.2, 4.6))
    plot_df = frame.assign(setting=np.where(frame["waterfront"] == 1, "Waterfront", "Inland"))
    sns.boxplot(
        data=plot_df,
        x="setting",
        y="price",
        hue="setting",
        palette={"Inland": NAVY, "Waterfront": TEAL},
        legend=False,
        showfliers=False,
        ax=ax,
    )
    ax.set_ylabel("Sale price")
    ax.set_xlabel("")
    ax.set_title("Sale price by waterfront status (extreme points hidden)")
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"{v/1000:.0f}k"))
    _save(fig, "fig2_waterfront_price.png")

    fig, ax = plt.subplots(figsize=(7.4, 4.8))
    inland = frame["waterfront"] == 0
    ax.scatter(
        frame.loc[inland, "sqft_living"],
        frame.loc[inland, "price"],
        s=8,
        alpha=0.25,
        c=NAVY,
        label="Inland",
        linewidths=0,
    )
    ax.scatter(
        frame.loc[~inland, "sqft_living"],
        frame.loc[~inland, "price"],
        s=22,
        alpha=0.8,
        c=TEAL,
        label="Waterfront",
        linewidths=0,
    )
    ax.set_xlabel("Internal floor space (sqft_living)")
    ax.set_ylabel("Sale price")
    ax.set_title("Floor space and price")
    ax.legend(frameon=False)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"{v/1e6:.1f}m"))
    _save(fig, "fig3_floorspace_price.png")

    cols = ["price", "sqft_living", "sqft_total", "floors", "living_area", "built", "waterfront", "grade"]
    corr = frame[cols].corr(method="spearman")
    fig, ax = plt.subplots(figsize=(7.6, 6.2))
    sns.heatmap(corr, annot=True, fmt=".2f", cmap="Blues", vmin=-0.2, vmax=1, ax=ax, square=True)
    ax.set_title("Spearman correlations")
    _save(fig, "fig4_correlation_heatmap.png")

    fig, ax = plt.subplots(figsize=(7.6, 4.8))
    sample = frame.sample(n=min(4000, len(frame)), random_state=24)
    ax.scatter(sample["built"], sample["price"], s=8, alpha=0.2, c=SLATE, linewidths=0)
    decade = frame.groupby("decade")["price"].median()
    ax.plot(decade.index + 5, decade.values, color=AMBER, lw=2.4, marker="o", label="Median price by decade")
    ax.set_xlabel("Year built")
    ax.set_ylabel("Sale price")
    ax.set_title("Year built against sale price")
    ax.legend(frameon=True, loc="upper left", framealpha=0.92)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"{v/1000:.0f}k"))
    _save(fig, "fig5_built_price.png")

    fig, ax = plt.subplots(figsize=(7.4, 4.6))
    order = sorted(frame["floors"].unique())
    sns.pointplot(
        data=frame,
        x="floors",
        y="price_per_sqft",
        order=order,
        color=NAVY,
        errorbar=("ci", 95),
        ax=ax,
    )
    ax.set_xlabel("Number of floors")
    ax.set_ylabel("Price per square foot of living space")
    ax.set_title("Floor count and price intensity")
    _save(fig, "fig6_floors_intensity.png")

    decade = (
        frame.groupby("decade")["price"].median().rename("median_price").reset_index()
    )
    fig, ax = plt.subplots(figsize=(7.6, 4.4))
    ax.bar(decade["decade"].astype(str), decade["median_price"], color=NAVY, width=0.8)
    ax.set_xlabel("Decade built")
    ax.set_ylabel("Median sale price")
    ax.set_title("Median price by decade of construction")
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"{v/1000:.0f}k"))
    ax.tick_params(axis="x", rotation=45)
    _save(fig, "fig7_decade_bars.png")


def write_execution_log(lines: list[str]) -> None:
    (ROOT / "execution_log.txt").write_text("\n".join(lines) + "\n", encoding="utf-8")
    shown = lines[:42]
    fig, ax = plt.subplots(figsize=(8.4, 6.6))
    ax.axis("off")
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.add_patch(plt.Rectangle((0, 0), 1, 1, color="#1B2430"))
    ax.text(0.03, 0.97, "\n".join(shown), va="top", ha="left", family="monospace", fontsize=7.4, color="#D6F5E3")
    fig.savefig(FIG / "fig0_execution_evidence.png", dpi=160, facecolor="#1B2430")
    plt.close(fig)


def main() -> dict:
    configure_plots()
    log: list[str] = [f"COM7024 analysis | student {STUDENT_ID}"]
    study = HousingStudy(DATA)
    study.load()
    study.correct_types()
    before = study.initial_profile()
    study.preprocess()
    after = study.cleaned_profile()
    frame = study.clean
    water = waterfront_tests(frame)
    floors = floor_space_tests(frame)
    built = built_price_tests(frame)
    model = baseline_model(frame)
    distributions = distribution_tests(study.typed, frame)
    patterns = taught_patterns(frame)
    comparisons = comparison_tests(frame)
    make_figures(frame, study.typed)

    log.append(f"Imported rows: {study.audit['rows_imported']}")
    log.append(f"Columns: {', '.join(study.audit['columns_imported'])}")
    log.append("Dtypes on import: " + ", ".join(f"{k}:{v}" for k, v in study.audit["dtypes_on_import"].items()))
    log.append(
        f"Cleaning: removed {study.audit['rows_removed']} rows "
        f"({study.audit['exact_duplicate_rows']} exact duplicates, "
        f"{study.audit['invalid_rows_removed']} invalid)."
    )
    log.append(f"Rows analysed: {study.audit['rows_after_cleaning']}")
    log.append(
        f"Waterfront n={water['n_waterfront']} median={water['median_price_waterfront']:.0f} | "
        f"inland n={water['n_inland']} median={water['median_price_inland']:.0f} | "
        f"ratio={water['median_ratio']} | Mann-Whitney p={water['mannwhitney_price_p']:.3e}"
    )
    log.append(
        f"sqft_living vs price Spearman r={floors['with_price']['sqft_living']['spearman_r']} "
        f"p={floors['with_price']['sqft_living']['spearman_p']:.3e}"
    )
    log.append(
        f"built vs price Pearson r={built['overall']['pearson_r']} "
        f"Spearman r={built['overall']['spearman_r']} "
        f"p={built['overall']['spearman_p']:.3e}"
    )
    log.append(
        f"Baseline hold-out R2 (log price)={model['r2_log_price']} "
        f"MAE={model['mae_currency']:.0f}"
    )
    for line in log:
        print(line)

    payload = {
        "student_id": STUDENT_ID,
        "audit": study.audit,
        "profile_before": before,
        "profile_after": after,
        "waterfront": water,
        "floor_space": floors,
        "built_price": built,
        "model": model,
        "distributions": distributions,
        "patterns": patterns,
        "comparisons": comparisons,
        "log": log,
    }
    (ROOT / "results.json").write_text(json.dumps(payload, indent=2), encoding="utf-8")
    write_execution_log(log)
    print("Saved results.json, execution_log.txt and figures/")
    return payload


if __name__ == "__main__":
    main()
