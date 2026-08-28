"""
Verrentungs_code_streamlit.py

Flexible engine for:
    reading MSCI World + REXP + CPI from an Excel file
    building nominal and real return series with optional fees
    constructing a portfolio return series from configurable weights
    running a constant-withdrawal simulation with configurable rate and horizon
    modelling capital gains tax on realised gains via a cost-basis approach
    producing charts (fan chart, success curve, distribution)
    exporting all historical cohort paths and a cohort summary table

All important scenario parameters are configured once in the SCENARIO SETTINGS
section below.

This file contains both:
    1) The full engine (functions + plotting helpers)
    2) A Streamlit front-end (no separate engine file required)

VARIANTE B (IMPORTIERBAR FÜR TESTS):
    Der komplette Streamlit-Code ist in run_app() gekapselt.
    Beim Import (z.B. durch pytest) wird die UI NICHT ausgeführt.

Run with:
    streamlit run Verrentungs_code_streamlit.py
"""

from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Optional

from matplotlib.ticker import FuncFormatter
from matplotlib.figure import Figure
from matplotlib.backends.backend_pdf import PdfPages

import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

import streamlit as st
import io

st.markdown("""
<style>
@media print {
    /* Hide Streamlit chrome */
    [data-testid="stSidebar"],
    header,
    footer,
    .stDeployButton,
    #MainMenu {
        display: none !important;
    }

    html, body {
        margin: 0 !important;
        padding: 0 !important;
        width: 210mm !important;
        height: 297mm !important;
        overflow: hidden !important;
        background: white !important;
        color: black !important;
        font-size: 9pt !important;
    }

    [data-testid="stAppViewContainer"],
    .main,
    .block-container {
        margin: 0 !important;
        padding: 0.2cm 0.4cm !important;
        max-width: 100% !important;
        width: 100% !important;
        overflow: hidden !important;
        background: white !important;
    }

    h1, h2, h3, h4, h5, h6 {
        margin-top: 0.2rem !important;
        margin-bottom: 0.2rem !important;
        line-height: 1.1 !important;
    }

    p, div, label, span {
        margin-top: 0 !important;
        margin-bottom: 0 !important;
        line-height: 1.1 !important;
    }

    [data-testid="stHorizontalBlock"] {
        display: flex !important;
        flex-direction: row !important;
        flex-wrap: nowrap !important;
        gap: 0.4rem !important;
    }

    [data-testid="column"] {
        flex: 1 1 0 !important;
        min-width: 0 !important;
        width: auto !important;
    }

    [data-testid="stMetric"] {
        padding: 0 !important;
        margin: 0 !important;
    }

    [data-testid="stMetricLabel"] {
        font-size: 7pt !important;
        line-height: 1.0 !important;
        margin: 0 !important;
    }

    [data-testid="stMetricValue"] {
        font-size: 10pt !important;
        line-height: 1.0 !important;
        white-space: nowrap !important;
        margin: 0 !important;
    }

    [data-testid="stVerticalBlock"] > div {
        padding-top: 0 !important;
        padding-bottom: 0 !important;
        margin-top: 0 !important;
        margin-bottom: 0 !important;
    }

    button, input, textarea, select {
        display: none !important;
    }

    [role="tablist"] {
        display: none !important;
    }

    canvas, img, svg {
        max-width: 100% !important;
        max-height: 150mm !important;
        height: auto !important;
        page-break-inside: avoid !important;
        break-inside: avoid !important;
    }

    [data-testid="stImage"],
    [data-testid="stPlotlyChart"],
    [data-testid="stVegaLiteChart"],
    iframe,
    figure {
        margin: 0 !important;
        padding: 0 !important;
        page-break-inside: avoid !important;
        break-inside: avoid !important;
        max-height: 150mm !important;
    }

    body {
        zoom: 1 !important;
    }

    @page {
        size: A4 portrait;
        margin: 0.5cm;
    }

    * {
        page-break-before: avoid !important;
        page-break-after: avoid !important;
        page-break-inside: avoid !important;
        break-before: avoid !important;
        break-after: avoid !important;
        break-inside: avoid !important;
        box-sizing: border-box !important;
    }
}
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# SCENARIO SETTINGS (ADJUST HERE)
# ---------------------------------------------------------------------------

DATA_FILE = Path("Daten Verrentung_EUR.xlsx")

PORTFOLIO_WEIGHTS = {
    "msci_world": 0.60,
    "rexp": 0.35,
    "Gold": 0.05,
}

ANNUAL_FEE = 0.0184
APPLY_FEES = True

CAPITAL_GAINS_TAX_RATE = 0.25
APPLY_TAX = True

USE_INFLATION = True

WITHDRAWAL_RATE = 0.05
INITIAL_WEALTH = 1_000_000.0
HORIZON_YEARS = 30
PERIODS_PER_YEAR = 12

SUCCESS_RATE_MIN = 0.02
SUCCESS_RATE_MAX = 0.08
SUCCESS_RATE_STEP = 0.0025


# ---------------------------------------------------------------------------
# CONFIGURATION DATACLASSES
# ---------------------------------------------------------------------------

@dataclass
class DataConfig:
    excel_path: Path
    sheet_name: str = "Import_Daten"
    date_column: str = "Dates"
    cpi_column: str = "Inflation DE"
    asset_columns: Dict[str, str] = field(
        default_factory=lambda: {
            "msci_world": "NDDUWI Index",
            "rexp": "REXP Index",
            "Gold": "Gold",
        }
    )


@dataclass
class PortfolioConfig:
    name: str
    weights: Dict[str, float]
    annual_fee: float = 0.0
    use_fees: bool = True
    use_inflation: bool = True


@dataclass
class SimulationConfig:
    annual_withdrawal_rate: float = 0.04
    initial_wealth: float = 100.0
    periods_per_year: int = 12
    horizon_years: Optional[int] = 30


# ---------------------------------------------------------------------------
# DATA LOADING AND NOMINAL RETURNS
# ---------------------------------------------------------------------------

def load_market_data(cfg: DataConfig) -> pd.DataFrame:
    df = pd.read_excel(cfg.excel_path, sheet_name=cfg.sheet_name)
    df[cfg.date_column] = pd.to_datetime(df[cfg.date_column])
    df = df.set_index(cfg.date_column).sort_index()

    cols = {}
    for asset_name, col_name in cfg.asset_columns.items():
        if col_name not in df.columns:
            raise KeyError(f"Column {col_name!r} for asset {asset_name!r} not found")
        cols[asset_name] = df[col_name]

    if cfg.cpi_column not in df.columns:
        raise KeyError(f"CPI column {cfg.cpi_column!r} not found")
    cols["cpi"] = df[cfg.cpi_column]

    panel = pd.DataFrame(cols).dropna()
    return panel


def compute_nominal_returns(panel: pd.DataFrame) -> pd.DataFrame:
    asset_names = [c for c in panel.columns if c != "cpi"]
    asset_levels = panel[asset_names]
    rets = asset_levels.pct_change().dropna()

    cpi = panel["cpi"].pct_change().dropna()
    cpi = cpi.reindex(rets.index)

    result = rets.copy()
    result["inflation"] = cpi
    return result


# ---------------------------------------------------------------------------
# FEES AND REAL RETURNS
# ---------------------------------------------------------------------------

def apply_annual_fee_to_returns(
    rets: pd.DataFrame,
    annual_fee: float,
    asset_cols=None,
) -> pd.DataFrame:
    if asset_cols is None:
        asset_cols = [c for c in rets.columns if c != "inflation"]

    monthly_fee = 1.0 - (1.0 - annual_fee) ** (1.0 / 12.0)
    out = rets.copy()

    for col in asset_cols:
        if col == "inflation":
            continue
        gross = 1.0 + out[col]
        out[col + "_net"] = gross * (1.0 - monthly_fee) - 1.0

    return out


def make_real_returns(rets: pd.DataFrame, use_net: bool = False) -> pd.DataFrame:
    infl = rets["inflation"]
    base_cols = []

    for col in rets.columns:
        if col == "inflation":
            continue
        is_net = col.endswith("_net")
        if use_net and is_net:
            base_cols.append(col)
        elif (not use_net) and (not is_net):
            base_cols.append(col)

    real = {}
    for col in base_cols:
        r = rets[col]
        base_name = col.replace("_net", "")
        suffix = "_real_net" if col.endswith("_net") else "_real"
        real_name = base_name + suffix
        real[real_name] = (1.0 + r) / (1.0 + infl) - 1.0

    return pd.DataFrame(real, index=rets.index)


# ---------------------------------------------------------------------------
# PORTFOLIO CONSTRUCTION
# ---------------------------------------------------------------------------

def build_portfolio_returns(
    rets: pd.DataFrame,
    weights: Dict[str, float],
    use_net: bool,
    use_real: bool,
    portfolio_name: str,
) -> pd.Series:
    columns_to_use: Dict[str, float] = {}

    for asset, w in weights.items():
        if use_real and use_net:
            col = f"{asset}_real_net"
        elif use_real and not use_net:
            col = f"{asset}_real"
        elif (not use_real) and use_net:
            col = f"{asset}_net"
        else:
            col = asset

        if col not in rets.columns:
            raise KeyError(f"Expected column {col!r} for asset {asset!r} not found")
        columns_to_use[col] = w

    port = pd.Series(0.0, index=rets.index, name=portfolio_name)
    for col, w in columns_to_use.items():
        port = port + w * rets[col]

    return port


def prepare_portfolio_returns(
    panel: pd.DataFrame,
    portfolio_cfg: PortfolioConfig,
) -> pd.Series:
    rets = compute_nominal_returns(panel)

    if "cash" not in rets.columns:
        rets["cash"] = 0.0

    if portfolio_cfg.use_fees and portfolio_cfg.annual_fee > 0:
        rets = apply_annual_fee_to_returns(rets, portfolio_cfg.annual_fee)

    real_gross = make_real_returns(rets, use_net=False)
    rets_full = rets.join(real_gross)

    if portfolio_cfg.use_fees and portfolio_cfg.annual_fee > 0:
        real_net = make_real_returns(rets, use_net=True)
        rets_full = rets_full.join(real_net)

    port_rets = build_portfolio_returns(
        rets_full,
        weights=portfolio_cfg.weights,
        use_net=portfolio_cfg.use_fees,
        use_real=portfolio_cfg.use_inflation,
        portfolio_name=portfolio_cfg.name,
    )

    return port_rets


# ---------------------------------------------------------------------------
# WITHDRAWAL SIMULATION WITH COST-BASIS TAX
# ---------------------------------------------------------------------------

def simulate_constant_withdrawal(
    portfolio_returns: pd.Series,
    sim_cfg: SimulationConfig,
    tax_rate: float = 0.0,
) -> pd.DataFrame:
    if sim_cfg.horizon_years is not None:
        max_periods = sim_cfg.horizon_years * sim_cfg.periods_per_year
        returns = portfolio_returns.iloc[:max_periods]
    else:
        returns = portfolio_returns

    withdraw_per_period = (
        sim_cfg.initial_wealth
        * sim_cfg.annual_withdrawal_rate
        / sim_cfg.periods_per_year
    )

    wealth_values = []
    basis_values = []
    withdrawals_net = []
    taxes_paid = []
    gross_sales = []

    wealth = sim_cfg.initial_wealth
    basis = sim_cfg.initial_wealth

    for r in returns:
        if wealth <= 0:
            wealth_values.append(0.0)
            basis_values.append(0.0)
            withdrawals_net.append(0.0)
            taxes_paid.append(0.0)
            gross_sales.append(0.0)
            wealth = 0.0
            basis = 0.0
            continue

        wealth_pre = wealth * (1.0 + r)

        if tax_rate <= 0.0:
            sale = min(withdraw_per_period, wealth_pre)
            tax = 0.0
            net_withdraw = sale
            principal_fraction = (basis / wealth_pre) if wealth_pre > 0 else 0.0
            principal_sold = sale * principal_fraction
            wealth = wealth_pre - sale
            basis = max(basis - principal_sold, 0.0)
        else:
            if wealth_pre > basis:
                total_gain = wealth_pre - basis
                gain_fraction = total_gain / wealth_pre
            else:
                total_gain = 0.0
                gain_fraction = 0.0

            principal_fraction = 1.0 - gain_fraction

            if gain_fraction > 0:
                net_factor = 1.0 - tax_rate * gain_fraction
            else:
                net_factor = 1.0

            if net_factor <= 0.0:
                sale = wealth_pre
            else:
                sale = withdraw_per_period / net_factor

            sale = min(sale, wealth_pre)
            realised_gain = sale * gain_fraction
            tax = tax_rate * realised_gain
            net_withdraw = sale - tax
            principal_sold = sale * principal_fraction
            wealth = wealth_pre - sale
            basis = max(basis - principal_sold, 0.0)

        wealth_values.append(wealth)
        basis_values.append(basis)
        withdrawals_net.append(net_withdraw)
        taxes_paid.append(tax)
        gross_sales.append(sale)

    path = pd.DataFrame(
        {
            "wealth": wealth_values,
            "basis": basis_values,
            "withdrawal_net": withdrawals_net,
            "tax_paid": taxes_paid,
            "gross_sale": gross_sales,
            "return": returns.values,
        },
        index=returns.index,
    )

    return path


# ---------------------------------------------------------------------------
# HELPERS FOR TITLES AND LABELS
# ---------------------------------------------------------------------------

def pretty_asset_name(asset: str) -> str:
    mapping = {
        "msci_world": "MSCI Welt",
        "rexp": "REXP",
        "Gold": "Gold",
        "cash": "Liquidität",
    }
    return mapping.get(asset, asset)


def format_weights(weights: Dict[str, float]) -> str:
    parts = []
    for asset, w in weights.items():
        if w <= 1e-6:
            continue
        parts.append(f"{w*100:.0f} % {pretty_asset_name(asset)}")
    return ", ".join(parts)


def format_chart_title(
    prefix: str,
    port_cfg: PortfolioConfig,
    sim_cfg: SimulationConfig,
    extra: str = "",
) -> str:
    w_str = format_weights(port_cfg.weights)

    if port_cfg.use_fees and port_cfg.annual_fee > 0:
        fee_str = f"{port_cfg.annual_fee*100:.2f} % Gebühren"
    else:
        fee_str = "ohne Gebühren"

    infl_str = "inflationsbereinigt" if port_cfg.use_inflation else "nominal"

    base = (
        f"{prefix}: {w_str}, "
        f"{sim_cfg.annual_withdrawal_rate*100:.1f} % Entnahme p.a., "
        f"{fee_str}, {infl_str}"
    )

    if extra:
        base += f", {extra}"

    return base


# ---------------------------------------------------------------------------
# STYLE SETTINGS
# ---------------------------------------------------------------------------

COMPANY_BLUE = "#003c71"
BAND_50_COLOR = "#7da6d8"
BAND_80_COLOR = "#c0d2ea"
BAND_100_COLOR = "#edf1f7"

MEDIAN_COLOR = COMPANY_BLUE
WORST_COLOR = "#b85c5c"
BEST_COLOR = "#2e7d32"

FIGSIZE_16_9 = (10, 5.625)
DPI_EXPORT = 200

FONT_FAMILY = "DejaVu Sans"
FONT_SIZE_BASE = 11
FONT_SIZE_TITLE = 8
FONT_SIZE_LABEL = 12
FONT_SIZE_TICKS = 10
FONT_SIZE_LEGEND = 9
FONT_SIZE_ANNOT = 9

plt.rcParams.update(
    {
        "font.family": FONT_FAMILY,
        "font.size": FONT_SIZE_BASE,
        "axes.titlesize": FONT_SIZE_TITLE,
        "axes.labelsize": FONT_SIZE_LABEL,
        "xtick.labelsize": FONT_SIZE_TICKS,
        "ytick.labelsize": FONT_SIZE_TICKS,
        "legend.fontsize": FONT_SIZE_LEGEND,
    }
)


def euro_formatter(x, pos):
    return f"{x:,.0f} €".replace(",", ".")


def set_euro_yaxis(ax):
    ax.yaxis.set_major_formatter(FuncFormatter(euro_formatter))


# ---------------------------------------------------------------------------
# CHART FUNCTIONS
# ---------------------------------------------------------------------------

def build_wealth_matrix(
    portfolio_returns: pd.Series,
    sim_cfg: SimulationConfig,
    tax_rate: float,
) -> pd.DataFrame:
    periods = sim_cfg.horizon_years * sim_cfg.periods_per_year
    wealth_paths = []
    start_dates = []

    n_cohorts = len(portfolio_returns) - periods + 1

    for start in range(n_cohorts):
        window = portfolio_returns.iloc[start:start + periods]
        start_dates.append(window.index[0])

        cfg = SimulationConfig(
            annual_withdrawal_rate=sim_cfg.annual_withdrawal_rate,
            initial_wealth=sim_cfg.initial_wealth,
            periods_per_year=sim_cfg.periods_per_year,
            horizon_years=sim_cfg.horizon_years,
        )
        path = simulate_constant_withdrawal(window, cfg, tax_rate=tax_rate)
        wealth_paths.append(path["wealth"].values)

    matrix = pd.DataFrame(wealth_paths).T
    matrix.index = range(1, periods + 1)
    matrix.index.name = "month_in_retirement"
    matrix.columns = pd.to_datetime(start_dates)
    matrix.columns.name = "start_date"
    matrix.attrs["n_cohorts"] = n_cohorts

    return matrix


def summarise_cohorts(matrix: pd.DataFrame) -> pd.DataFrame:
    start_dates = matrix.columns
    terminal_wealth = matrix.iloc[-1]
    min_wealth = matrix.min(axis=0)
    max_wealth = matrix.max(axis=0)
    ruined_mask = (matrix == 0.0).any(axis=0)

    ruin_month = []
    for col in matrix.columns:
        series = matrix[col]
        zero_idx = (series <= 0.0).to_numpy().nonzero()[0]
        ruin_month.append(int(zero_idx[0] + 1) if len(zero_idx) > 0 else None)

    ruin_year = [
        (m / 12.0) if m is not None else None
        for m in ruin_month
    ]

    summary = pd.DataFrame(
        {
            "start_date": start_dates,
            "terminal_wealth": terminal_wealth.values,
            "min_wealth": min_wealth.values,
            "max_wealth": max_wealth.values,
            "ruined": ruined_mask.values,
            "ruin_month": ruin_month,
            "ruin_year": ruin_year,
        }
    )

    summary = summary.sort_values("terminal_wealth", ascending=False).reset_index(drop=True)
    summary.attrs["n_cohorts"] = matrix.attrs.get("n_cohorts", len(summary))
    return summary


def plot_fan_chart(matrix: pd.DataFrame, title: str) -> Figure:
    p0   = matrix.min(axis=1)
    p10  = matrix.quantile(0.10, axis=1)
    p25  = matrix.quantile(0.25, axis=1)
    p50  = matrix.quantile(0.50, axis=1)
    p75  = matrix.quantile(0.75, axis=1)
    p90  = matrix.quantile(0.90, axis=1)
    p100 = matrix.max(axis=1)

    months = matrix.index.values
    years  = months / 12.0

    terminals    = matrix.iloc[-1]
    tw_min       = terminals.min()
    tw_max       = terminals.max()
    tw_mean      = terminals.mean()
    tw_median    = terminals.median()

    start_min    = terminals.idxmin()
    start_max    = terminals.idxmax()
    median_start = (terminals - tw_median).abs().idxmin()

    def fmt_ym(d):
        return d.strftime("%Y-%m") if isinstance(d, pd.Timestamp) else str(d)

    label_min    = f"Schlechtester Verlauf: {euro_formatter(tw_min, None)} ({fmt_ym(start_min)})"
    label_max    = f"Bester Verlauf: {euro_formatter(tw_max, None)} ({fmt_ym(start_max)})"
    label_median = f"Median: {euro_formatter(tw_median, None)} ({fmt_ym(median_start)})"
    label_mean   = f"Mittelwert: {euro_formatter(tw_mean, None)}"

    fig, ax = plt.subplots(figsize=FIGSIZE_16_9, dpi=DPI_EXPORT)

    ax.fill_between(years, p0,  p100, color=BAND_100_COLOR, alpha=1.0, label="100 % Band (min–max)")
    ax.fill_between(years, p10, p90,  color=BAND_80_COLOR,  alpha=1.0, label="80 % Band")
    ax.fill_between(years, p25, p75,  color=BAND_50_COLOR,  alpha=1.0, label="50 % Band")

    ax.plot(years, p50,  color=MEDIAN_COLOR, linewidth=2.0,  label=label_median)
    ax.plot(years, p0,   color=WORST_COLOR,  linestyle="--", linewidth=1.3, label=label_min)
    ax.plot(years, p100, color=BEST_COLOR,   linestyle="--", linewidth=1.3, label=label_max)
    ax.plot([], [],      color="grey",       linestyle=":",  linewidth=1.3, label=label_mean)

    n_cohorts = matrix.attrs.get("n_cohorts")
    if n_cohorts is not None:
        ax.text(0.99, 0.02, f"{n_cohorts} historische Läufe",
                transform=ax.transAxes, ha="right", va="bottom", fontsize=FONT_SIZE_ANNOT)

    ax.set_title(title)
    ax.set_xlabel("Jahre im Ruhestand")
    ax.set_ylabel("Vermögen")
    set_euro_yaxis(ax)
    ax.legend()
    plt.tight_layout()
    return fig


def plot_all_wealth_paths(matrix: pd.DataFrame, title: str) -> Figure:
    months    = matrix.index.values
    years     = months / 12.0
    n_cohorts = matrix.attrs.get("n_cohorts", matrix.shape[1])

    fig, ax = plt.subplots(figsize=FIGSIZE_16_9, dpi=DPI_EXPORT)

    for col in matrix.columns:
        ax.plot(years, matrix[col].values, alpha=0.12, linewidth=1, color=BAND_50_COLOR)

    median_path = matrix.median(axis=1)
    ax.plot(years, median_path.values, linewidth=2.0, color=COMPANY_BLUE, label="Median")

    ax.set_title(title)
    ax.set_xlabel("Jahre im Ruhestand")
    ax.set_ylabel("Vermögen")
    set_euro_yaxis(ax)
    ax.legend()
    ax.text(0.99, 0.02, f"{n_cohorts} historische Läufe",
            transform=ax.transAxes, ha="right", va="bottom", fontsize=FONT_SIZE_ANNOT)

    plt.tight_layout()
    return fig


def success_rate_for_withdrawal(
    portfolio_returns: pd.Series,
    annual_withdrawal_rate: float,
    sim_cfg_template: SimulationConfig,
    tax_rate: float,
) -> float:
    periods   = sim_cfg_template.horizon_years * sim_cfg_template.periods_per_year
    n_cohorts = len(portfolio_returns) - periods + 1
    successes = 0

    for start in range(n_cohorts):
        window = portfolio_returns.iloc[start:start + periods]
        cfg = SimulationConfig(
            annual_withdrawal_rate=annual_withdrawal_rate,
            initial_wealth=sim_cfg_template.initial_wealth,
            periods_per_year=sim_cfg_template.periods_per_year,
            horizon_years=sim_cfg_template.horizon_years,
        )
        path = simulate_constant_withdrawal(window, cfg, tax_rate=tax_rate)
        if (path["wealth"] > 0).all():
            successes += 1

    return successes / n_cohorts if n_cohorts > 0 else float("nan")


def plot_success_curve(
    gross_returns: pd.Series,
    net_returns: pd.Series,
    sim_cfg: SimulationConfig,
    port_cfg_net: PortfolioConfig,
    rate_min: float,
    rate_max: float,
    rate_step: float,
    tax_rate: float,
    show_gross_line: bool = False,
) -> Figure:
    rates = np.arange(rate_min, rate_max + 1e-9, rate_step)

    success_net = [
        success_rate_for_withdrawal(net_returns, r, sim_cfg, tax_rate=tax_rate)
        for r in rates
    ]

    success_gross = None
    if show_gross_line:
        success_gross = [
            success_rate_for_withdrawal(gross_returns, r, sim_cfg, tax_rate=tax_rate)
            for r in rates
        ]

    periods   = sim_cfg.horizon_years * sim_cfg.periods_per_year
    n_cohorts = len(net_returns) - periods + 1

    x_vals          = rates * 100.0
    success_net_pct = np.array(success_net) * 100.0
    success_gross_pct = (
        np.array(success_gross) * 100.0
        if show_gross_line and success_gross is not None
        else None
    )

    fig, ax = plt.subplots(figsize=FIGSIZE_16_9, dpi=DPI_EXPORT)

    if show_gross_line and success_gross_pct is not None:
        ax.plot(x_vals, success_net_pct,   label="mit Gebühren",   color=COMPANY_BLUE)
        ax.plot(x_vals, success_gross_pct, label="ohne Gebühren",  color="grey")
    else:
        ax.plot(x_vals, success_net_pct, color=COMPANY_BLUE)

    ax.set_xlabel("Entnahmesatz in Prozent p.a.", fontsize=FONT_SIZE_LABEL + 1)
    ax.set_ylabel("Historische Erfolgsquote in %", fontsize=FONT_SIZE_LABEL + 1)

    ax.yaxis.set_major_formatter(FuncFormatter(lambda y, _: f"{y:.0f} %"))
    ax.xaxis.set_major_formatter(FuncFormatter(lambda x, _: f"{x:.1f} %"))
    ax.set_ylim(0.0, 105.0)
    ax.tick_params(axis="both", labelsize=FONT_SIZE_TICKS + 1)

    tax_str    = f"mit Steuer {tax_rate*100:.2f} %" if tax_rate > 0 else "ohne Steuer"
    compare_str = ", Vergleich ohne Gebühren" if show_gross_line else ""

    extra = (
        f"Entnahmesätze {rate_min*100:.1f}–{rate_max*100:.1f} %, "
        f"{n_cohorts} Läufe, {tax_str}{compare_str}"
    )

    ax.set_title(
        format_chart_title(
            "Erfolgswahrscheinlichkeit versus Entnahmesatz",
            port_cfg_net,
            sim_cfg,
            extra=extra,
        ),
        fontsize=FONT_SIZE_TITLE + 1,
    )

    if show_gross_line and success_gross_pct is not None:
        legend = ax.legend()
        for text in legend.get_texts():
            text.set_fontsize(FONT_SIZE_LEGEND + 1)

    plt.tight_layout()
    return fig


def compute_terminal_wealth_distribution(
    portfolio_returns: pd.Series,
    sim_cfg: SimulationConfig,
    tax_rate: float,
) -> pd.Series:
    periods   = sim_cfg.horizon_years * sim_cfg.periods_per_year
    n_cohorts = len(portfolio_returns) - periods + 1

    terminal_wealth = []
    start_dates     = []

    for start in range(n_cohorts):
        window = portfolio_returns.iloc[start:start + periods]
        start_dates.append(window.index[0])
        cfg = SimulationConfig(
            annual_withdrawal_rate=sim_cfg.annual_withdrawal_rate,
            initial_wealth=sim_cfg.initial_wealth,
            periods_per_year=sim_cfg.periods_per_year,
            horizon_years=sim_cfg.horizon_years,
        )
        path = simulate_constant_withdrawal(window, cfg, tax_rate=tax_rate)
        terminal_wealth.append(path["wealth"].iloc[-1])

    series = pd.Series(terminal_wealth, index=pd.to_datetime(start_dates))
    series.index.name = "start_date"
    series.attrs["n_cohorts"] = n_cohorts
    return series


def plot_terminal_wealth_hist(
    term_wealth: pd.Series,
    title: str,
    sim_cfg: SimulationConfig,
) -> Figure:
    fig, ax = plt.subplots(figsize=FIGSIZE_16_9, dpi=DPI_EXPORT)
    ax.hist(term_wealth.values, bins=30, color=COMPANY_BLUE, alpha=0.8)
    ax.xaxis.set_major_formatter(FuncFormatter(euro_formatter))

    min_val    = term_wealth.min()
    max_val    = term_wealth.max()
    mean_val   = term_wealth.mean()
    median_val = term_wealth.median()

    start_min    = term_wealth.idxmin()
    start_max    = term_wealth.idxmax()
    start_median = (term_wealth - median_val).abs().idxmin()

    def fmt_ym(d):
        return d.strftime("%Y-%m") if isinstance(d, pd.Timestamp) else str(d)

    label_min    = f"Minimum: {euro_formatter(min_val, None)} ({fmt_ym(start_min)})"
    label_max    = f"Maximum: {euro_formatter(max_val, None)} ({fmt_ym(start_max)})"
    label_mean   = f"Mittelwert: {euro_formatter(mean_val, None)}"
    label_median = f"Median: {euro_formatter(median_val, None)} ({fmt_ym(start_median)})"

    ax.axvline(min_val,    color=WORST_COLOR,   linestyle="--", linewidth=1.2, label=label_min)
    ax.axvline(max_val,    color=BEST_COLOR,    linestyle="--", linewidth=1.2, label=label_max)
    ax.axvline(mean_val,   color="grey",        linestyle=":",  linewidth=1.2, label=label_mean)
    ax.axvline(median_val, color=COMPANY_BLUE,  linestyle="-",  linewidth=1.5, label=label_median)

    n_cohorts = term_wealth.attrs.get("n_cohorts")
    if n_cohorts is not None:
        ax.text(0.99, 0.02, f"{n_cohorts} historische Läufe",
                transform=ax.transAxes, ha="right", va="bottom", fontsize=FONT_SIZE_ANNOT)

    ax.set_title(title)
    ax.set_xlabel(f"Restvermögen nach {sim_cfg.horizon_years} Jahren")
    ax.set_ylabel("Anzahl Kohorten")
    ax.legend(loc="upper right")
    plt.tight_layout()
    return fig


def plot_cumulative_withdrawals(
    path: pd.DataFrame,
    sim_cfg: SimulationConfig,
    title: str,
    inflation: Optional[pd.Series] = None,
    real_mode: bool = True,
) -> Figure:
    months = np.arange(1, len(path) + 1)
    years  = months / sim_cfg.periods_per_year

    cum_withdrawals = path["withdrawal_net"].cumsum()

    fig, ax = plt.subplots(figsize=FIGSIZE_16_9, dpi=DPI_EXPORT)

    label_base = "Kumulierte Entnahmen (real)" if real_mode else "Kumulierte Entnahmen (nominal)"
    ax.plot(years, cum_withdrawals, color=COMPANY_BLUE, linewidth=2.0, label=label_base)

    total_nominal = None

    if inflation is not None and real_mode:
        infl           = inflation.reindex(path.index).fillna(0.0)
        infl_factor    = (1.0 + infl).cumprod()
        withdrawals_nominal = path["withdrawal_net"] * infl_factor
        cum_nominal    = withdrawals_nominal.cumsum()

        ax.plot(years, cum_nominal, color=BAND_50_COLOR, linestyle="--", linewidth=2.0,
                label="Kumulierte Entnahmen (nominal, inflationsindexiert)")
        total_nominal = cum_nominal.iloc[-1]

    ax.set_xlabel("Jahre im Ruhestand")
    ax.set_ylabel("Kumulierte Entnahmen")
    set_euro_yaxis(ax)

    if total_nominal is not None:
        text = (f"Summe nach {sim_cfg.horizon_years} Jahren (nominal): "
                f"{euro_formatter(total_nominal, None)}")
    else:
        total  = cum_withdrawals.iloc[-1]
        suffix = "real" if real_mode else "nominal"
        text   = (f"Summe nach {sim_cfg.horizon_years} Jahren ({suffix}): "
                  f"{euro_formatter(total, None)}")

    ax.text(0.99, 0.02, text, transform=ax.transAxes,
            ha="right", va="bottom", fontsize=FONT_SIZE_ANNOT)

    ax.legend()
    ax.set_title(title)
    plt.tight_layout()
    return fig


# ---------------------------------------------------------------------------
# NIESSBVRAUCH-BERECHNUNG (§ 14 BewG)
# ---------------------------------------------------------------------------

# Vervielfältiger-Tabelle gemäß § 14 BewG (BMF-Tabelle)
# Schlüssel = Restlebenserwartung in Jahren (ganzzahlig), Wert = Vervielfältiger
_VERVIELFAELTIGER_TABLE: Dict[int, float] = {
     1: 0.9,   2: 1.8,   3: 2.7,   4: 3.5,   5: 4.3,
     6: 5.1,   7: 5.8,   8: 6.5,   9: 7.2,  10: 7.8,
    11: 8.4,  12: 8.9,  13: 9.4,  14: 9.9,  15: 10.3,
    16: 10.7, 17: 11.1, 18: 11.4, 19: 11.8, 20: 12.1,
    21: 12.4, 22: 12.7, 23: 12.9, 24: 13.2, 25: 13.4,
    26: 13.6, 27: 13.8, 28: 14.0, 29: 14.2, 30: 14.3,
    31: 14.5, 32: 14.6, 33: 14.8, 34: 14.9, 35: 15.0,
    36: 15.1, 37: 15.2, 38: 15.3, 39: 15.4, 40: 15.5,
    41: 15.6, 42: 15.6, 43: 15.7, 44: 15.8, 45: 15.8,
    46: 15.9, 47: 15.9, 48: 16.0, 49: 16.0, 50: 16.0,
    51: 16.1, 52: 16.1, 53: 16.1, 54: 16.1, 55: 16.1,
    56: 16.1, 57: 16.1, 58: 16.1, 59: 16.0, 60: 16.0,
    61: 16.0, 62: 15.9, 63: 15.9, 64: 15.8, 65: 15.8,
    66: 15.7, 67: 15.6, 68: 15.5, 69: 15.4, 70: 15.3,
    71: 15.2, 72: 15.0, 73: 14.9, 74: 14.7, 75: 14.5,
    76: 14.3, 77: 14.1, 78: 13.9, 79: 13.7, 80: 13.4,
}


def get_vervielfaeltiger(restlebenserwartung_jahre: float) -> float:
    """
    Interpoliert den Vervielfältiger aus der §-14-BewG-Tabelle für eine
    beliebige (auch nicht-ganzzahlige) Restlebenserwartung in Jahren.

    Unterhalb von 1 Jahr  → 0.9  (Minimalwert der Tabelle).
    Oberhalb von 80 Jahren → 13.4 (Maximalwert, BMF-Deckelung).
    """
    if restlebenserwartung_jahre <= 1:
        return _VERVIELFAELTIGER_TABLE[1]

    lo = int(restlebenserwartung_jahre)
    hi = lo + 1

    if lo >= 80:
        return _VERVIELFAELTIGER_TABLE[80]

    v_lo  = _VERVIELFAELTIGER_TABLE.get(lo, _VERVIELFAELTIGER_TABLE[80])
    v_hi  = _VERVIELFAELTIGER_TABLE.get(hi, _VERVIELFAELTIGER_TABLE[80])
    frac  = restlebenserwartung_jahre - lo
    return v_lo + frac * (v_hi - v_lo)


def compute_niessbrauch(
    initial_wealth: float,
    weights: Dict[str, float],
    dividendenrendite: float,
    kuponrendite: float,
    goldrendite: float,
    restlebenserwartung: float,
    cap_jahreswert: bool = True,
) -> Dict[str, float]:
    """
    Berechnet den steuerlichen Nießbrauchswert nach § 14 BewG.

    Jahreswert
    ----------
    Gewichtete laufende Rendite des Portfolios × Startvermögen:
        jahreswert = Σ (gewicht_i × rendite_i) × startvermögen

    Gemäß § 16 BewG ist der Jahreswert auf 1/18,6 des Kapitals gedeckelt.

    Nießbrauchswert
    ---------------
        nwert = jahreswert × vervielfältiger
    """
    w_aktien = weights.get("msci_world", 0.0)
    w_renten = weights.get("rexp",       0.0)
    w_gold   = weights.get("Gold",       0.0)
    # cash / Liquidität bringt 0 % laufenden Ertrag

    rendite_gewichtet = (
        w_aktien * dividendenrendite
        + w_renten * kuponrendite
        + w_gold   * goldrendite
    )

    jahreswert_roh = rendite_gewichtet * initial_wealth

    # § 16 BewG: Deckelung auf 1/18,6 des Vermögens
    jahreswert_cap = initial_wealth / 18.6
    cap_aktiv      = cap_jahreswert and (jahreswert_roh > jahreswert_cap)
    jahreswert     = jahreswert_cap if cap_aktiv else jahreswert_roh

    vervielfaeltiger = get_vervielfaeltiger(restlebenserwartung)
    niessbrauchswert = jahreswert * vervielfaeltiger

    return {
        "jahreswert_roh":    jahreswert_roh,
        "jahreswert":        jahreswert,
        "jahreswert_cap":    jahreswert_cap,
        "vervielfaeltiger":  vervielfaeltiger,
        "niessbrauchswert":  niessbrauchswert,
        "rendite_gewichtet": rendite_gewichtet,
        "cap_aktiv":         cap_aktiv,
    }


# ---------------------------------------------------------------------------
# DEPOT-IMPORT FÜR DEN NIESSBRAUCH
# ---------------------------------------------------------------------------

# Spalten, die auf Personenbezug hindeuten. Enthält die Datei eine davon,
# wird der Upload abgelehnt (keine Kundendaten im Tool).
DEPOT_GESPERRTE_FRAGMENTE = (
    "kunde", "inhaber", "mandant", "iban", "bic", "depotnummer", "depot_nr",
    "depotnr", "kontonummer", "konto_nr", "kontonr", "geburts", "steuernummer",
    "steuer_id", "steueridentifikation", "adresse", "strasse", "telefon",
    "mail", "vorname", "nachname", "berater",
)

DEPOT_PFLICHTSPALTEN = ("name", "assetklasse", "kurswert_eur", "ertragsart")

DEPOT_SPALTEN_ALIASE = {
    "isin": "isin",
    "wkn": "wkn",
    "name": "name",
    "wp_name": "name",
    "bezeichnung": "name",
    "assetklasse": "assetklasse",
    "bestand": "bestand",
    "nominal": "bestand",
    "waehrung": "waehrung",
    "kurs": "kurs",
    "devisenkurs": "devisenkurs",
    "aktueller_devisenkurs": "devisenkurs",
    "kurswert_eur": "kurswert_eur",
    "kurswert": "kurswert_eur",
    "ertrag_basis": "ertrag_basis",
    "ertrag_je_einheit": "ertrag_je_einheit",
    "auszahlung_in_eur_pro_titel": "ertrag_je_einheit",
    "ertrag_pa_eur": "ertrag_pa_eur",
    "auszahlung_in_eur": "ertrag_pa_eur",
    "ertragsart": "ertragsart",
    "zahlungsfrequenz": "zahlungsfrequenz",
    "faelligkeit": "faelligkeit",
    "ertrag_quelle": "ertrag_quelle",
    "stichtag": "stichtag",
}

_RENTEN_KLASSEN = {"renten", "anleihen", "anleihe", "bonds", "renten_eur"}
_KEIN_ERTRAG_ARTEN = {"thesaurierend", "kein_ertrag", "keine", "thesaurierung"}


def _normalise_column(col: str) -> str:
    """Spaltenname auf Kleinschreibung ohne Umlaute und Sonderzeichen bringen."""
    s = str(col).strip().lower()
    for a, b in (("ä", "ae"), ("ö", "oe"), ("ü", "ue"), ("ß", "ss")):
        s = s.replace(a, b)
    for ch in ("[", "]", "%", "(", ")", ".", "/", "€"):
        s = s.replace(ch, " ")
    s = "_".join(s.split())
    return s.strip("_")


def _to_float(value) -> float:
    """Robuster Parser für deutsche Zahlenformate ('1.234,56', '35.000', '97,87')."""
    if value is None:
        return float("nan")
    if isinstance(value, (int, float, np.integer, np.floating)):
        return float(value)

    s = str(value).strip().replace("\u00a0", "").replace(" ", "")
    s = s.replace("€", "").replace("%", "")
    if s == "" or s.lower() in {"nan", "none", "-", "n/a"}:
        return float("nan")

    if "," in s:
        # Komma = Dezimaltrenner, Punkt = Tausendertrenner
        s = s.replace(".", "").replace(",", ".")
    elif "." in s:
        # Nur Punkte: '35.000' ist Tausender, '97.87' ist Dezimal
        teile = s.split(".")
        if len(teile) > 1 and all(len(t) == 3 for t in teile[1:]) and len(teile[0]) <= 3:
            s = s.replace(".", "")

    try:
        return float(s)
    except ValueError:
        return float("nan")


def pruefe_personenbezug(columns) -> list:
    """Gibt alle Spalten zurück, die auf personenbezogene Daten hindeuten."""
    treffer = []
    for col in columns:
        norm = _normalise_column(col)
        if any(frag in norm for frag in DEPOT_GESPERRTE_FRAGMENTE):
            treffer.append(str(col))
    return treffer


def load_depot_csv(source) -> pd.DataFrame:
    """
    Liest eine Depot-CSV (Semikolon, UTF-8) ein, normalisiert die Spaltennamen
    und lehnt Dateien mit personenbezogenen Spalten ab.

    Raises
    ------
    PermissionError
        Wenn die Datei mutmaßlich personenbezogene Spalten enthält.
    KeyError
        Wenn Pflichtspalten fehlen.
    ValueError
        Wenn die Datei nicht als CSV lesbar ist.
    """
    if hasattr(source, "seek"):
        source.seek(0)

    df = None
    for sep in (";", ",", "\t"):
        if hasattr(source, "seek"):
            source.seek(0)
        try:
            kandidat = pd.read_csv(source, sep=sep, dtype=str, encoding="utf-8-sig")
        except Exception:
            continue
        if kandidat.shape[1] > 1:
            df = kandidat
            break

    if df is None:
        raise ValueError(
            "Die Datei konnte nicht als CSV gelesen werden. "
            "Erwartet wird eine semikolongetrennte Datei in UTF-8."
        )

    treffer = pruefe_personenbezug(df.columns)
    if treffer:
        raise PermissionError(
            "Die Datei enthält mutmaßlich personenbezogene Spalten: "
            + ", ".join(treffer)
            + ". Bitte diese Spalten vor dem Upload entfernen."
        )

    umbenennung = {}
    for col in df.columns:
        norm = _normalise_column(col)
        umbenennung[col] = DEPOT_SPALTEN_ALIASE.get(norm, norm)
    df = df.rename(columns=umbenennung)
    df = df.loc[:, ~df.columns.duplicated()]

    fehlend = [c for c in DEPOT_PFLICHTSPALTEN if c not in df.columns]
    if fehlend:
        raise KeyError("Pflichtspalten fehlen: " + ", ".join(fehlend))

    if "isin" not in df.columns and "wkn" not in df.columns:
        raise KeyError("Es muss mindestens eine der Spalten 'isin' oder 'wkn' vorhanden sein.")

    for col in ("bestand", "kurs", "devisenkurs", "kurswert_eur",
                "ertrag_je_einheit", "ertrag_pa_eur"):
        if col in df.columns:
            df[col] = df[col].map(_to_float)

    for col in ("assetklasse", "ertragsart", "ertrag_basis"):
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.lower()

    if "devisenkurs" not in df.columns:
        df["devisenkurs"] = 1.0
    df["devisenkurs"] = df["devisenkurs"].fillna(1.0).replace(0.0, 1.0)

    df = df[df["kurswert_eur"].notna()].reset_index(drop=True)
    return df


def berechne_ertraege(df: pd.DataFrame) -> tuple:
    """
    Ermittelt je Position den Jahresertrag in EUR und prüft ihn gegen einen
    gelieferten Wert.

    Returns
    -------
    tuple
        (DataFrame mit zusätzlicher Spalte 'ertrag_pa_eur_calc', Liste von Warnungen)
    """
    warnungen = []
    out = df.copy()
    berechnet = []

    for i, row in out.iterrows():
        art   = str(row.get("ertragsart", "")).strip().lower()
        basis = str(row.get("ertrag_basis", "")).strip().lower()
        je    = _to_float(row.get("ertrag_je_einheit"))
        best  = _to_float(row.get("bestand"))
        fx    = _to_float(row.get("devisenkurs"))
        fx    = fx if np.isfinite(fx) and fx != 0.0 else 1.0
        gelie = _to_float(row.get("ertrag_pa_eur"))
        name  = row.get("name", f"Zeile {i + 2}")

        if art in _KEIN_ERTRAG_ARTEN:
            berechnet.append(0.0)
            if art.startswith("thesaur"):
                warnungen.append(
                    f"{name}: thesaurierend, Jahreswert vereinbarungsgemäß mit 0 € angesetzt."
                )
            continue

        wert = float("nan")
        if basis == "betrag_gesamt":
            wert = gelie
        elif basis == "je_stueck" and np.isfinite(je) and np.isfinite(best):
            wert = je * best / fx
        elif basis == "prozent_nominal" and np.isfinite(je) and np.isfinite(best):
            wert = best * je / 100.0 / fx

        if not np.isfinite(wert):
            if np.isfinite(gelie):
                wert = gelie
                if basis not in ("betrag_gesamt", "", "nan"):
                    warnungen.append(
                        f"{name}: Ertrag nicht aus 'ertrag_basis' herleitbar, "
                        f"gelieferter Wert 'ertrag_pa_eur' übernommen."
                    )
            else:
                wert = 0.0
                warnungen.append(f"{name}: kein Ertrag ermittelbar, mit 0 € angesetzt.")
        elif np.isfinite(gelie) and basis != "betrag_gesamt":
            if abs(wert) > 1e-9 and abs(gelie - wert) / abs(wert) > 0.01:
                warnungen.append(
                    f"{name}: gelieferter Jahresertrag {_de_num(gelie)} € weicht von der "
                    f"Berechnung {_de_num(wert)} € ab (Faktor {gelie / wert:.2f}). "
                    f"Bitte 'ertrag_basis' und 'ertrag_je_einheit' prüfen."
                )

        berechnet.append(float(wert))

    out["ertrag_pa_eur_calc"] = berechnet
    return out, warnungen


def _de_num(x: float, nachkomma: int = 2) -> str:
    """Zahl im deutschen Format, ohne Währungszeichen."""
    if not np.isfinite(x):
        return "–"
    s = f"{x:,.{nachkomma}f}"
    return s.replace(",", "#").replace(".", ",").replace("#", ".")


def compute_niessbrauch_aus_depot(
    df: pd.DataFrame,
    restlebenserwartung: float,
    wiederanlagerendite: float,
    stichtag: Optional[pd.Timestamp] = None,
    cap_jahreswert: bool = True,
) -> Dict[str, object]:
    """
    Nießbrauchswert auf Basis der konkreten Depotausschüttungen.

    Rentenpositionen: Kupon bis zur Fälligkeit, danach Wiederanlage des
    Rückzahlungsbetrags (Nominal zu 100 %) zur unterstellten Wiederanlagerendite.
    Beide Phasen werden über die Restlebenserwartung T zeitgewichtet gemittelt:

        jahreswert = (min(t_f, T) * ertrag_heute + max(0, T - t_f) * ertrag_wiederanlage) / T

    Die Gewichtung ist undiskontiert und damit eine Näherung, weil der
    Vervielfältiger nach § 14 BewG ein Barwertfaktor ist. Späte Perioden werden
    dadurch leicht übergewichtet. Die exakte Alternative wäre eine Aufteilung in
    einen befristeten und einen aufgeschobenen Nießbrauch mit getrennten
    Vervielfältigern; das ist mit dem Steuerberater abzustimmen.

    Deckelung nach § 16 BewG auf Depotwert / 18,6 (Summe der Kurswerte aus der Datei).
    """
    if stichtag is None:
        if "stichtag" in df.columns and df["stichtag"].notna().any():
            stichtag = pd.to_datetime(df["stichtag"].dropna().iloc[0], errors="coerce")
    if stichtag is None or pd.isna(stichtag):
        stichtag = pd.Timestamp.today().normalize()

    T = float(restlebenserwartung)
    daten, warnungen = berechne_ertraege(df)

    jahreswerte       = []
    restlaufzeiten    = []
    ertrag_nach_liste = []

    for _, row in daten.iterrows():
        ertrag_heute = float(row["ertrag_pa_eur_calc"])
        klasse       = str(row.get("assetklasse", "")).strip().lower()
        name         = row.get("name", "")
        rl           = float("nan")
        e_nach       = float("nan")

        if klasse in _RENTEN_KLASSEN:
            faellig = pd.to_datetime(row.get("faelligkeit"), errors="coerce")
            if pd.isna(faellig):
                warnungen.append(
                    f"{name}: keine Fälligkeit angegeben, Kupon wird unbefristet "
                    f"fortgeschrieben (keine Wiederanlage unterstellt)."
                )
                jahreswerte.append(ertrag_heute)
                restlaufzeiten.append(float("nan"))
                ertrag_nach_liste.append(float("nan"))
                continue

            t_f = max(0.0, (faellig - stichtag).days / 365.25)
            rl  = t_f

            nominal = _to_float(row.get("bestand"))
            fx      = _to_float(row.get("devisenkurs"))
            fx      = fx if np.isfinite(fx) and fx != 0.0 else 1.0
            rueckzahlung = (nominal / fx) if np.isfinite(nominal) else float(row["kurswert_eur"])
            e_nach = rueckzahlung * wiederanlagerendite

            if T <= 0:
                jw = ertrag_heute
            elif t_f >= T:
                jw = ertrag_heute
            else:
                jw = (t_f * ertrag_heute + (T - t_f) * e_nach) / T
            jahreswerte.append(jw)
        else:
            jahreswerte.append(ertrag_heute)

        restlaufzeiten.append(rl)
        ertrag_nach_liste.append(e_nach)

    daten["restlaufzeit_jahre"]  = restlaufzeiten
    daten["ertrag_wiederanlage"] = ertrag_nach_liste
    daten["jahreswert_effektiv"] = jahreswerte

    vermoegen      = float(daten["kurswert_eur"].sum())
    jahreswert_roh = float(daten["jahreswert_effektiv"].sum())
    ertrag_heute   = float(daten["ertrag_pa_eur_calc"].sum())

    jahreswert_cap = vermoegen / 18.6
    cap_aktiv      = cap_jahreswert and (jahreswert_roh > jahreswert_cap)
    jahreswert     = jahreswert_cap if cap_aktiv else jahreswert_roh

    vervielfaeltiger = get_vervielfaeltiger(T)

    renten = daten[daten["assetklasse"].isin(_RENTEN_KLASSEN)]
    if len(renten) > 0 and renten["restlaufzeit_jahre"].notna().any():
        maske = renten["restlaufzeit_jahre"].notna()
        gew   = renten.loc[maske, "kurswert_eur"]
        rlz   = float((renten.loc[maske, "restlaufzeit_jahre"] * gew).sum() / gew.sum()) \
            if gew.sum() > 0 else float("nan")
    else:
        rlz = float("nan")

    return {
        "daten":               daten,
        "warnungen":           warnungen,
        "stichtag":            stichtag,
        "vermoegen":           vermoegen,
        "ertrag_heute":        ertrag_heute,
        "jahreswert_roh":      jahreswert_roh,
        "jahreswert":          jahreswert,
        "jahreswert_cap":      jahreswert_cap,
        "cap_aktiv":           cap_aktiv,
        "vervielfaeltiger":    vervielfaeltiger,
        "niessbrauchswert":    jahreswert * vervielfaeltiger,
        "rendite_gewichtet":   (jahreswert_roh / vermoegen) if vermoegen > 0 else 0.0,
        "restlaufzeit_renten": rlz,
        "rentenanteil":        float(renten["kurswert_eur"].sum() / vermoegen) if vermoegen > 0 else 0.0,
    }


DEPOT_MUSTER_CSV = (
    "isin;wkn;name;assetklasse;bestand;waehrung;kurs;devisenkurs;kurswert_eur;"
    "ertrag_basis;ertrag_je_einheit;ertrag_pa_eur;ertragsart;zahlungsfrequenz;"
    "faelligkeit;ertrag_quelle;stichtag\n"
    "US00287Y1091;US00287Y1091;AbbVie Inc.;aktien;116;USD;258,15;1,1652;25.700,79;"
    "je_stueck;6,63;660,36;dividende;quartal;;letzte_ausschuettung;2026-08-28\n"
    "DE0008404005;840400;Allianz SE;aktien;67;EUR;450,50;1;30.183,50;"
    "je_stueck;17,03;1.140,89;dividende;jaehrlich;;letzte_ausschuettung;2026-08-28\n"
    "XS3229496180;;Deutsche Post IHS 3,0 % 25.11.31;renten;35.000;EUR;97,87;1;35.049,52;"
    "prozent_nominal;3,00;1.050,00;kupon;jaehrlich;2031-11-25;prospekt_kupon;2026-08-28\n"
    "XS3274805376;;E.ON IHS 3,448 % 19.01.34;renten;35.000;EUR;96,77;1;34.601,59;"
    "prozent_nominal;3,448;1.206,80;kupon;jaehrlich;2034-01-19;prospekt_kupon;2026-08-28\n"
    "DE000A0S9GB0;A0S9GB;Gold (Xetra-Gold);edelmetall;300;EUR;72,50;1;21.750,00;"
    "betrag_gesamt;;0,00;kein_ertrag;;;;2026-08-28\n"
    ";;Liquiditaet;liquiditaet;;EUR;;1;12.500,00;"
    "betrag_gesamt;;0,00;kein_ertrag;;;;2026-08-28\n"
)


# ---------------------------------------------------------------------------
# LOGIN AUTHENTICATION
# ---------------------------------------------------------------------------

def check_login():
    import streamlit as st

    USERS = st.secrets["passwords"]

    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False
        st.session_state.username  = ""

    def verify_password():
        username = st.session_state.get("username_input", "")
        password = st.session_state.get("password_input", "")
        if username in USERS and USERS[username] == password:
            st.session_state.logged_in = True
            st.session_state.username  = username
            return True
        return False

    if not st.session_state.logged_in:
        st.title("Ausschüttungs-VV Rechner | Fürst Fugger Privatbank")
        st.write("Bitte melden Sie sich an, um fortzufahren.")

        st.text_input("Benutzername", key="username_input")
        st.text_input("Passwort", type="password", key="password_input")

        if st.button("Einloggen"):
            if verify_password():
                st.success("Erfolgreich eingeloggt!")
                st.rerun()
            else:
                st.error("❌ Falscher Benutzername oder Passwort")

        return False
    else:
        with st.sidebar:
            st.write(f"👤 Angemeldet als: **{st.session_state.username}**")
            if st.button("Ausloggen"):
                st.session_state.logged_in = False
                st.session_state.username  = ""
                st.rerun()

        return True


# ---------------------------------------------------------------------------
# STREAMLIT APP
# ---------------------------------------------------------------------------

def run_app() -> None:
    import tempfile
    import streamlit as st

    if not check_login():
        st.stop()

    st.set_page_config(
        page_title="Verrentungs-Simulation (MSCI World + REXP + Gold)",
        layout="wide",
    )
    st.title("Verrentungs-Simulation: MSCI World, REXP und Gold")
    st.markdown(
        "Diese Anwendung ist für Beratungsgespräche gedacht. "
        "Sie dient der Visualisierung Ihrer Möglichkeiten im Beratungsgespräch und ist als Ergänzung "
        "zu unserer Broschüre konzipiert. Bitte beachten Sie, dass alle rechtsverbindlichen Details "
        "sowie die wichtigen Risikohinweise vollständig in der begleitenden Broschüre enthalten sind. "
        'Änderungen werden erst nach Klick auf „Berechnung starten" übernommen.'
    )

    default_file_exists = DATA_FILE.is_file()

    def _euro_str(x: float) -> str:
        return f"{x:,.0f} €".replace(",", ".")

    def _fmt_date_ym(d: pd.Timestamp) -> str:
        if isinstance(d, pd.Timestamp):
            return d.strftime("%Y-%m")
        return str(d)

    def _simulate_path_for_start_date(
        portfolio_returns: pd.Series,
        sim_cfg: SimulationConfig,
        tax_rate: float,
        start_date: pd.Timestamp,
    ) -> pd.DataFrame:
        periods = sim_cfg.horizon_years * sim_cfg.periods_per_year
        idx     = portfolio_returns.index.get_indexer([pd.to_datetime(start_date)])[0]
        if idx < 0:
            raise ValueError("Startdatum der Kohorte wurde in der Renditereihe nicht gefunden.")
        window = portfolio_returns.iloc[idx:idx + periods]
        cfg = SimulationConfig(
            annual_withdrawal_rate=sim_cfg.annual_withdrawal_rate,
            initial_wealth=sim_cfg.initial_wealth,
            periods_per_year=sim_cfg.periods_per_year,
            horizon_years=sim_cfg.horizon_years,
        )
        return simulate_constant_withdrawal(window, cfg, tax_rate=tax_rate)

    def _init_defaults() -> None:
        st.session_state.setdefault("erweitert", False)

        if default_file_exists:
            st.session_state.setdefault("datenquelle", "Standardpfad verwenden")
        else:
            st.session_state.setdefault("datenquelle", "Excel-Datei hochladen")

        st.session_state.setdefault("excel_path_input", str(DATA_FILE))
        st.session_state.setdefault("uploaded_file", None)

        st.session_state.setdefault("sheet_name",  "Import_Daten")
        st.session_state.setdefault("date_column", "Dates")
        st.session_state.setdefault("cpi_column",  "Inflation DE")
        st.session_state.setdefault("col_msci",    "NDDUWI Index")
        st.session_state.setdefault("col_rexp",    "REXP Index")
        st.session_state.setdefault("col_gold",    "Gold")

        st.session_state.setdefault("w_msci_pct", float(PORTFOLIO_WEIGHTS.get("msci_world", 0.60) * 100.0))
        st.session_state.setdefault("w_rexp_pct", float(PORTFOLIO_WEIGHTS.get("rexp",       0.35) * 100.0))
        st.session_state.setdefault("w_gold_pct", float(PORTFOLIO_WEIGHTS.get("Gold",       0.05) * 100.0))

        st.session_state.setdefault("apply_fees",    APPLY_FEES)
        st.session_state.setdefault("annual_fee_pct", float(ANNUAL_FEE * 100.0))

        st.session_state.setdefault("apply_tax",    APPLY_TAX)
        st.session_state.setdefault("tax_rate_pct", float(CAPITAL_GAINS_TAX_RATE * 100.0))

        st.session_state.setdefault("use_inflation",  USE_INFLATION)
        st.session_state.setdefault("withdrawal_eur", float(WITHDRAWAL_RATE * INITIAL_WEALTH))
        st.session_state.setdefault("initial_wealth", float(INITIAL_WEALTH))
        st.session_state.setdefault("horizon_years",  int(HORIZON_YEARS))

        st.session_state.setdefault("rate_min_pct",  float(SUCCESS_RATE_MIN  * 100.0))
        st.session_state.setdefault("rate_max_pct",  float(SUCCESS_RATE_MAX  * 100.0))
        st.session_state.setdefault("rate_step_pp",  float(SUCCESS_RATE_STEP * 100.0))
        st.session_state.setdefault("show_gross_line", False)

        st.session_state.setdefault("results",                None)
        st.session_state.setdefault("last_config_fingerprint", None)

        # Nießbrauch-Defaults
        st.session_state.setdefault("nb_restleben",      20.0)
        st.session_state.setdefault("nb_dividende_pct",   1.0)
        st.session_state.setdefault("nb_kupon_pct",        3.0)
        st.session_state.setdefault("nb_gold_pct",         0.0)

        # Depot-Import für den Nießbrauch
        st.session_state.setdefault("nb_wiederanlage_pct", 3.0)
        st.session_state.setdefault("nb_depot_fehler",     None)

    def _reset_settings() -> None:
        keys_to_clear = [
            "erweitert", "datenquelle", "excel_path_input", "uploaded_file",
            "sheet_name", "date_column", "cpi_column", "col_msci", "col_rexp", "col_gold",
            "w_msci_pct", "w_rexp_pct", "w_gold_pct",
            "apply_fees", "annual_fee_pct", "apply_tax", "tax_rate_pct",
            "use_inflation", "withdrawal_eur", "initial_wealth", "horizon_years",
            "rate_min_pct", "rate_max_pct", "rate_step_pp", "show_gross_line",
            "results", "last_config_fingerprint",
            "nb_restleben", "nb_dividende_pct", "nb_kupon_pct", "nb_gold_pct",
            "nb_wiederanlage_pct", "nb_depot_fehler", "nb_depot_upload",
        ]
        for k in keys_to_clear:
            if k in st.session_state:
                del st.session_state[k]
        _init_defaults()

    def _resolve_excel_path() -> Optional[Path]:
        if st.session_state["datenquelle"] == "Standardpfad verwenden":
            return DATA_FILE if default_file_exists else None
        if st.session_state["datenquelle"] == "Pfad zur Excel-Datei eingeben":
            p = Path(st.session_state["excel_path_input"])
            return p if p.is_file() else None
        return None

    def _load_panel_from_source() -> pd.DataFrame:
        asset_columns = {
            "msci_world": st.session_state["col_msci"],
            "rexp":       st.session_state["col_rexp"],
            "Gold":       st.session_state["col_gold"],
        }

        datenquelle = st.session_state["datenquelle"]

        if datenquelle in ("Standardpfad verwenden", "Pfad zur Excel-Datei eingeben"):
            excel_path = _resolve_excel_path()
            if excel_path is None:
                raise FileNotFoundError("Excel-Datei wurde nicht gefunden.")
            data_cfg = DataConfig(
                excel_path=excel_path,
                sheet_name=st.session_state["sheet_name"],
                date_column=st.session_state["date_column"],
                cpi_column=st.session_state["cpi_column"],
                asset_columns=asset_columns,
            )
            return load_market_data(data_cfg)

        uploaded_file = st.session_state.get("uploaded_file", None)
        if uploaded_file is None:
            raise ValueError("Keine Excel-Datei hochgeladen.")

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(uploaded_file.getbuffer())
            temp_path = Path(tmp.name)

        data_cfg = DataConfig(
            excel_path=temp_path,
            sheet_name=st.session_state["sheet_name"],
            date_column=st.session_state["date_column"],
            cpi_column=st.session_state["cpi_column"],
            asset_columns=asset_columns,
        )
        return load_market_data(data_cfg)

    def _current_config_fingerprint() -> tuple:
        uploaded     = st.session_state.get("uploaded_file", None)
        uploaded_sig = None
        if uploaded is not None:
            try:
                uploaded_sig = (uploaded.name, uploaded.size)
            except Exception:
                uploaded_sig = "uploaded"

        return (
            st.session_state.get("datenquelle"),
            st.session_state.get("excel_path_input"),
            uploaded_sig,
            st.session_state.get("sheet_name"),
            st.session_state.get("date_column"),
            st.session_state.get("cpi_column"),
            st.session_state.get("col_msci"),
            st.session_state.get("col_rexp"),
            st.session_state.get("col_gold"),
            float(st.session_state.get("w_msci_pct")),
            float(st.session_state.get("w_rexp_pct")),
            float(st.session_state.get("w_gold_pct")),
            bool(st.session_state.get("apply_fees")),
            float(st.session_state.get("annual_fee_pct")),
            bool(st.session_state.get("apply_tax")),
            float(st.session_state.get("tax_rate_pct")),
            bool(st.session_state.get("use_inflation")),
            float(st.session_state.get("withdrawal_eur")),
            float(st.session_state.get("initial_wealth")),
            int(st.session_state.get("horizon_years")),
            float(st.session_state.get("rate_min_pct")),
            float(st.session_state.get("rate_max_pct")),
            float(st.session_state.get("rate_step_pp")),
            bool(st.session_state.get("show_gross_line")),
        )

    _init_defaults()

    # ---------------------------------------------------------------------------
    # SIDEBAR
    # ---------------------------------------------------------------------------

    st.sidebar.header("Einstellungen")

    c_btn1, c_btn2 = st.sidebar.columns(2)
    with c_btn1:
        run_clicked = st.button("Berechnung starten", type="primary")
    with c_btn2:
        if st.button("Zurücksetzen"):
            _reset_settings()
            st.rerun()

    st.sidebar.markdown("")

    st.session_state["erweitert"] = st.sidebar.checkbox(
        "Erweiterte Einstellungen anzeigen",
        value=st.session_state["erweitert"],
        help="Wenn deaktiviert, werden nur die wichtigsten Parameter angezeigt.",
    )

    if default_file_exists:
        daten_optionen = ("Standardpfad verwenden", "Excel-Datei hochladen")
    else:
        daten_optionen = ("Excel-Datei hochladen",)

    if st.session_state["erweitert"]:
        daten_optionen = daten_optionen + ("Pfad zur Excel-Datei eingeben",)

    st.session_state["datenquelle"] = st.sidebar.radio(
        "Datenquelle",
        daten_optionen,
        index=list(daten_optionen).index(st.session_state["datenquelle"])
        if st.session_state["datenquelle"] in daten_optionen
        else 0,
    )

    if st.session_state["datenquelle"] == "Standardpfad verwenden":
        st.sidebar.caption(f"Verwendeter Pfad: {DATA_FILE}")
        if not default_file_exists:
            st.sidebar.warning("Standardpfad nicht gefunden. Bitte Datei hochladen oder Pfad eingeben.")
    elif st.session_state["datenquelle"] == "Pfad zur Excel-Datei eingeben":
        st.session_state["excel_path_input"] = st.sidebar.text_input(
            "Excel-Pfad", value=st.session_state["excel_path_input"]
        )
    else:
        st.session_state["uploaded_file"] = st.sidebar.file_uploader(
            "Excel-Datei auswählen", type=["xlsx", "xls"]
        )

    if st.session_state["erweitert"]:
        with st.sidebar.expander("Datenstruktur (Sheet und Spalten)", expanded=False):
            st.session_state["sheet_name"]  = st.text_input("Sheet-Name",               value=st.session_state["sheet_name"])
            st.session_state["date_column"] = st.text_input("Datums-Spalte",             value=st.session_state["date_column"])
            st.session_state["cpi_column"]  = st.text_input("Inflations-Spalte (CPI)",   value=st.session_state["cpi_column"])
            st.session_state["col_msci"]    = st.text_input("MSCI World Spalte",         value=st.session_state["col_msci"])
            st.session_state["col_rexp"]    = st.text_input("REXP Spalte",               value=st.session_state["col_rexp"])
            st.session_state["col_gold"]    = st.text_input("Gold Spalte",               value=st.session_state["col_gold"])

    st.sidebar.divider()
    st.sidebar.subheader("Portfolio")

    st.session_state["w_msci_pct"] = st.sidebar.number_input(
        "MSCI World (%)", min_value=0.0, max_value=100.0,
        value=float(st.session_state["w_msci_pct"]), step=1.0, format="%.1f",
    )
    st.session_state["w_rexp_pct"] = st.sidebar.number_input(
        "REXP (%)", min_value=0.0, max_value=100.0,
        value=float(st.session_state["w_rexp_pct"]), step=1.0, format="%.1f",
    )
    st.session_state["w_gold_pct"] = st.sidebar.number_input(
        "Gold (%)", min_value=0.0, max_value=100.0,
        value=float(st.session_state["w_gold_pct"]), step=1.0, format="%.1f",
    )

    sum_weights_pct = float(
        st.session_state["w_msci_pct"]
        + st.session_state["w_rexp_pct"]
        + st.session_state["w_gold_pct"]
    )
    rest_cash_pct = float(max(0.0, 100.0 - sum_weights_pct))

    m1, m2 = st.sidebar.columns(2)
    with m1:
        st.metric("Summe",      f"{sum_weights_pct:.0f} %")
    with m2:
        st.metric("Liquidität", f"{rest_cash_pct:.0f} %")

    if sum_weights_pct > 100.0 + 1e-9:
        st.sidebar.warning(
            f"Achtung: Die Summe der Gewichte beträgt {sum_weights_pct:.1f} % und ist größer als 100 %.\n"
            "Bitte Gewichte reduzieren. Die Berechnung wird sonst nicht gestartet."
        )
    else:
        st.sidebar.caption("Wenn die Summe unter 100 % liegt, wird der Rest automatisch als Liquidität ergänzt.")

    st.sidebar.divider()
    st.sidebar.subheader("Kosten und Entnahme")

    st.session_state["apply_fees"] = st.sidebar.checkbox("Gebühren berücksichtigen", value=st.session_state["apply_fees"])
    st.session_state["annual_fee_pct"] = st.sidebar.number_input(
        "Gebühr pro Jahr (%)",
        min_value=0.0, max_value=10.0,
        value=float(st.session_state["annual_fee_pct"]),
        step=0.01, format="%.2f",
        disabled=not st.session_state["apply_fees"],
    )

    st.session_state["apply_tax"] = st.sidebar.checkbox("Steuer berücksichtigen", value=st.session_state["apply_tax"])
    st.session_state["tax_rate_pct"] = st.sidebar.number_input(
        "Steuersatz (%)",
        min_value=0.0, max_value=50.0,
        value=float(st.session_state["tax_rate_pct"]),
        step=0.10, format="%.2f",
        disabled=not st.session_state["apply_tax"],
    )

    st.session_state["use_inflation"] = st.sidebar.radio(
        "Rechnungsmodus",
        ("Inflationsbereinigt (real)", "Nominal"),
        index=0 if st.session_state["use_inflation"] else 1,
    ) == "Inflationsbereinigt (real)"

    st.session_state["initial_wealth"] = st.sidebar.number_input(
        "Startvermögen (€)",
        min_value=0.0, max_value=100_000_000.0,
        value=float(st.session_state["initial_wealth"]),
        step=10_000.0, format="%.0f",
    )

    st.session_state["withdrawal_eur"] = st.sidebar.number_input(
        "Entnahme p.a. (€)",
        min_value=0.0,
        max_value=float(st.session_state["initial_wealth"]) if st.session_state["initial_wealth"] > 0 else 10_000_000.0,
        value=float(st.session_state["withdrawal_eur"]),
        step=1_000.0, format="%.0f",
        help="Jährlicher Entnahmebetrag in Euro. Wird automatisch ins Verhältnis zum Startvermögen gesetzt.",
    )

    _initial_wealth_for_rate = float(st.session_state["initial_wealth"])
    _withdrawal_eur_val      = float(st.session_state["withdrawal_eur"])
    _derived_rate_pct        = (
        (_withdrawal_eur_val / _initial_wealth_for_rate * 100.0)
        if _initial_wealth_for_rate > 0 else 0.0
    )
    st.sidebar.caption(f"→ Entspricht einem Entnahmesatz von **{_derived_rate_pct:.2f} % p.a.**")

    st.session_state["horizon_years"] = st.sidebar.number_input(
        "Horizont (Jahre)",
        min_value=5, max_value=60,
        value=int(st.session_state["horizon_years"]),
        step=1,
    )

    if st.session_state["erweitert"]:
        st.sidebar.divider()
        st.sidebar.subheader("Erfolgskurve (optional)")
        st.session_state["rate_min_pct"], st.session_state["rate_max_pct"] = st.sidebar.slider(
            "Spannweite (% p.a.)", 1.0, 10.0,
            (float(st.session_state["rate_min_pct"]), float(st.session_state["rate_max_pct"])),
            0.25, format="%.2f",
        )
        st.session_state["rate_step_pp"] = st.sidebar.number_input(
            "Schritt (Prozentpunkte)",
            min_value=0.05, max_value=2.0,
            value=float(st.session_state["rate_step_pp"]),
            step=0.05, format="%.2f",
        )
        st.session_state["show_gross_line"] = st.sidebar.checkbox(
            "Vergleich ohne Gebühren anzeigen",
            value=st.session_state["show_gross_line"],
        )

    # ---------------------------------------------------------------------------
    # BERECHNUNG BEI BUTTON-KLICK
    # ---------------------------------------------------------------------------

    current_fp = _current_config_fingerprint()

    if st.session_state["results"] is None:
        st.info('Bitte Einstellungen links wählen und dann „Berechnung starten" klicken.')
    else:
        if st.session_state.get("last_config_fingerprint") != current_fp:
            st.warning('Einstellungen wurden geändert. Bitte „Berechnung starten" klicken, um die Ergebnisse zu aktualisieren.')

    if run_clicked:
        sum_weights_pct_run = float(
            st.session_state["w_msci_pct"]
            + st.session_state["w_rexp_pct"]
            + st.session_state["w_gold_pct"]
        )

        if sum_weights_pct_run > 100.0 + 1e-9:
            st.error(
                f"Die Summe der Portfolio-Gewichte beträgt {sum_weights_pct_run:.1f} % und überschreitet 100 %.\n\n"
                "Bitte reduzieren Sie die Gewichte so, dass die Summe maximal 100 % ist."
            )
            st.stop()

        rest_cash_pct_run = float(max(0.0, 100.0 - sum_weights_pct_run))

        rate_min  = float(st.session_state["rate_min_pct"]  / 100.0)
        rate_max  = float(st.session_state["rate_max_pct"]  / 100.0)
        rate_step = float(st.session_state["rate_step_pp"]  / 100.0)

        if rate_min >= rate_max:
            st.error("Für die Erfolgskurve muss die minimale Rate kleiner als die maximale Rate sein.")
            st.stop()

        weights = {
            "msci_world": float(st.session_state["w_msci_pct"] / 100.0),
            "rexp":       float(st.session_state["w_rexp_pct"] / 100.0),
            "Gold":       float(st.session_state["w_gold_pct"] / 100.0),
        }
        if rest_cash_pct_run > 1e-9:
            weights["cash"] = float(rest_cash_pct_run / 100.0)

        annual_fee = float(st.session_state["annual_fee_pct"] / 100.0) if st.session_state["apply_fees"] else 0.0
        tax_rate   = float(st.session_state["tax_rate_pct"]   / 100.0) if st.session_state["apply_tax"]  else 0.0

        _initial_wealth_run = float(st.session_state["initial_wealth"])
        _withdrawal_eur_run = float(st.session_state["withdrawal_eur"])
        withdrawal_rate     = (_withdrawal_eur_run / _initial_wealth_run) if _initial_wealth_run > 0 else 0.0

        sim_cfg = SimulationConfig(
            annual_withdrawal_rate=withdrawal_rate,
            initial_wealth=_initial_wealth_run,
            periods_per_year=PERIODS_PER_YEAR,
            horizon_years=int(st.session_state["horizon_years"]),
        )

        port_cfg_net = PortfolioConfig(
            name="portfolio_net",
            weights=weights,
            annual_fee=annual_fee,
            use_fees=bool(st.session_state["apply_fees"]),
            use_inflation=bool(st.session_state["use_inflation"]),
        )

        port_cfg_gross = PortfolioConfig(
            name="portfolio_gross",
            weights=weights,
            annual_fee=0.0,
            use_fees=False,
            use_inflation=bool(st.session_state["use_inflation"]),
        )

        with st.spinner("Berechnung läuft..."):
            try:
                panel = _load_panel_from_source()
            except Exception as e:
                st.error(f"Fehler beim Laden der Marktdaten: {e}")
                st.stop()

            nominal_rets     = compute_nominal_returns(panel)
            inflation_series = nominal_rets["inflation"]

            portfolio_returns_net   = prepare_portfolio_returns(panel, port_cfg_net)
            portfolio_returns_gross = prepare_portfolio_returns(panel, port_cfg_gross)

            periods_needed = sim_cfg.horizon_years * sim_cfg.periods_per_year
            if len(portfolio_returns_net) < periods_needed:
                st.error(
                    f"Nicht genügend Daten für den gewählten Horizont. "
                    f"Benötigt: {periods_needed} Monate, verfügbar: {len(portfolio_returns_net)}."
                )
                st.stop()

            tax_str = f"mit Steuer {tax_rate*100:.2f} %" if tax_rate > 0 else "ohne Steuer"

            matrix        = build_wealth_matrix(portfolio_returns_net, sim_cfg, tax_rate=tax_rate)
            n_cohorts     = matrix.attrs.get("n_cohorts", matrix.shape[1])
            cohort_summary = summarise_cohorts(matrix)
            term_wealth    = compute_terminal_wealth_distribution(portfolio_returns_net, sim_cfg, tax_rate=tax_rate)

            terminals    = matrix.iloc[-1]
            best_start   = terminals.idxmax()
            worst_start  = terminals.idxmin()
            median_start = (terminals - terminals.median()).abs().idxmin()

            path_median = _simulate_path_for_start_date(portfolio_returns_net, sim_cfg, tax_rate, median_start)
            path_best   = _simulate_path_for_start_date(portfolio_returns_net, sim_cfg, tax_rate, best_start)
            path_worst  = _simulate_path_for_start_date(portfolio_returns_net, sim_cfg, tax_rate, worst_start)

            erfolg_aktuell = float(((matrix > 0).all(axis=0)).mean())

            st.session_state["results"] = {
                "panel":                  panel,
                "inflation_series":       inflation_series,
                "sim_cfg":                sim_cfg,
                "port_cfg_net":           port_cfg_net,
                "portfolio_returns_net":  portfolio_returns_net,
                "portfolio_returns_gross": portfolio_returns_gross,
                "tax_rate":               tax_rate,
                "tax_str":                tax_str,
                "matrix":                 matrix,
                "n_cohorts":              n_cohorts,
                "cohort_summary":         cohort_summary,
                "term_wealth":            term_wealth,
                "best_start":             best_start,
                "worst_start":            worst_start,
                "median_start":           median_start,
                "path_median":            path_median,
                "path_best":              path_best,
                "path_worst":             path_worst,
                "erfolg_aktuell":         erfolg_aktuell,
                "withdrawal_eur":         _withdrawal_eur_run,
                "rate_min":               rate_min,
                "rate_max":               rate_max,
                "rate_step":              rate_step,
                "show_gross_line":        bool(st.session_state.get("show_gross_line", False)),
            }

            st.session_state["last_config_fingerprint"] = current_fp

        st.rerun()

    # ---------------------------------------------------------------------------
    # OUTPUT
    # ---------------------------------------------------------------------------

    if st.session_state["results"] is not None:
        r = st.session_state["results"]

        st.subheader("Übersicht")

        col1, col2, col3, col4, col5 = st.columns(5)
        with col1:
            st.metric("Startvermögen",      _euro_str(r["sim_cfg"].initial_wealth))
        with col2:
            st.metric("Entnahme p.a.",      _euro_str(r["withdrawal_eur"]))
        with col3:
            st.metric("Entnahmesatz p.a.",  f"{r['sim_cfg'].annual_withdrawal_rate*100:.2f} %")
        with col4:
            st.metric("Horizont",           f"{r['sim_cfg'].horizon_years} Jahre")
        with col5:
            st.metric("Historische Läufe",  f"{r['n_cohorts']}")

        st.markdown(f"Historische Erfolgsquote für den gewählten Entnahmesatz: **{r['erfolg_aktuell']*100:.1f} %**")
        st.markdown(f"Portfolio-Mix: **{format_weights(r['port_cfg_net'].weights)}**")

        tabs = st.tabs(
            [
                "1. Historische Bandbreite",
                "2. Kumulierte Entnahmen",
                "3. Erfolgskurve",
                "4. Nießbrauchswert",
            ]
        )

        # -----------------------------------------------------------------------
        # TAB 1 – Fan-Chart
        # -----------------------------------------------------------------------
        with tabs[0]:
            st.header("Historische Bandbreite (Fan-Chart)")

            extra_fan  = r["tax_str"] + f", {r['n_cohorts']} Läufe"
            title_fan  = format_chart_title(
                "Historische Bandbreite", r["port_cfg_net"], r["sim_cfg"], extra=extra_fan
            )
            fig_fan = plot_fan_chart(r["matrix"], title=title_fan)
            st.pyplot(fig_fan, use_container_width=True)
            plt.close(fig_fan)

        # -----------------------------------------------------------------------
        # TAB 2 – Kumulierte Entnahmen
        # -----------------------------------------------------------------------
        with tabs[1]:
            st.header("Kumulierte Entnahmen")

            optionen = ("Median-Kohorte", "Beste Kohorte", "Schlechteste Kohorte")
            auswahl  = st.selectbox("Beispielkohorte auswählen", optionen, index=0)

            if auswahl == "Beste Kohorte":
                path       = r["path_best"]
                start_date = r["best_start"]
            elif auswahl == "Schlechteste Kohorte":
                path       = r["path_worst"]
                start_date = r["worst_start"]
            else:
                path       = r["path_median"]
                start_date = r["median_start"]

            st.caption(f"Beispiel-Startdatum: {_fmt_date_ym(pd.to_datetime(start_date))}")

            title_cum    = format_chart_title(
                "Kumulierte Entnahmen", r["port_cfg_net"], r["sim_cfg"], extra=r["tax_str"]
            )
            infl_for_plot = r["inflation_series"] if r["port_cfg_net"].use_inflation else None

            fig_cum = plot_cumulative_withdrawals(
                path, r["sim_cfg"], title_cum,
                inflation=infl_for_plot,
                real_mode=r["port_cfg_net"].use_inflation,
            )
            st.pyplot(fig_cum, use_container_width=True)
            plt.close(fig_cum)

        # -----------------------------------------------------------------------
        # TAB 3 – Erfolgskurve
        # -----------------------------------------------------------------------
        with tabs[2]:
            st.header("Erfolgskurve")

            with st.spinner("Erfolgskurve wird berechnet..."):
                fig_success = plot_success_curve(
                    r["portfolio_returns_gross"],
                    r["portfolio_returns_net"],
                    r["sim_cfg"],
                    r["port_cfg_net"],
                    rate_min=r["rate_min"],
                    rate_max=r["rate_max"],
                    rate_step=r["rate_step"],
                    tax_rate=r["tax_rate"],
                    show_gross_line=r["show_gross_line"],
                )
            st.pyplot(fig_success, use_container_width=True)
            plt.close(fig_success)

        # -----------------------------------------------------------------------
        # TAB 4 – Nießbrauchswert
        # -----------------------------------------------------------------------
        with tabs[3]:
            st.header("Nießbrauchswert des Portfolios (§ 14 BewG)")
            st.markdown(
                "Der **Nießbrauchswert** ist der steuerlich maßgebliche Kapitalwert eines "
                "Nießbrauchsrechts am Portfolio. Er ergibt sich aus dem jährlichen Ertrag "
                "(**Jahreswert**) multipliziert mit dem **Vervielfältiger** nach § 14 BewG, "
                "der von der statistischen Restlebenserwartung des Nießbrauchers abhängt. "
                "Der Jahreswert ist gemäß § 16 BewG auf 1/18,6 des Vermögens gedeckelt."
            )

            st.divider()

            nb_col_left, nb_col_right = st.columns([1, 1], gap="large")

            with nb_col_left:
                st.subheader("Eingaben")

                nb_restleben = st.number_input(
                    "Statistische Restlebenserwartung (Jahre)",
                    min_value=1.0,
                    max_value=80.0,
                    value=float(st.session_state["nb_restleben"]),
                    step=0.5,
                    format="%.1f",
                    key="nb_restleben_input",
                    help=(
                        "Statistische Restlebenserwartung laut amtlicher Sterbetafel "
                        "(z. B. Statistisches Bundesamt). Maßgeblich für den "
                        "Vervielfältiger nach § 14 BewG."
                    ),
                )
                st.session_state["nb_restleben"] = nb_restleben

                st.markdown("**Laufende Ertragsrenditen des Portfolios**")

                nb_dividende_pct = st.number_input(
                    "Ø Dividendenrendite – Aktienanteil (% p.a.)",
                    min_value=0.0,
                    max_value=20.0,
                    value=float(st.session_state["nb_dividende_pct"]),
                    step=0.1,
                    format="%.2f",
                    key="nb_dividende_input",
                    help=(
                        "Durchschnittliche laufende Dividendenrendite des Aktienanteils "
                        "(z. B. MSCI World Dividendenrendite ca. 1,2–1,5 % p.a.)."
                    ),
                )
                st.session_state["nb_dividende_pct"] = nb_dividende_pct

                nb_kupon_pct = st.number_input(
                    "Ø Kuponrendite – Rentenanteil (% p.a.)",
                    min_value=0.0,
                    max_value=20.0,
                    value=float(st.session_state["nb_kupon_pct"]),
                    step=0.1,
                    format="%.2f",
                    key="nb_kupon_input",
                    help=(
                        "Durchschnittliche laufende Kuponrendite des Rentenanteils "
                        "(z. B. REXP-Kupon historisch ca. 3–4 % p.a.)."
                    ),
                )
                st.session_state["nb_kupon_pct"] = nb_kupon_pct

                nb_gold_pct = st.number_input(
                    "Ø Rendite – Gold / Sonstige (% p.a.)",
                    min_value=0.0,
                    max_value=20.0,
                    value=float(st.session_state["nb_gold_pct"]),
                    step=0.1,
                    format="%.2f",
                    key="nb_gold_input",
                    help=(
                        "Laufende Rendite des Gold- bzw. Liquiditätsanteils. "
                        "Gold schüttet üblicherweise 0 % aus, Liquidität ebenfalls 0 %."
                    ),
                )
                st.session_state["nb_gold_pct"] = nb_gold_pct

                # Aktuelle Portfolio-Gewichte anzeigen
                st.markdown("**Verwendete Portfolio-Gewichte** *(aus der Berechnung übernommen)*")
                w = r["port_cfg_net"].weights
                gewichte_df = pd.DataFrame(
                    {
                        "Asset":   ["MSCI World", "REXP", "Gold", "Liquidität"],
                        "Gewicht": [
                            f"{w.get('msci_world', 0)*100:.1f} %",
                            f"{w.get('rexp',       0)*100:.1f} %",
                            f"{w.get('Gold',       0)*100:.1f} %",
                            f"{w.get('cash',       0)*100:.1f} %",
                        ],
                        "Rendite": [
                            f"{nb_dividende_pct:.2f} %",
                            f"{nb_kupon_pct:.2f} %",
                            f"{nb_gold_pct:.2f} %",
                            "0,00 %",
                        ],
                    }
                )
                st.dataframe(gewichte_df, hide_index=True, use_container_width=True)

            # Live-Berechnung (kein Button nötig – reine Formelauswertung)
            nb_result = compute_niessbrauch(
                initial_wealth=r["sim_cfg"].initial_wealth,
                weights=r["port_cfg_net"].weights,
                dividendenrendite=nb_dividende_pct / 100.0,
                kuponrendite=nb_kupon_pct      / 100.0,
                goldrendite=nb_gold_pct        / 100.0,
                restlebenserwartung=nb_restleben,
            )

            with nb_col_right:
                st.subheader("Ergebnis")

                # Farbige Hervorhebung des Nießbrauchswerts
                niessbrauch_formatted = _euro_str(nb_result["niessbrauchswert"])
                st.markdown(
                    f"""
                    <div style="
                        background-color: #003c71;
                        color: white;
                        border-radius: 8px;
                        padding: 1.2rem 1.5rem;
                        margin-bottom: 1rem;
                        text-align: center;
                    ">
                        <div style="font-size: 0.9rem; opacity: 0.85; margin-bottom: 0.3rem;">
                            Nießbrauchswert (§ 14 BewG)
                        </div>
                        <div style="font-size: 2rem; font-weight: bold; letter-spacing: 0.02em;">
                            {niessbrauch_formatted}
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                r1c1, r1c2 = st.columns(2)
                with r1c1:
                    st.metric(
                        "Gewichtete Portfoliorendite",
                        f"{nb_result['rendite_gewichtet']*100:.4f} % p.a.",
                        help="Gewichteter Durchschnittsertrag (Dividenden + Kupons + Sonstige).",
                    )
                with r1c2:
                    st.metric(
                        "Vervielfältiger (§ 14 BewG)",
                        f"{nb_result['vervielfaeltiger']:.1f}",
                        help=(
                            f"Aus der BMF-Tabelle interpoliert für "
                            f"{nb_restleben:.1f} Jahre Restlebenserwartung."
                        ),
                    )

                r2c1, r2c2 = st.columns(2)
                with r2c1:
                    st.metric(
                        "Jahreswert (vor Deckelung)",
                        _euro_str(nb_result["jahreswert_roh"]),
                        help="Gewichtete Rendite × Startvermögen.",
                    )
                with r2c2:
                    if nb_result["cap_aktiv"]:
                        st.metric(
                            "Jahreswert (§ 16 BewG, gedeckelt ✓)",
                            _euro_str(nb_result["jahreswert"]),
                            delta=f"Deckel: {_euro_str(nb_result['jahreswert_cap'])}",
                            delta_color="off",
                            help="Deckelung auf Startvermögen ÷ 18,6 ist aktiv.",
                        )
                    else:
                        st.metric(
                            "Jahreswert (maßgeblich)",
                            _euro_str(nb_result["jahreswert"]),
                            help="Deckelung nach § 16 BewG greift nicht.",
                        )

                st.metric(
                    "Startvermögen (Referenz)",
                    _euro_str(r["sim_cfg"].initial_wealth),
                )

            # Rechenweg-Tabelle
            st.divider()
            st.subheader("Rechenweg (transparent)")

            w = r["port_cfg_net"].weights
            rendite_zeile = " + ".join(
                [
                    f"{w.get('msci_world', 0)*100:.0f} % × {nb_dividende_pct:.2f} %",
                    f"{w.get('rexp', 0)*100:.0f} % × {nb_kupon_pct:.2f} %",
                    f"{w.get('Gold', 0)*100:.0f} % × {nb_gold_pct:.2f} %",
                ]
            )

            cap_hinweis = "*(aktiv – Jahreswert wurde gedeckelt)*" if nb_result["cap_aktiv"] else "*(nicht aktiv)*"

            st.markdown(
                f"""
| Schritt | Formel / Basis | Ergebnis |
|:---|:---|---:|
| Gewichtete Portfoliorendite | {rendite_zeile} | **{nb_result['rendite_gewichtet']*100:.4f} % p.a.** |
| Jahreswert (vor Deckelung) | {nb_result['rendite_gewichtet']*100:.4f} % × {_euro_str(r['sim_cfg'].initial_wealth)} | **{_euro_str(nb_result['jahreswert_roh'])}** |
| Deckelungsgrenze § 16 BewG | {_euro_str(r['sim_cfg'].initial_wealth)} ÷ 18,6 | **{_euro_str(nb_result['jahreswert_cap'])}** {cap_hinweis} |
| Maßgeblicher Jahreswert | | **{_euro_str(nb_result['jahreswert'])}** |
| Vervielfältiger § 14 BewG | Restlebenserwartung {nb_restleben:.1f} Jahre | **{nb_result['vervielfaeltiger']:.1f}** |
| **Nießbrauchswert** | {_euro_str(nb_result['jahreswert'])} × {nb_result['vervielfaeltiger']:.1f} | **{niessbrauch_formatted}** |
"""
            )

            st.caption(
                "⚠️ **Hinweis:** Diese Berechnung dient ausschließlich der Orientierung im Beratungsgespräch. "
                "Für die steuerlich verbindliche Bewertung ist ein Steuerberater hinzuzuziehen. "
                "Vervielfältiger gemäß BMF-Schreiben zu § 14 BewG, Jahreswertdeckelung gemäß § 16 BewG (Fassung 2024)."
            )

            # ---------------------------------------------------------------
            # DEPOTBASIERTE BERECHNUNG (optional)
            # ---------------------------------------------------------------
            st.divider()
            st.subheader("Depotbasierte Berechnung (optional)")
            st.markdown(
                "Statt pauschaler Renditeannahmen können die **konkreten Ausschüttungen** "
                "des Depots verwendet werden. Der Jahreswert ergibt sich dann als Summe der "
                "Dividenden und Kupons je Position, die Deckelungsgrenze nach § 16 BewG aus "
                "der Summe der Kurswerte in der Datei. Die Datei darf **keine personenbezogenen "
                "Daten** enthalten (Depot- oder Kundennummer, Name des Inhabers, IBAN, Adresse); "
                "solche Dateien werden abgelehnt."
            )

            dl_col, ul_col = st.columns([1, 2], gap="large")
            with dl_col:
                st.download_button(
                    "Muster-CSV herunterladen",
                    data=DEPOT_MUSTER_CSV.encode("utf-8-sig"),
                    file_name="depot_muster.csv",
                    mime="text/csv",
                )
            with ul_col:
                depot_datei = st.file_uploader(
                    "Depot-CSV hochladen (Semikolon, UTF-8)",
                    type=["csv"],
                    key="nb_depot_upload",
                )

            nb_wiederanlage_pct = st.number_input(
                "Wiederanlagerendite nach Fälligkeit von Anleihen (% p.a.)",
                min_value=0.0,
                max_value=20.0,
                value=float(st.session_state["nb_wiederanlage_pct"]),
                step=0.1,
                format="%.2f",
                key="nb_wiederanlage_input",
                help=(
                    "Nach Fälligkeit einer Anleihe wird der Rückzahlungsbetrag (Nominal zu 100 %) "
                    "zu diesem Satz wieder angelegt. Der Jahreswert ist der zeitgewichtete "
                    "Durchschnitt aus Kuponphase und Wiederanlagephase."
                ),
            )
            st.session_state["nb_wiederanlage_pct"] = nb_wiederanlage_pct

            depot_df = None
            if depot_datei is not None:
                try:
                    depot_df = load_depot_csv(depot_datei)
                    st.session_state["nb_depot_fehler"] = None
                except PermissionError as e:
                    st.session_state["nb_depot_fehler"] = str(e)
                except KeyError as e:
                    st.session_state["nb_depot_fehler"] = f"Spaltenfehler: {e}"
                except Exception as e:
                    st.session_state["nb_depot_fehler"] = f"Datei konnte nicht gelesen werden: {e}"
            else:
                st.session_state["nb_depot_fehler"] = None

            if st.session_state["nb_depot_fehler"]:
                st.error("🚫 " + str(st.session_state["nb_depot_fehler"]))

            if depot_df is not None:
                if len(depot_df) == 0:
                    st.warning("Die Datei enthält keine auswertbaren Positionen.")
                else:
                    nb_depot = compute_niessbrauch_aus_depot(
                        depot_df,
                        restlebenserwartung=nb_restleben,
                        wiederanlagerendite=nb_wiederanlage_pct / 100.0,
                    )

                    st.markdown(
                        f"""
                        <div style="
                            background-color: #003c71;
                            color: white;
                            border-radius: 8px;
                            padding: 1.2rem 1.5rem;
                            margin: 1rem 0;
                            text-align: center;
                        ">
                            <div style="font-size: 0.9rem; opacity: 0.85; margin-bottom: 0.3rem;">
                                Nießbrauchswert auf Depotbasis (§ 14 BewG)
                            </div>
                            <div style="font-size: 2rem; font-weight: bold; letter-spacing: 0.02em;">
                                {_euro_str(nb_depot['niessbrauchswert'])}
                            </div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

                    d1, d2, d3, d4 = st.columns(4)
                    with d1:
                        st.metric(
                            "Depotwert (Summe Kurswerte)",
                            _euro_str(nb_depot["vermoegen"]),
                            help="Basis der Deckelung nach § 16 BewG.",
                        )
                    with d2:
                        st.metric(
                            "Ausschüttungen heute p.a.",
                            _euro_str(nb_depot["ertrag_heute"]),
                            help="Summe der aktuellen Dividenden und Kupons.",
                        )
                    with d3:
                        st.metric(
                            "Jahreswert (nachhaltig)",
                            _euro_str(nb_depot["jahreswert_roh"]),
                            help="Nach Berücksichtigung der Wiederanlage fälliger Anleihen.",
                        )
                    with d4:
                        st.metric(
                            "Vervielfältiger (§ 14 BewG)",
                            f"{nb_depot['vervielfaeltiger']:.1f}",
                            help=f"Für {nb_restleben:.1f} Jahre Restlebenserwartung.",
                        )

                    d5, d6, d7 = st.columns(3)
                    with d5:
                        st.metric(
                            "Laufende Rendite des Depots",
                            f"{nb_depot['rendite_gewichtet']*100:.4f} % p.a.",
                            help="Nachhaltiger Jahreswert geteilt durch Depotwert.",
                        )
                    with d6:
                        if np.isfinite(nb_depot["restlaufzeit_renten"]):
                            st.metric(
                                "Ø Restlaufzeit Renten",
                                f"{nb_depot['restlaufzeit_renten']:.1f} Jahre",
                                help=(
                                    "Kurswertgewichtet. Je kürzer die Restlaufzeit, desto "
                                    "stärker wirkt die Wiederanlageannahme auf den Jahreswert."
                                ),
                            )
                        else:
                            st.metric("Ø Restlaufzeit Renten", "–")
                    with d7:
                        if nb_depot["cap_aktiv"]:
                            st.metric(
                                "Jahreswert (§ 16 BewG, gedeckelt ✓)",
                                _euro_str(nb_depot["jahreswert"]),
                                delta=f"Deckel: {_euro_str(nb_depot['jahreswert_cap'])}",
                                delta_color="off",
                                help="Deckelung auf Depotwert ÷ 18,6 ist aktiv.",
                            )
                        else:
                            st.metric(
                                "Jahreswert (maßgeblich)",
                                _euro_str(nb_depot["jahreswert"]),
                                help="Deckelung nach § 16 BewG greift nicht.",
                            )

                    if nb_depot["warnungen"]:
                        with st.expander(
                            f"⚠️ Hinweise zur Datei ({len(nb_depot['warnungen'])})",
                            expanded=True,
                        ):
                            for hinweis in nb_depot["warnungen"]:
                                st.markdown(f"- {hinweis}")

                    with st.expander("Positionen im Detail", expanded=False):
                        anzeige = nb_depot["daten"].copy()
                        spalten = [
                            c for c in [
                                "name", "assetklasse", "kurswert_eur",
                                "ertrag_pa_eur_calc", "restlaufzeit_jahre",
                                "ertrag_wiederanlage", "jahreswert_effektiv",
                            ] if c in anzeige.columns
                        ]
                        anzeige = anzeige[spalten].rename(columns={
                            "name":                "Position",
                            "assetklasse":         "Assetklasse",
                            "kurswert_eur":        "Kurswert (€)",
                            "ertrag_pa_eur_calc":  "Ertrag heute p.a. (€)",
                            "restlaufzeit_jahre":  "Restlaufzeit (J.)",
                            "ertrag_wiederanlage": "Ertrag n. Wiederanlage (€)",
                            "jahreswert_effektiv": "Jahreswert (€)",
                        })
                        st.dataframe(anzeige, hide_index=True, use_container_width=True)

                        st.download_button(
                            "Positionsübersicht als CSV (für den Steuerberater)",
                            data=anzeige.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig"),
                            file_name="niessbrauch_positionen.csv",
                            mime="text/csv",
                        )

                    st.markdown(
                        f"""
| Schritt | Formel / Basis | Ergebnis |
|:---|:---|---:|
| Ausschüttungen heute | Summe Dividenden und Kupons laut Datei | **{_euro_str(nb_depot['ertrag_heute'])}** |
| Nachhaltiger Jahreswert | Kupon bis Fälligkeit, danach Wiederanlage zu {nb_wiederanlage_pct:.2f} % | **{_euro_str(nb_depot['jahreswert_roh'])}** |
| Deckelungsgrenze § 16 BewG | {_euro_str(nb_depot['vermoegen'])} ÷ 18,6 | **{_euro_str(nb_depot['jahreswert_cap'])}** {"*(aktiv)*" if nb_depot["cap_aktiv"] else "*(nicht aktiv)*"} |
| Maßgeblicher Jahreswert | | **{_euro_str(nb_depot['jahreswert'])}** |
| Vervielfältiger § 14 BewG | Restlebenserwartung {nb_restleben:.1f} Jahre | **{nb_depot['vervielfaeltiger']:.1f}** |
| **Nießbrauchswert (Depotbasis)** | {_euro_str(nb_depot['jahreswert'])} × {nb_depot['vervielfaeltiger']:.1f} | **{_euro_str(nb_depot['niessbrauchswert'])}** |
"""
                    )

                    st.caption(
                        f"Stichtag der Bewertung: {nb_depot['stichtag'].strftime('%d.%m.%Y')}. "
                        "Der Jahreswert ist ein **zeitgewichteter Durchschnitt** aus laufendem Kupon "
                        "und unterstellter Wiederanlage; die Gewichtung ist undiskontiert und damit "
                        "eine Näherung, da der Vervielfältiger ein Barwertfaktor ist. Eine Aufteilung "
                        "in einen befristeten und einen aufgeschobenen Nießbrauch mit getrennten "
                        "Vervielfältigern wäre exakter und ist mit dem Steuerberater abzustimmen. "
                        "**Laufende Kosten** (Vermögensverwaltungs- und Depotgebühren) sind hier "
                        "**nicht** vom Jahreswert abgezogen; ob der Nießbraucher sie zu tragen hat, "
                        "richtet sich nach der Nießbrauchsvereinbarung. Thesaurierende Positionen "
                        "werden mit 0 € angesetzt. Die Wiederanlagerendite ist eine Annahme und "
                        "keine Prognose."
                    )


if __name__ == "__main__":
    run_app()
