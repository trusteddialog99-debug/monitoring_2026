import datetime
from typing import Optional
import io

import numpy as np
import pandas as pd
import streamlit as st
import plotly.graph_objects as go


def safe_pct_change(current: float, baseline: float) -> Optional[float]:
    """Calculate a safe percent change and avoid division by zero."""
    if pd.isna(current) or pd.isna(baseline):
        return np.nan
    if baseline == 0:
        if current == 0:
            return 0.0
        return np.sign(current) * np.inf
    return (current - baseline) / baseline * 100


def safe_gap_percent(total: float, signed: float) -> Optional[float]:
    """Calculate percent gap between Gesamt and signed volume 0."""
    if pd.isna(total) or pd.isna(signed) or total == 0:
        return np.nan
    return (total - signed) / total * 100


def get_previous_iso_week(year: int, week: int) -> tuple[int, int]:
    """Return the previous ISO week for a given year and week."""
    if week > 1:
        return year, week - 1

    previous_year = year - 1
    previous_week = datetime.date(previous_year, 12, 28).isocalendar()[1]
    return previous_year, previous_week


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(col).strip() for col in df.columns]
    return df


def validate_required_columns(df: pd.DataFrame, required: list[str]) -> None:
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise ValueError(f"Folgende Spalten fehlen in der Datei: {', '.join(missing)}")


def normalize_year_column(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize alternative year column names to a unified column name."""
    if "Jahr von Date" in df.columns and "Jahr" not in df.columns:
        df = df.rename(columns={"Jahr von Date": "Jahr"})
    return df


def map_column_aliases(df: pd.DataFrame) -> pd.DataFrame:
    """Map common alternative column names to expected canonical names."""
    import re

    def norm(s: str) -> str:
        return re.sub(r"[^a-z0-9]", "", str(s).lower())

    cols = list(df.columns)
    norm_map = {norm(c): c for c in cols}
    col_map = {}

    # Jahr aliases
    for a in ["jahrvondate", "jahr", "year"]:
        if a in norm_map and "Jahr" not in df.columns:
            col_map[norm_map[a]] = "Jahr"
            break

    # KW aliases
    for a in ["kw", "kalenderwoche", "kwvondate"]:
        if a in norm_map and "KW" not in df.columns:
            col_map[norm_map[a]] = "KW"
            break

    # Customer aliases
    for a in ["tdmcustomerohnesc", "tdmcustomer", "customer", "kunde", "tdmcustomerohnesc"]:
        if a in norm_map and "tDM Customer ohne SC" not in df.columns:
            col_map[norm_map[a]] = "tDM Customer ohne SC"
            break

    # Gesamt aliases
    for a in ["gesamt", "total", "summe", "alle", "gesamtsumme"]:
        if a in norm_map and "Gesamt" not in df.columns:
            col_map[norm_map[a]] = "Gesamt"
            break

    # Column '0'
    if "0" not in df.columns:
        for c in cols:
            if norm(c) == "0":
                col_map[c] = "0"
                break

    if col_map:
        df = df.rename(columns=col_map)

    return df


def read_input_data(uploaded_file) -> pd.DataFrame:
    """Read uploaded Excel or CSV data into a DataFrame."""
    filename = getattr(uploaded_file, "name", "").lower()

    if filename.endswith(".csv"):
        raw = uploaded_file.read()
        if isinstance(raw, str):
            raw = raw.encode("utf-8")

        encodings = ["utf-8-sig", "utf-8", "utf-16", "utf-16-le", "utf-16-be", "cp1252", "latin1"]
        separators = [(";", ","), (",", "."), ("\t", ",")]
        last_exc = None

        for encoding in encodings:
            for sep, decimal in separators:
                try:
                    read_kwargs = dict(
                        sep=sep,
                        decimal=decimal,
                        engine="python",
                        encoding=encoding,
                    )
                    if decimal == ",":
                        read_kwargs["thousands"] = "."
                    df = pd.read_csv(io.BytesIO(raw), **read_kwargs)
                    if df.shape[1] > 1:
                        return df
                except Exception as exc:
                    last_exc = exc

        # Fallback
        for enc in encodings:
            try:
                decoded = raw.decode(enc, errors="replace")
            except Exception:
                decoded = None

            if decoded is None:
                continue

            for sep, decimal in separators:
                try:
                    read_kwargs = dict(sep=sep, decimal=decimal, engine="python")
                    if decimal == ",":
                        read_kwargs["thousands"] = "."
                    df = pd.read_csv(io.StringIO(decoded), **read_kwargs)
                    if df.shape[1] > 1:
                        return df
                except Exception:
                    continue

            try:
                df = pd.read_csv(io.StringIO(decoded), engine="python")
                return df
            except Exception:
                continue

        raise last_exc if last_exc is not None else ValueError("Fehler beim Einlesen der CSV-Datei.")

    uploaded_file.seek(0)
    return pd.read_excel(uploaded_file, engine="openpyxl")


def parse_int_series(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").astype("Int64")


def parse_float_series(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").astype(float)


def fmt_thousands_point(x) -> str:
    """Format number with German thousands separator (.) and no decimals."""
    if pd.isna(x):
        return "N/A"
    try:
        return f"{x:,.0f}".replace(",", ".")
    except Exception:
        return str(x)


def fmt_percent_no_decimal(x) -> str:
    """Format percent without decimal places and with comma as decimal separator if needed."""
    if pd.isna(x):
        return "N/A"
    try:
        s = f"{x:.0f}%"
        return s.replace(".", ",")
    except Exception:
        return str(x)


def build_monitoring_table(
    df: pd.DataFrame,
    threshold: float,
    min_volume: float,
    use_threshold: bool = True
) -> pd.DataFrame:
    df = normalize_columns(df)
    df = normalize_year_column(df)
    df = map_column_aliases(df)

    required_columns = ["Jahr", "KW", "tDM Customer ohne SC", "0", "Gesamt"]
    validate_required_columns(df, required_columns)

    df["Jahr"] = parse_int_series(df["Jahr"]).astype(int)
    df["KW"] = parse_int_series(df["KW"]).astype(int)
    df["tDM Customer ohne SC"] = df["tDM Customer ohne SC"].astype(str).str.strip()
    df["volume_0"] = parse_float_series(df["0"]).fillna(0.0)
    df["gesamt"] = parse_float_series(df["Gesamt"]).fillna(0.0)

    df = df[["Jahr", "KW", "tDM Customer ohne SC", "volume_0", "gesamt"]].copy()

    current_year = int(df.loc[df["Jahr"].idxmax(), "Jahr"])
    latest_years = df[df["Jahr"] == current_year]
    if not latest_years.empty:
        current_kw = int(latest_years["KW"].max())
    else:
        raise ValueError("Konnte die aktuelle Kalenderwoche nicht bestimmen.")

    prev_year_1, prev_kw_1 = get_previous_iso_week(current_year, current_kw)
    prev_year_2, prev_kw_2 = get_previous_iso_week(prev_year_1, prev_kw_1)

    results = []
    grouped = df.groupby("tDM Customer ohne SC", sort=True)

    for customer, group in grouped:
        group = group.sort_values(["Jahr", "KW"]).reset_index(drop=True)
        current_mask = (group["Jahr"] == current_year) & (group["KW"] == current_kw)
        if not current_mask.any():
            continue

        current_value = float(group.loc[current_mask, "volume_0"].iloc[0])
        current_total = float(group.loc[current_mask, "gesamt"].iloc[0])

        prev_mask = (group["Jahr"] == prev_year_1) & (group["KW"] == prev_kw_1)
        prev_prev_mask = (group["Jahr"] == prev_year_2) & (group["KW"] == prev_kw_2)

        prev_value = float(group.loc[prev_mask, "volume_0"].iloc[0]) if prev_mask.any() else np.nan
        prev_total = float(group.loc[prev_mask, "gesamt"].iloc[0]) if prev_mask.any() else np.nan

        prev_prev_value = float(group.loc[prev_prev_mask, "volume_0"].iloc[0]) if prev_prev_mask.any() else np.nan
        prev_prev_total = float(group.loc[prev_prev_mask, "gesamt"].iloc[0]) if prev_prev_mask.any() else np.nan

        current_index = group.index[current_mask][0]
        prior_rows = group.loc[:current_index - 1, "volume_0"] if current_index > 0 else pd.Series(dtype=float)
        prior_totals = group.loc[:current_index - 1, "gesamt"] if current_index > 0 else pd.Series(dtype=float)

        avg_4 = prior_rows.tail(4).mean() if len(prior_rows) > 0 else np.nan
        avg_8 = prior_rows.tail(8).mean() if len(prior_rows) > 0 else np.nan
        avg_4_total = prior_totals.tail(4).mean() if len(prior_totals) > 0 else np.nan
        avg_8_total = prior_totals.tail(8).mean() if len(prior_totals) > 0 else np.nan

        gap_current = safe_gap_percent(current_total, current_value)
        gap_prev = safe_gap_percent(prev_total, prev_value)
        gap_prev_prev = safe_gap_percent(prev_prev_total, prev_prev_value)
        gap_avg4 = safe_gap_percent(avg_4_total, avg_4)
        gap_avg8 = safe_gap_percent(avg_8_total, avg_8)

        diff_prev = safe_pct_change(current_value, prev_value)
        diff_avg4 = safe_pct_change(current_value, avg_4)
        diff_avg8 = safe_pct_change(current_value, avg_8)

        comparison_value = avg_4 if not np.isnan(avg_4) else prev_value if not np.isnan(prev_value) else avg_8
        comparison_diff = safe_pct_change(current_value, comparison_value)

        results.append(
            {
                "tDM Customer ohne SC": customer,
                "Jahr": current_year,
                "KW": current_kw,
                "Aktuelle Woche (0)": current_value,
                "Aktuelle Woche (Gesamt)": current_total,
                "Gap Gesamt vs 0 (%)": gap_current,
                "Vorwoche": prev_value,
                "Vorwoche Gesamt": prev_total,
                "Gap Vorwoche (%)": gap_prev,
                "Vor-Vorwoche": prev_prev_value,
                "Vor-Vorwoche Gesamt": prev_prev_total,
                "Gap Vor-Vorwoche (%)": gap_prev_prev,
                "Durchschnitt 4 Wochen": avg_4,
                "Durchschnitt 4 Wochen Gesamt": avg_4_total,
                "Gap Ø 4 Wochen (%)": gap_avg4,
                "Durchschnitt 8 Wochen": avg_8,
                "Durchschnitt 8 Wochen Gesamt": avg_8_total,
                "Gap Ø 8 Wochen (%)": gap_avg8,
                "% Veränderung vs Vorwoche": diff_prev,
                "% Veränderung vs Ø 4 Wochen": diff_avg4,
                "% Veränderung vs Ø 8 Wochen": diff_avg8,
                "% Veränderung (Hauptvergleich)": comparison_diff,
            }
        )

    result_df = pd.DataFrame(results)
    if result_df.empty:
        return result_df

    if use_threshold:
        filtered = result_df[
            (
                (result_df["Aktuelle Woche (0)"] == 0)
                | (
                    (result_df["Aktuelle Woche (0)"] >= float(min_volume))
                    & (result_df["% Veränderung (Hauptvergleich)"] < float(threshold))
                )
            )
        ].copy()
    else:
        filtered = result_df[result_df["Aktuelle Woche (0)"] >= float(min_volume)].
