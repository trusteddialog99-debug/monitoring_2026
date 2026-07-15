import datetime
from typing import Optional
import io
import numpy as np
import pandas as pd
import streamlit as st
import plotly.graph_objects as go


def safe_pct_change(current: float, baseline: float) -> Optional[float]:
    if pd.isna(current) or pd.isna(baseline):
        return np.nan
    if baseline == 0:
        if current == 0:
            return 0.0
        return np.sign(current) * np.inf
    return (current - baseline) / baseline * 100


def safe_gap_percent(total: float, signed: float) -> Optional[float]:
    if pd.isna(total) or pd.isna(signed) or total == 0:
        return np.nan
    return (total - signed) / total * 100


def get_previous_iso_week(year: int, week: int) -> tuple[int, int]:
    if week > 1:
        return year, week - 1
    previous_year = year - 1
    previous_week = datetime.date(previous_year, 12, 28).isocalendar()[1]
    return previous_year, previous_week


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(col).strip() for col in df.columns]
    return df


def parse_int_series(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").astype("Int64")


def parse_float_series(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").astype(float)


def main():

    st.set_page_config(page_title="E-Mail Versandmonitoring", layout="wide")

    uploaded_file = st.file_uploader("Datei hochladen", type=["xlsx", "xls", "csv"])
    if uploaded_file is None:
        st.stop()

    data = pd.read_excel(uploaded_file)

    data = normalize_columns(data)

    data["Jahr"] = parse_int_series(data["Jahr"]).astype(int)
    data["KW"] = parse_int_series(data["KW"]).astype(int)
    data["tDM Customer ohne SC"] = data["tDM Customer ohne SC"].astype(str).str.strip()
    data["volume_0"] = parse_float_series(data["0"]).fillna(0.0)
    data["gesamt"] = parse_float_series(data["Gesamt"]).fillna(0.0)

    customers = sorted(data["tDM Customer ohne SC"].unique().tolist())

    # ✅ FIX START
    if "selected_customer" not in st.session_state:
        st.session_state.selected_customer = customers[0]

    search_term = st.text_input("Kunde suchen")

    filtered_customers = (
        [c for c in customers if search_term.lower() in c.lower()]
        if search_term else customers
    )

    if not filtered_customers:
        st.warning("Kein Kunde gefunden.")
        filtered_customers = customers

    if st.session_state.selected_customer not in filtered_customers:
        st.session_state.selected_customer = filtered_customers[0]

    selected = st.selectbox(
        "Kunde auswählen für Historie",
        filtered_customers,
        key="selected_customer"
    )
    # ✅ FIX END

    st.subheader(f"Historisches Mailvolumen für: {selected}")

    df_cust = data[data["tDM Customer ohne SC"] == selected]

    agg = df_cust.groupby(["Jahr", "KW"], as_index=False).agg(
        volume_0=("volume_0", "sum"),
        gesamt=("gesamt", "sum"),
    )

    years = sorted(agg["Jahr"].unique())

    full = pd.MultiIndex.from_product(
        [years, list(range(1, 54))],
        names=["Jahr", "KW"]
    ).to_frame(index=False)

    full = full.merge(agg, on=["Jahr", "KW"], how="left")

    fig = go.Figure()

    for year in years:
        year_data = full[full["Jahr"] == year]

        fig.add_trace(
            go.Scatter(
                x=year_data["KW"],
                y=year_data["volume_0"],
                mode="lines",
                name=f"{year}"
            )
        )

    fig.update_layout(
        xaxis_title="Kalenderwoche",
        yaxis_title="Mailvolumen",
        hovermode="x unified"
    )

    st.plotly_chart(fig, use_container_width=True)


if __name__ == "__main__":
    main()
