"""Streamlit tool for comparing annual bid results stored in Excel workbooks."""

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, IO, Optional, Union

import altair as alt
import pandas as pd
import streamlit as st

BASE_DIR = Path(__file__).parent
BIDS_SHEET_NAME = "Bids"
ITEMS_SHEET_NAME = "Items"
ITEM_COLUMN_ALIASES = ("Item No", "Item Number", "Item #", "Item")
NUMERIC_COLUMNS = ("Item No", "Quantity", "Price", "Total Cost", "Bid Rank")
WorkbookSource = Union[Path, IO[bytes]]

st.set_page_config(page_title="Bid Comparison", page_icon="📊", layout="wide")


def _resolve_sheet_name(xls: pd.ExcelFile, target: str) -> Optional[str]:
    """Return the actual sheet name that matches target (case-insensitive)."""
    target_lower = target.lower()
    for sheet in xls.sheet_names:
        if sheet.lower() == target_lower:
            return sheet
    return None


def _discover_workbooks(search_dir: Path) -> Dict[str, Path]:
    """Return a dict of year label -> workbook path discovered in search_dir."""
    workbooks: Dict[str, Path] = {}
    for path in sorted(search_dir.glob("*.xlsx")):
        year_label = path.stem.strip()
        if not year_label.isdigit():
            continue
        workbooks[year_label] = path
    return workbooks


def _discover_bid_type_dirs() -> Dict[str, Path]:
    """Return bid-type folders that contain at least one .xlsx workbook."""
    bid_dirs: Dict[str, Path] = {}
    for path in sorted(BASE_DIR.iterdir()):
        if not path.is_dir():
            continue
        if path.name.startswith("."):
            continue
        has_workbook = any(
            child.is_file() and child.suffix.lower() == ".xlsx"
            for child in path.iterdir()
        )
        if has_workbook:
            bid_dirs[path.name] = path
    return bid_dirs


def _ensure_item_column(df: pd.DataFrame) -> pd.DataFrame:
    """Rename known item-number aliases to 'Item No'."""
    if "Item No" in df.columns:
        return df

    normalized = {str(col).strip().lower(): col for col in df.columns}
    target_lower = "item no"
    if target_lower in normalized:
        source_col = normalized[target_lower]
        if source_col != "Item No":
            df = df.rename(columns={source_col: "Item No"})
        return df

    for alias in ITEM_COLUMN_ALIASES:
        alias_lower = alias.strip().lower()
        if alias_lower in normalized:
            df = df.rename(columns={normalized[alias_lower]: "Item No"})
            break
    return df


def _deduplicate_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Ensure column labels are unique by suffixing duplicates."""
    df = df.copy()
    counts: Dict[str, int] = {}
    new_columns = []
    for col in df.columns:
        label = str(col)
        count = counts.get(label, 0)
        if count:
            new_columns.append(f"{label}_{count}")
        else:
            new_columns.append(label)
        counts[label] = count + 1
    df.columns = new_columns
    return df


def _format_item_identifier(value: Any) -> str:
    """Return a friendly string for an item number."""
    if pd.isna(value):
        return ""
    if isinstance(value, (int,)):
        return str(value)
    if isinstance(value, float):
        return str(int(value)) if value.is_integer() else f"{value}"
    text = str(value).strip()
    if not text:
        return ""
    try:
        numeric = float(text)
    except ValueError:
        return text
    return str(int(numeric)) if numeric.is_integer() else text


def _add_item_labels(df: pd.DataFrame) -> pd.DataFrame:
    """Attach a descriptive Item Label column for display/filtering."""
    if "Item No" not in df.columns:
        return df

    df = df.copy()
    info_cols = ["Item No"]
    description_column = None
    if "Item Description" in df.columns:
        info_cols.append("Item Description")
        description_column = "Item Description"
    elif "Description" in df.columns:
        info_cols.append("Description")
        description_column = "Description"

    info = df[info_cols].dropna(subset=["Item No"]).drop_duplicates("Item No")
    fallback = info["Item No"].apply(_format_item_identifier).replace("", "Unknown")

    if description_column:
        labels = info[description_column].fillna("").astype(str).str.strip()
        labels = labels.where(labels != "", "Item " + fallback)
    else:
        labels = "Item " + fallback

    duplicates = labels.duplicated(keep=False)
    labels.loc[duplicates] = labels.loc[duplicates] + " (Item " + fallback.loc[duplicates] + ")"

    info["Item Label"] = labels
    label_map = dict(zip(info["Item No"], info["Item Label"]))

    df["Item Label"] = df["Item No"].map(label_map)
    fallback_series = df["Item No"].apply(_format_item_identifier)
    df["Item Label"] = df["Item Label"].fillna("Item " + fallback_series)
    return df


def _coerce_numeric(df: pd.DataFrame, columns: Union[tuple[str, ...], list[str]]) -> pd.DataFrame:
    """Convert specified columns to numeric values when possible."""
    df = df.copy()
    for col in columns:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    return df


def _prepare_bids(df: pd.DataFrame, year_label: str) -> pd.DataFrame:
    df = _ensure_item_column(df)
    df = _coerce_numeric(df, NUMERIC_COLUMNS)
    df["Year"] = int(year_label)
    return df


def _prepare_items(df: pd.DataFrame) -> pd.DataFrame:
    df = _ensure_item_column(df)
    df = _coerce_numeric(df, ("Item No",))
    rename_map = {col: f"Item {col}" for col in df.columns if col != "Item No"}
    df = df.rename(columns=rename_map)
    return df


def _load_year_dataset(
    source: WorkbookSource, label: str, year_label: str
) -> Optional[pd.DataFrame]:
    """Read the Bids and Items sheets for a workbook."""
    try:
        xls = pd.ExcelFile(source, engine="openpyxl")
    except PermissionError:
        st.error(
            f"Unable to open '{label}'. Close it in Excel if it's currently open and retry."
        )
        return None
    except Exception as exc:  # pragma: no cover - surfaced in UI
        st.error(f"Failed to load '{label}': {exc}")
        return None

    bids_sheet = _resolve_sheet_name(xls, BIDS_SHEET_NAME)
    if bids_sheet is None:
        st.error(f"Workbook '{label}' is missing the '{BIDS_SHEET_NAME}' sheet.")
        return None

    bids_df = pd.read_excel(xls, sheet_name=bids_sheet, engine="openpyxl")
    bids_df = _deduplicate_columns(bids_df)
    bids_df = _prepare_bids(bids_df, year_label)

    items_df: Optional[pd.DataFrame] = None
    items_sheet = _resolve_sheet_name(xls, ITEMS_SHEET_NAME)
    if items_sheet is not None:
        raw_items = pd.read_excel(xls, sheet_name=items_sheet, engine="openpyxl")
        raw_items = _deduplicate_columns(raw_items)
        items_df = _prepare_items(raw_items)
    else:
        st.warning(
            f"Workbook '{label}' does not have an '{ITEMS_SHEET_NAME}' sheet. "
            "Item descriptions will be blank."
        )

    if items_df is not None and "Item No" in items_df.columns:
        dataset = bids_df.merge(items_df, on="Item No", how="left")
    else:
        dataset = bids_df
    dataset = _add_item_labels(dataset)
    return dataset


def load_datasets() -> Dict[str, pd.DataFrame]:
    """Load all detected Excel workbooks within the working directory."""
    datasets: Dict[str, pd.DataFrame] = {}
    with st.sidebar:
        st.header("Data Sources")
        bid_dirs = _discover_bid_type_dirs()
        if bid_dirs:
            bid_type = st.selectbox("Bid type", options=sorted(bid_dirs.keys()))
            search_dir = bid_dirs[bid_type]
        else:
            st.info("No bid-type folders found. Looking for workbooks in the repo root.")
            search_dir = BASE_DIR

        workbooks = _discover_workbooks(search_dir)
        if not workbooks:
            st.error("No year-specific workbooks (e.g., 2025.xlsx) were found.")
            return {}

        for year, path in workbooks.items():

            st.markdown(f"**{year}** — {path.name}")
            if not path.exists():
                st.warning(f"Missing file: {path}")
                continue

            dataset = _load_year_dataset(path, label=str(path), year_label=year)
            if dataset is not None:
                datasets[year] = dataset
                st.caption(f"Loaded {len(dataset):,} bid rows.")
            else:
                st.error(f"Failed to load workbook for {year}.")

        st.divider()
    return datasets


def combine_datasets(datasets: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    if not datasets:
        return pd.DataFrame()
    frames = [df for _, df in sorted(datasets.items())]
    return pd.concat(frames, ignore_index=True)


def apply_filters(df: pd.DataFrame) -> tuple[pd.DataFrame, str]:
    if df.empty or "Item No" not in df.columns:
        return df, "All bids"

    with st.sidebar:
        st.header("Options")
        mode = st.radio(
            "Bid view",
            options=["All bids", "Lowest bidders"],
            index=0,
            help="Choose whether to analyze all bids or only the lowest bidder per item/year.",
        )
        st.subheader("Filters")
        year_values = sorted(df["Year"].dropna().unique().tolist())
        year_selection = st.multiselect("Year", year_values, default=year_values)

        item_numbers = df["Item No"].dropna().unique().tolist()
        item_info = (
            df[["Item No", "Item Label"]]
            .dropna(subset=["Item No"])
            .drop_duplicates("Item No")
            if "Item Label" in df.columns
            else None
        )
        if item_info is not None and not item_info.empty:
            item_info = item_info.sort_values("Item Label")
            item_options = item_info["Item Label"].tolist()
            selected_labels = st.multiselect(
                "Items (by description)", item_options, default=item_options
            )
            if selected_labels:
                selected_numbers = item_info[
                    item_info["Item Label"].isin(selected_labels)
                ]["Item No"].tolist()
            else:
                selected_numbers = item_info["Item No"].tolist()
        else:
            item_values = sorted(item_numbers)
            selected_numbers = st.multiselect(
                "Item No", item_values, default=item_values
            )

        org_values = (
            sorted(df["Organization Name"].dropna().unique().tolist())
            if "Organization Name" in df.columns
            else []
        )
        org_selection = st.multiselect("Bidders", org_values)

    filtered = df[df["Year"].isin(year_selection)]
    filtered = filtered[filtered["Item No"].isin(selected_numbers)]
    if org_selection:
        filtered = filtered[filtered["Organization Name"].isin(org_selection)]
    return filtered, mode

def _extract_lowest_bids(df: pd.DataFrame) -> pd.DataFrame:
    """Return lowest bid per item/year with original columns."""
    required = {"Item No", "Year", "Price"}
    if df.empty or not required.issubset(df.columns):
        return pd.DataFrame(columns=df.columns)
    subset = df.dropna(subset=["Price"]).copy()
    if subset.empty:
        return pd.DataFrame(columns=df.columns)
    sort_fields = ["Item No", "Year", "Price"]
    if "Bid Rank" in subset.columns:
        sort_fields.append("Bid Rank")
    subset = subset.sort_values(sort_fields)
    winners = (
        subset.groupby(["Item No", "Year"], as_index=False)
        .first()
        .copy()
    )
    return winners


def compute_lowest_bid_view(df: pd.DataFrame) -> pd.DataFrame:
    """Return the lowest bidders with the same columns as the filtered dataset."""
    winners = _extract_lowest_bids(df)
    return winners


def _weighted_average_all(df: pd.DataFrame) -> float:
    if df.empty or not {"Price", "Quantity"}.issubset(df.columns):
        return float("nan")
    subset = df.dropna(subset=["Price", "Quantity"])
    if subset.empty:
        return float("nan")
    total_quantity = subset["Quantity"].sum()
    if not total_quantity or pd.isna(total_quantity):
        return float("nan")
    return (subset["Price"] * subset["Quantity"]).sum() / total_quantity


def _weighted_average_lowest(df: pd.DataFrame) -> float:
    winners = _extract_lowest_bids(df)
    if winners.empty or "Quantity" not in winners.columns:
        return float("nan")
    subset = winners.dropna(subset=["Price", "Quantity"])
    if subset.empty:
        return float("nan")
    total_quantity = subset["Quantity"].sum()
    if not total_quantity or pd.isna(total_quantity):
        return float("nan")
    return (subset["Price"] * subset["Quantity"]).sum() / total_quantity


def show_metrics(df: pd.DataFrame, mode: str) -> None:
    if df.empty:
        st.info("No bids match the current filters.")
        return

    col_items, col_orgs, col_avg, col_weighted = st.columns(4)
    distinct_items = df["Item No"].nunique()
    col_items.metric("Distinct items", f"{distinct_items:,}")

    distinct_bidders = (
        df["Organization Name"].nunique() if "Organization Name" in df.columns else 0
    )
    col_orgs.metric("Bidders", f"{distinct_bidders:,}")

    avg_price = df["Price"].mean() if "Price" in df.columns else float("nan")
    avg_price_display = f"${avg_price:,.2f}" if pd.notna(avg_price) else "N/A"
    col_avg.metric("Average unit price", avg_price_display)

    if mode == "Lowest bidders":
        weighted_avg = _weighted_average_lowest(df)
        label = "Weighted avg (lowest)"
    else:
        weighted_avg = _weighted_average_all(df)
        label = "Weighted avg (all)"
    weighted_display = f"${weighted_avg:,.2f}" if pd.notna(weighted_avg) else "N/A"
    col_weighted.metric(label, weighted_display)


def prepare_price_history(df: pd.DataFrame, mode: str) -> pd.DataFrame:
    if df.empty or "Price" not in df.columns or "Year" not in df.columns:
        return pd.DataFrame()

    base = df.dropna(subset=["Price", "Year"]).copy()
    if base.empty:
        return pd.DataFrame()

    base["Year"] = pd.to_numeric(base["Year"], errors="coerce")
    base = base.dropna(subset=["Year"])
    if base.empty:
        return pd.DataFrame()
    base["Year"] = base["Year"].astype(int)

    if mode == "Lowest bidders":
        if "Item Label" not in base.columns:
            base["Item Label"] = "Item " + base["Item No"].astype(str)
        result = base[["Item No", "Item Label", "Year", "Price"]].copy()
        return result

    summary = base.groupby("Year", as_index=False)["Price"].mean()
    summary["Price"] = summary["Price"].round(2)
    summary["Item Label"] = "Average of all bids"
    return summary


def compute_lowest_bid_table(df: pd.DataFrame) -> pd.DataFrame:
    """Return a per-item view of the lowest bidder per year."""
    required_cols = {"Item No", "Year", "Price", "Organization Name"}
    if df.empty or not required_cols.issubset(df.columns):
        return pd.DataFrame()

    sortable = df.dropna(subset=["Price"]).copy()
    if sortable.empty:
        return pd.DataFrame()

    sort_fields = ["Item No", "Year", "Price"]
    if "Bid Rank" in sortable.columns:
        sort_fields.append("Bid Rank")
    sortable = sortable.sort_values(sort_fields)

    winners = (
        sortable.groupby(["Item No", "Year"], as_index=False)
        .first()
        .copy()
    )
    if winners.empty:
        return pd.DataFrame()

    metadata_cols = [col for col in df.columns if col.startswith("Item ")]
    metadata = (
        df[["Item No"] + metadata_cols]
        .drop_duplicates("Item No")
        .pipe(_deduplicate_columns)
        if metadata_cols
        else None
    )

    pivot_cols = [
        col
        for col in ["Organization Name", "Price", "Total Cost", "Bid Rank"]
        if col in winners.columns
    ]
    pivot = winners.pivot(index="Item No", columns="Year", values=pivot_cols)
    pivot.columns = [f"{col}_{int(year)}" for col, year in pivot.columns]
    result = pivot.reset_index()

    rename_map: Dict[str, str] = {}
    for col in list(result.columns):
        if "_" not in col:
            continue
        metric, year_token = col.rsplit("_", 1)
        if not year_token.isdigit():
            continue
        metric_label = {
            "Organization Name": "Winner",
            "Price": "Price",
            "Total Cost": "Total Cost",
            "Bid Rank": "Bid Rank",
        }.get(metric, metric)
        rename_map[col] = f"{metric_label} {year_token}"

    result = result.rename(columns=rename_map)

    if metadata is not None and not metadata.empty:
        result = result.merge(metadata, on="Item No", how="left")

    price_columns = sorted(
        [
            (int(col.split(" ", 1)[-1]), col)
            for col in result.columns
            if col.startswith("Price ") and col.split(" ", 1)[-1].isdigit()
        ]
    )
    if len(price_columns) >= 2:
        earliest_year, earliest_col = price_columns[0]
        latest_year, latest_col = price_columns[-1]
        result["Price Δ"] = result[latest_col] - result[earliest_col]
        result["Price Δ%"] = (
            (result[latest_col] - result[earliest_col]) / result[earliest_col]
        ) * 100

    return result


def main() -> None:
    st.title("Bid Comparison From Excel")
    st.caption("Compare item-level bids across years using the cleaned Excel workbooks.")

    datasets = load_datasets()
    if not datasets:
        st.warning("Load at least one workbook to begin.")
        return

    combined = combine_datasets(datasets)
    if combined.empty:
        st.warning("No bid data found in the selected workbooks.")
        return

    filtered, mode = apply_filters(combined)
    if mode == "Lowest bidders":
        display_df = compute_lowest_bid_view(filtered)
    else:
        display_df = filtered

    show_metrics(display_df, mode)

    st.subheader("Bid detail")
    st.dataframe(display_df, use_container_width=True)

    lowest = compute_lowest_bid_table(filtered)
    st.subheader("Lowest bidder by item")
    if lowest.empty:
        st.info("Select both years with price data to compare lowest bidders.")
    else:
        st.dataframe(lowest, use_container_width=True)
        st.download_button(
            "Download lowest-bid comparison",
            data=lowest.to_csv(index=False),
            file_name="lowest_bid_comparison.csv",
            mime="text/csv",
        )

    st.subheader("Price history")
    history = prepare_price_history(display_df, mode)
    if history.empty:
        st.info("No price history available for the selected scope.")
    else:
        if mode == "Lowest bidders":
            chart_item_info = (
                history[["Item No", "Item Label"]]
                .dropna(subset=["Item No"])
                .drop_duplicates(["Item No", "Item Label"])
                .sort_values("Item Label")
            )
            label_options = chart_item_info["Item Label"].tolist()
            default_labels = label_options[: min(5, len(label_options))]
            selected_labels = st.multiselect(
                "Items to chart (by description)",
                options=label_options,
                default=default_labels,
                help="Select which items to visualize.",
            )
            if not selected_labels:
                st.info("Select at least one item to display the chart.")
            else:
                selected_items = chart_item_info[
                    chart_item_info["Item Label"].isin(selected_labels)
                ]["Item No"].tolist()
                chart_data = history[history["Item No"].isin(selected_items)]
                chart = (
                    alt.Chart(chart_data)
                    .mark_line(point=True)
                    .encode(
                        x=alt.X("Year:O", title="Year"),
                        y=alt.Y("Price:Q", title="Unit price"),
                        color=alt.Color("Item Label:N", title="Item"),
                        tooltip=[
                            alt.Tooltip("Item Label:N", title="Item"),
                            alt.Tooltip("Year:O", title="Year"),
                            alt.Tooltip("Price:Q", title="Price", format="$.2f"),
                        ],
                    )
                    .properties(height=400)
                )
                st.altair_chart(chart, use_container_width=True)
        else:
            chart = (
                alt.Chart(history)
                .mark_bar()
                .encode(
                    x=alt.X("Year:O", title="Year"),
                    y=alt.Y("Price:Q", title="Average unit price"),
                    tooltip=[
                        alt.Tooltip("Year:O", title="Year"),
                        alt.Tooltip("Price:Q", title="Average Price", format="$.2f"),
                    ],
                )
                .properties(height=400)
            )
            st.altair_chart(chart, use_container_width=True)

    st.download_button(
        "Download filtered bids",
        data=display_df.to_csv(index=False),
        file_name="filtered_bids.csv",
        mime="text/csv",
        disabled=display_df.empty,
    )


if __name__ == "__main__":
    main()
