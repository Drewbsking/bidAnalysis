"""Streamlit tool for comparing annual bid results stored in Excel workbooks."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Dict, IO, Optional, Union

import altair as alt
import pandas as pd
import streamlit as st

BASE_DIR = Path(__file__).parent
ALL_SHEET_NAME = "All"
MASTER_WORKBOOK_NAME = "PRS25-7-3.xlsx"
MASTER_SHEET_NAME = "Sheet"
MASTER_ITEM_COLUMN = "Item"
MASTER_COMB_COLUMN = "Comb"
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


@st.cache_data(show_spinner=False)
def _load_master_comb_map() -> Dict[float, str]:
    """Load Item -> Comb mapping from the master workbook."""
    path = BASE_DIR / MASTER_WORKBOOK_NAME
    if not path.exists():
        return {}
    try:
        master = pd.read_excel(path, sheet_name=MASTER_SHEET_NAME, engine="openpyxl")
    except Exception:
        return {}
    required_cols = {MASTER_ITEM_COLUMN, MASTER_COMB_COLUMN}
    if not required_cols.issubset(master.columns):
        return {}

    item_numbers = pd.to_numeric(master[MASTER_ITEM_COLUMN], errors="coerce")
    comb_values = master[MASTER_COMB_COLUMN].fillna("").astype(str).str.strip()
    mapping_df = pd.DataFrame({"Item No": item_numbers, "Comb": comb_values})
    mapping_df = mapping_df.dropna(subset=["Item No"])
    mapping_df = mapping_df[mapping_df["Comb"] != ""]
    mapping_df = mapping_df.drop_duplicates("Item No", keep="first")
    return dict(zip(mapping_df["Item No"], mapping_df["Comb"]))


def _apply_master_descriptions(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize Description using master Item -> Comb mapping."""
    if df.empty or "Item No" not in df.columns:
        return df
    comb_map = _load_master_comb_map()
    if not comb_map:
        return df

    df = df.copy()
    mapped_desc = df["Item No"].map(comb_map)
    if "Description" in df.columns:
        existing_desc = df["Description"].fillna("").astype(str).str.strip()
        df["Description"] = mapped_desc.where(mapped_desc.notna(), existing_desc)
    else:
        df["Description"] = mapped_desc
    return df


def _load_all_sheet(
    xls: pd.ExcelFile, label: str, year_label: str
) -> Optional[pd.DataFrame]:
    """Parse the wide-format 'All' sheet into the long bid format."""
    df_raw = pd.read_excel(xls, sheet_name=label, header=None, engine="openpyxl")
    if df_raw.shape[0] < 3:
        st.error(f"Workbook '{label}' has an '{label}' sheet but no data rows.")
        return None

    vendor_starts = [i for i, val in enumerate(df_raw.iloc[0]) if pd.notna(val)]
    if not vendor_starts:
        st.error(
            f"Workbook '{label}' has an '{label}' sheet but no vendor names in row 1."
        )
        return None

    base_end = vendor_starts[0]
    raw_headers = df_raw.iloc[1, :base_end].tolist()
    base_headers = [
        str(h).strip() if str(h).strip() else f"Column_{idx}"
        for idx, h in enumerate(raw_headers)
    ]
    base_data = df_raw.iloc[2:, :base_end].copy()
    base_data.columns = base_headers
    if "Quantity" in base_data.columns:
        base_data = base_data.rename(columns={"Quantity": "Item Quantity"})
    base_data = base_data.dropna(how="all").reset_index(drop=True)

    # Derive Item No directly from the All sheet.
    base_data = _deduplicate_columns(base_data)
    base_data = _ensure_item_column(base_data)
    code_item_no = pd.Series(index=base_data.index, dtype="float64")
    if "Description" in base_data.columns:
        code_item_no = pd.to_numeric(
            base_data["Description"]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.extract(r"^([0-9]{5,})")[0],
            errors="coerce",
        )
    if "Item No" not in base_data.columns:
        # Try extracting leading digits from Description or first column.
        desc_col = None
        if "Description" in base_data.columns:
            desc_col = "Description"
        elif base_data.columns:
            desc_col = base_data.columns[0]
        if desc_col:
            extracted = (
                base_data[desc_col]
                .astype(str)
                .str.extract(r"^([0-9]+)")
                .iloc[:, 0]
                .astype(float)
            )
            base_data["Item No"] = pd.to_numeric(extracted, errors="coerce")
    else:
        base_data["Item No"] = pd.to_numeric(base_data["Item No"], errors="coerce")

    if code_item_no.notna().any():
        base_data["Item No"] = code_item_no.combine_first(base_data["Item No"])

    if "Item No" not in base_data.columns:
        base_data["Item No"] = pd.RangeIndex(start=1, stop=len(base_data) + 1)
    allowed_items = base_data["Item No"].dropna().unique()

    rows: list[pd.DataFrame] = []
    total_cols = df_raw.shape[1]
    for idx, start in enumerate(vendor_starts):
        vendor_name = str(df_raw.iat[0, start]).strip()
        next_start = (
            vendor_starts[idx + 1] if idx + 1 < len(vendor_starts) else total_cols
        )
        block_width = next_start - start
        vendor_cols = df_raw.iloc[1, start : start + block_width].tolist()
        vendor_block = df_raw.iloc[2:, start : start + block_width].copy()
        vendor_block.columns = vendor_cols
        vendor_block["Organization Name"] = vendor_name

        combined = pd.concat(
            [base_data.reset_index(drop=True), vendor_block.reset_index(drop=True)],
            axis=1,
        )
        combined["Year"] = int(year_label)
        rows.append(combined)

    dataset = pd.concat(rows, ignore_index=True)
    dataset = _deduplicate_columns(dataset)
    dataset = _ensure_item_column(dataset)
    dataset = _coerce_numeric(dataset, NUMERIC_COLUMNS)
    dataset = _apply_master_descriptions(dataset)
    if len(allowed_items):
        dataset = dataset[dataset["Item No"].isin(allowed_items)]
    if "Description" in dataset.columns:
        dataset = dataset.dropna(subset=["Item No", "Description"], how="all")
    else:
        dataset = dataset.dropna(subset=["Item No"], how="all")
    value_cols = [
        col
        for col in dataset.columns
        if col.startswith(("Price", "Total Cost", "Quantity"))
    ]
    if value_cols:
        dataset = dataset[dataset[value_cols].notna().any(axis=1)]
    if "Description" in dataset.columns:
        description_text = dataset["Description"].fillna("").astype(str).str.strip()
        dataset = dataset[description_text != ""]
    dataset = _add_item_labels(dataset)
    return dataset


def _load_year_dataset(
    source: WorkbookSource, label: str, year_label: str
) -> Optional[pd.DataFrame]:
    """Read the All sheet for a workbook."""
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

    all_sheet = _resolve_sheet_name(xls, ALL_SHEET_NAME)
    if all_sheet is None:
        st.error(f"Workbook '{label}' is missing an '{ALL_SHEET_NAME}' sheet.")
        return None

    return _load_all_sheet(xls, all_sheet, year_label)


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
            options=["All bids", "Lowest bid"],
            index=0,
            help="Choose whether to analyze all bids or only the lowest bid per item/year.",
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
            item_query = st.text_input(
                "Description contains",
                value="",
                help="Literal match. Example: R1-1 matches only descriptions containing R1-1.",
            ).strip()
            if item_query:
                literal_query = re.escape(item_query)
                item_info = item_info[
                    item_info["Item Label"].astype(str).str.contains(
                        literal_query, case=False, na=False, regex=True
                    )
                ]
            item_options = item_info["Item Label"].tolist()
            selector_mode = st.radio(
                "Item selector",
                options=["List view", "Tag view"],
                index=0,
                horizontal=True,
            )
            if selector_mode == "List view":
                selection_key = "item_label_selection_map"
                stored_selection = st.session_state.get(selection_key, {})
                if not isinstance(stored_selection, dict):
                    stored_selection = {}
                current_selection = {
                    label: bool(stored_selection.get(label, True))
                    for label in item_options
                }
                list_df = pd.DataFrame(
                    {
                        "Select": [current_selection[label] for label in item_options],
                        "Item": item_options,
                    }
                )
                edited_df = st.data_editor(
                    list_df,
                    hide_index=True,
                    use_container_width=True,
                    height=260,
                    disabled=["Item"],
                    column_config={
                        "Select": st.column_config.CheckboxColumn("Select"),
                        "Item": st.column_config.TextColumn("Item"),
                    },
                    key="item_selector_list_editor",
                )
                selected_labels = edited_df.loc[edited_df["Select"], "Item"].tolist()
                st.session_state[selection_key] = dict(
                    zip(edited_df["Item"], edited_df["Select"])
                )
            else:
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

    if mode == "Lowest bid":
        weighted_avg = _weighted_average_lowest(df)
        label = "Weighted avg (lowest bid)"
    else:
        weighted_avg = _weighted_average_all(df)
        label = "Weighted avg (all)"
    weighted_display = f"${weighted_avg:,.2f}" if pd.notna(weighted_avg) else "N/A"
    col_weighted.metric(label, weighted_display)


def prepare_price_history(
    df: pd.DataFrame, mode: str, year_domain: Optional[list[int]] = None
) -> pd.DataFrame:
    if df.empty or "Year" not in df.columns:
        return pd.DataFrame()

    base = df.copy()
    base["Year"] = pd.to_numeric(base["Year"], errors="coerce")
    base = base.dropna(subset=["Year"])
    if base.empty:
        return pd.DataFrame()
    base["Year"] = base["Year"].astype(int)

    if year_domain:
        years = sorted({int(y) for y in year_domain})
    else:
        years = sorted(base["Year"].dropna().unique().tolist())
    if not years:
        return pd.DataFrame()

    if mode == "Lowest bid":
        if "Item No" not in base.columns:
            return pd.DataFrame()
        if "Item Label" not in base.columns:
            base["Item Label"] = "Item " + base["Item No"].astype(str)
        item_info = (
            base[["Item No", "Item Label"]]
            .dropna(subset=["Item No"])
            .drop_duplicates("Item No")
        )
        if item_info.empty:
            return pd.DataFrame()

        if "Price" in base.columns:
            price_rows = base.dropna(subset=["Price"]).copy()
            if not price_rows.empty:
                price_rows = price_rows.groupby(["Item No", "Year"], as_index=False)[
                    "Price"
                ].mean()
            else:
                price_rows = pd.DataFrame(columns=["Item No", "Year", "Price"])
        else:
            price_rows = pd.DataFrame(columns=["Item No", "Year", "Price"])

        year_df = pd.DataFrame({"Year": years})
        year_df["_key"] = 1
        item_info = item_info.copy()
        item_info["_key"] = 1
        grid = item_info.merge(year_df, on="_key", how="inner").drop(columns=["_key"])
        result = grid.merge(price_rows, on=["Item No", "Year"], how="left")
        result["Bid Status"] = result["Price"].where(
            result["Price"].notna(), "Not bid this year"
        )
        result["Bid Status"] = result["Bid Status"].where(
            result["Bid Status"] == "Not bid this year", "Bid"
        )
        return result

    if "Price" not in base.columns:
        return pd.DataFrame()

    summary = (
        base.dropna(subset=["Price"]).groupby("Year", as_index=False)["Price"].mean()
    )
    result = pd.DataFrame({"Year": years}).merge(summary, on="Year", how="left")
    result["Price"] = result["Price"].round(2)
    result["Item Label"] = "Average of all bids"
    result["Bid Status"] = result["Price"].where(
        result["Price"].notna(), "Not bid this year"
    )
    result["Bid Status"] = result["Bid Status"].where(
        result["Bid Status"] == "Not bid this year", "Bid"
    )
    return result


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
        result[earliest_col] = pd.to_numeric(result[earliest_col], errors="coerce")
        result[latest_col] = pd.to_numeric(result[latest_col], errors="coerce")
        result["Price Δ"] = result[latest_col] - result[earliest_col]
        denom = result[earliest_col].replace(0, pd.NA)
        result["Price Δ%"] = ((result[latest_col] - result[earliest_col]) / denom) * 100

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
    if mode == "Lowest bid":
        display_df = compute_lowest_bid_view(filtered)
    else:
        display_df = filtered

    show_metrics(display_df, mode)

    st.subheader("Bid detail")
    st.dataframe(display_df, use_container_width=True)

    lowest = compute_lowest_bid_table(filtered)
    st.subheader("Lowest bid by item")
    if lowest.empty:
        st.info("Select both years with price data to compare lowest bids.")
    else:
        st.dataframe(lowest, use_container_width=True)
        st.download_button(
            "Download lowest-bid comparison",
            data=lowest.to_csv(index=False),
            file_name="lowest_bid_comparison.csv",
            mime="text/csv",
        )

    st.subheader("Price history")
    selected_year_domain = (
        sorted(filtered["Year"].dropna().unique().tolist())
        if "Year" in filtered.columns
        else None
    )
    history = prepare_price_history(display_df, mode, year_domain=selected_year_domain)
    if history.empty:
        st.info("No price history available for the selected scope.")
    else:
        if mode == "Lowest bid":
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
                chart_data = history[history["Item No"].isin(selected_items)].copy()
                year_sort = sorted(chart_data["Year"].dropna().unique().tolist())
                bid_rows = chart_data[chart_data["Price"].notna()]
                not_bid_rows = chart_data[chart_data["Price"].isna()]

                base_encoding = {
                    "x": alt.X("Year:O", title="Year", sort=year_sort),
                    "color": alt.Color("Item Label:N", title="Item"),
                }
                line = (
                    alt.Chart(bid_rows)
                    .mark_line(point=True)
                    .encode(
                        **base_encoding,
                        y=alt.Y("Price:Q", title="Unit price"),
                        tooltip=[
                            alt.Tooltip("Item Label:N", title="Item"),
                            alt.Tooltip("Year:O", title="Year"),
                            alt.Tooltip("Price:Q", title="Price", format="$.2f"),
                            alt.Tooltip("Bid Status:N", title="Status"),
                        ],
                    )
                )

                if not_bid_rows.empty:
                    chart = line.properties(height=400)
                else:
                    if len(selected_items) == 1:
                        not_bid_mark = alt.Chart(not_bid_rows).mark_text(
                            text="Not bid", dy=-8, fontSize=10
                        )
                    else:
                        not_bid_mark = alt.Chart(not_bid_rows).mark_point(
                            shape="cross", size=80
                        )
                    not_bid = not_bid_mark.encode(
                        **base_encoding,
                        y=alt.value(390),
                        tooltip=[
                            alt.Tooltip("Item Label:N", title="Item"),
                            alt.Tooltip("Year:O", title="Year"),
                            alt.Tooltip("Bid Status:N", title="Status"),
                        ],
                    )
                    chart = (line + not_bid).properties(height=400)

                st.altair_chart(chart, use_container_width=True)
        else:
            year_sort = sorted(history["Year"].dropna().unique().tolist())
            bid_rows = history[history["Price"].notna()]
            not_bid_rows = history[history["Price"].isna()]
            bars = (
                alt.Chart(bid_rows)
                .mark_bar()
                .encode(
                    x=alt.X("Year:O", title="Year", sort=year_sort),
                    y=alt.Y("Price:Q", title="Average unit price"),
                    tooltip=[
                        alt.Tooltip("Year:O", title="Year"),
                        alt.Tooltip("Price:Q", title="Average Price", format="$.2f"),
                        alt.Tooltip("Bid Status:N", title="Status"),
                    ],
                )
            )
            if not_bid_rows.empty:
                chart = bars.properties(height=400)
            else:
                notes = (
                    alt.Chart(not_bid_rows)
                    .mark_text(text="Not bid", dy=-8, fontSize=10)
                    .encode(
                        x=alt.X("Year:O", title="Year", sort=year_sort),
                        y=alt.value(390),
                        tooltip=[
                            alt.Tooltip("Year:O", title="Year"),
                            alt.Tooltip("Bid Status:N", title="Status"),
                        ],
                    )
                )
                chart = (bars + notes).properties(height=400)
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
