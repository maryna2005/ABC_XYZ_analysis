import pandas as pd
import numpy as np
import os
import sys
from typing import Literal

# --- PATH SETTINGS ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_PATH = os.path.join(BASE_DIR, 'data', 'input')
OUTPUT_PATH = os.path.join(BASE_DIR, 'data', 'output')

# Create output folder if it doesn't exist
if not os.path.exists(OUTPUT_PATH):
    os.makedirs(OUTPUT_PATH)


def load_stock_data(file_name: str = 'Stock.xlsx') -> pd.DataFrame:
    """Loads Stock data from the input directory."""
    file_path = os.path.join(INPUT_PATH, file_name)
    print(f"Loading data from: {file_path}")

    if not os.path.exists(file_path):
        print(f"\n❌ ERROR: File '{file_name}' not found in {INPUT_PATH}")
        sys.exit(1)

    try:
        df = pd.read_excel(file_path)
        return df
    except Exception as e:
        print(f"❌ Error reading Excel: {e}")
        sys.exit(1)


def prepare_data(stock: pd.DataFrame) -> pd.DataFrame:
    """Prepares data: creates Period column and renames Stock to Value."""
    print("Preparing data...")

    # Check for required columns
    if 'Date' not in stock.columns:
        raise ValueError("Column 'Date' is missing from the file")

    # Create Period column (YYYY-MM)
    stock['Period'] = pd.to_datetime(stock['Date']).dt.strftime('%Y-%m')
    del stock['Date']

    # Rename Stock -> Value for XYZ function universality
    if 'Stock' in stock.columns:
        stock.rename(columns={'Stock': 'Value'}, inplace=True)
    elif 'Value' not in stock.columns:
        # If neither Stock nor Value is present - it's an error
        raise ValueError("File must contain either 'Stock' (quantity) or 'Value' column")

    return stock


def assign_xyz_groups(
        df: pd.DataFrame,
        data_mode: Literal["dense", "sparse"] = "dense"
) -> pd.DataFrame:
    """
    Performs XYZ analysis (classification by stability).

    data_mode="dense" (default):
        Fills missing months with zeros. Crucial for correct stability calculation.
        If a product was out of stock for a month, it affects its reliability.
    """
    print(f"Executing XYZ analysis (mode: {data_mode})...")

    # --- 0. Validation ---
    if data_mode not in ["dense", "sparse"]:
        raise ValueError(f"data_mode must be 'dense' or 'sparse', not '{data_mode}'")

    required = {"SKU", "Period", "Value"}
    if not required.issubset(df.columns):
        missing = required - set(df.columns)
        raise ValueError(f"Missing columns: {missing}")

    df_out = df.copy()

    # --- 1. Convert Value to numeric ---
    if not pd.api.types.is_numeric_dtype(df_out["Value"]):
        df_out["Value"] = pd.to_numeric(df_out["Value"], errors="coerce").fillna(0.0)
    else:
        df_out["Value"] = df_out["Value"].fillna(0.0)

    # --- 2. Aggregation (SKU, Period) ---
    period_sum = (
        df_out.groupby(["SKU", "Period"], as_index=False)["Value"]
        .sum()
    )

    # --- 3. Data Densification ---
    if data_mode == "dense":
        # Create a grid of all possible SKUs and Periods
        all_skus = df_out['SKU'].unique()
        all_periods = df_out['Period'].unique()

        df_scaffold = pd.MultiIndex.from_product(
            [all_skus, all_periods],
            names=['SKU', 'Period']
        ).to_frame(index=False)

        # Merge actual data with the full grid
        df_dense = pd.merge(
            df_scaffold,
            period_sum,
            on=["SKU", "Period"],
            how="left"
        )
        # Fill gaps with zeros
        df_dense["Value"] = df_dense["Value"].fillna(0.0)
        stats_input_df = df_dense
    else:
        stats_input_df = period_sum

    # --- 4. Statistics Calculation (Mean and Standard Deviation) ---
    sku_stats = (
        stats_input_df.groupby("SKU", as_index=False)["Value"]
        .agg(n_periods="count", mean_value="mean", std_value="std")
    )

    # For SKUs with only one period, std will be NaN -> replace with 0
    sku_stats["std_value"] = sku_stats["std_value"].fillna(0.0)

    # --- 5. Coefficient of Variation (CV) Calculation ---
    # CV = std / mean
    sku_stats['cv'] = np.where(
        sku_stats['mean_value'] == 0,
        0.0,
        sku_stats['std_value'] / np.abs(sku_stats['mean_value'])
    )
    # Remove infinite values
    sku_stats['cv'] = sku_stats['cv'].replace([np.inf, -np.inf], np.nan)

    # --- 6. Define Thresholds (X, Y, Z) ---
    MIN_PERIODS = 2

    # Consider only SKUs with sales and enough periods
    is_eligible = (
            (sku_stats['n_periods'] >= MIN_PERIODS) &
            (sku_stats['mean_value'] != 0)
    )
    eligible_cvs = sku_stats.loc[is_eligible, 'cv'].dropna()

    if eligible_cvs.empty:
        x_threshold = 0.0
        y_threshold = 0.0
    else:
        # AUTOMATIC SPLIT: 33% / 33% / 33%
        quantiles = eligible_cvs.quantile([0.33, 0.66])
        x_threshold = quantiles[0.33]
        y_threshold = quantiles[0.66]

        # Manual mode (commented):
        # x_threshold = 0.5  (CV < 50% is stable)
        # y_threshold = 1.0  (CV < 100% is medium)

    print(f"CV thresholds determined automatically: X <= {x_threshold:.2f}, Y <= {y_threshold:.2f}")

    # --- 7. Classification ---
    conditions = [
        (sku_stats['n_periods'] < MIN_PERIODS),  # Not enough data
        (sku_stats['mean_value'] == 0),           # No stock/sales
        (sku_stats['cv'].isna()),                # CV error
        (sku_stats['cv'] <= x_threshold),        # Group X
        (sku_stats['cv'] <= y_threshold)         # Group Y
    ]

    choices = ["", "", "", "X", "Y"]
    default_choice = "Z"  # Everything else is Group Z

    sku_stats['XYZ'] = np.select(conditions, choices, default=default_choice)

    # --- 8. Merge results back to original data ---
    # Add XYZ labels to each row of the original table
    cols_to_merge = sku_stats[["SKU", "XYZ", "cv"]].rename(columns={"cv": "CV"})

    merged = pd.merge(
        df_out,
        cols_to_merge,
        on="SKU",
        how="left"
    )
    merged["XYZ"] = merged["XYZ"].fillna("")

    # Add thresholds for reference
    merged['x_threshold_33'] = x_threshold
    merged['y_threshold_66'] = y_threshold

    # Final column sorting
    new_cols = ["XYZ", "CV", "x_threshold_33", "y_threshold_66"]
    original_cols = [c for c in df.columns if c not in new_cols]
    final_df = merged[original_cols + new_cols]

    return final_df


def save_local_file(data: pd.DataFrame, name: str) -> None:
    """Saves results to Excel."""
    file_name = f"{name}.xlsx"
    save_path = os.path.join(OUTPUT_PATH, file_name)

    print(f"Saving file: {save_path}")

    engine = "openpyxl"
    try:
        import xlsxwriter
        engine = "xlsxwriter"
    except ImportError:
        pass

    try:
        with pd.ExcelWriter(save_path, engine=engine) as writer:
            data.to_excel(writer, sheet_name="XYZ_Analysis", index=False)
        print(f"✅ Successfully saved: {file_name}")
    except Exception as e:
        print(f"❌ Error saving file: {e}")


# --- MAIN EXECUTION BLOCK ---
if __name__ == "__main__":
    print("--- STARTING XYZ ANALYSIS ---")

    # 1. Load
    stock_df = load_stock_data('Stock.xlsx')

    # 2. Prepare
    try:
        xyz_input = prepare_data(stock_df)
    except ValueError as e:
        print(f"❌ Data error: {e}")
        sys.exit(1)

    # 3. Analyze
    result_df = assign_xyz_groups(xyz_input, data_mode="dense")

    # 4. Save
    save_local_file(result_df, "xyz_analysis_output")

    print("--- FINISHED ---")