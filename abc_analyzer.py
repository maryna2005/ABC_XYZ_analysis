import pandas as pd
import numpy as np
import os
import sys

# --- PATH SETTINGS ---
# Set the base path to the project folder.
# This path uses the MAIN project directory.
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_PATH = os.path.join(BASE_DIR, 'data', 'input')
OUTPUT_PATH = os.path.join(BASE_DIR, 'data', 'output')

# Create output folder if it doesn't exist
if not os.path.exists(OUTPUT_PATH):
    os.makedirs(OUTPUT_PATH)
    print(f"Created output directory: {OUTPUT_PATH}")


def load_data(stock_file: str = 'Stock.xlsx', cogs_file: str = 'COGS.xlsx') -> tuple[pd.DataFrame, pd.DataFrame]:
    """Loads Stock and COGS data from the specified input directory."""
    try:
        stock_path = os.path.join(INPUT_PATH, stock_file)
        cogs_path = os.path.join(INPUT_PATH, cogs_file)

        print(f"Loading Stock: {stock_path}")
        stock = pd.read_excel(stock_path)

        print(f"Loading COGS: {cogs_path}")
        cogs = pd.read_excel(cogs_path)

        # Check for required columns
        if 'Date' not in stock.columns or 'Stock' not in stock.columns:
            raise ValueError("Stock.xlsx must contain 'Date' and 'Stock' columns.")
        if 'SKU' not in stock.columns or 'SKU' not in cogs.columns:
            raise ValueError("Both files must contain the 'SKU' column.")
        if 'COGS' not in cogs.columns:
            raise ValueError("COGS.xlsx must contain the 'COGS' column.")

        return stock, cogs

    except FileNotFoundError as e:
        print("\n" + "=" * 80)
        print("ERROR: File not found!")
        print("Please ensure you have created the 'data/input' folder and placed these files there:")
        print(f"1. {stock_file}")
        print(f"2. {cogs_file}")
        print(f"Expected path: {INPUT_PATH}")
        print("=" * 80 + "\n")
        sys.exit(1)
    except ValueError as e:
        print(f"\nDATA FORMAT ERROR: {e}\n")
        sys.exit(1)


def transform_data(stock: pd.DataFrame, cogs: pd.DataFrame) -> pd.DataFrame:
    """Performs data transformation according to the instructions."""

    # --- 1. Create 'Period' column ---
    # New 'Period' column based on 'Date'
    print("Transformation: Creating 'Period' column...")
    stock['Period'] = pd.to_datetime(stock['Date']).dt.strftime('%Y-%m')
    del stock['Date']

    # --- 2. Vlookup COGS column to Stock table ---
    print("Transformation: Merging COGS (Vlookup)...")
    stock_cogs = pd.merge(stock, cogs, how='left', on='SKU')

    # --- 3. New Stock Value column creation ---
    print("Transformation: Creating 'Value' column...")
    # Handle NaN values in COGS after merge
    stock_cogs['COGS'] = stock_cogs['COGS'].fillna(0)
    stock_cogs['Value'] = stock_cogs['Stock'] * stock_cogs['COGS']

    # --- 4. Choose columns in a specific order ---
    print("Transformation: Selecting final columns...")
    required_cols = ['SKU', 'Period', 'Value']
    # Use .copy() to avoid SettingWithCopyWarning
    abc_input = stock_cogs[required_cols].copy()

    return abc_input


def assign_abc_groups(df: pd.DataFrame, a_threshold: float = 0.80, b_threshold: float = 0.95) -> pd.DataFrame:
    """
    Assigns ABC classification groups to SKUs by period 
    based on cumulative value contribution.
    """

    print("Executing ABC classification...")

    # === 1. Input validation ===
    required_columns = {'SKU', 'Period', 'Value'}
    if not required_columns.issubset(df.columns):
        missing = required_columns - set(df.columns)
        raise ValueError(f"Missing required columns: {missing}")

    if not pd.api.types.is_numeric_dtype(df['Value']):
        raise TypeError("'Value' column must be numeric (float or int).")

    # Fill missing values in the Value column
    df = df.copy()
    df['Value'] = df['Value'].fillna(0.0)

    # === 2. Aggregate total values per SKU/Period ===
    # Sum values to ensure a single row per SKU for each period
    agg_df = (
        df.groupby(['Period', 'SKU'], as_index=False)['Value']
        .sum()
    )

    # === 3. Sort, rank and compute cumulative metrics ===
    agg_df = agg_df.sort_values(['Period', 'Value'], ascending=[True, False])

    # Calculate total value per period
    agg_df['Total_Period_Value'] = agg_df.groupby('Period')['Value'].transform('sum')

    # Calculate cumulative value per period
    agg_df['Cumulative_Value'] = agg_df.groupby('Period')['Value'].cumsum()

    # Calculate cumulative percentage (protection against division by zero)
    agg_df['Cumulative_Percent'] = np.where(
        agg_df['Total_Period_Value'] == 0,
        1.0,
        agg_df['Cumulative_Value'] / agg_df['Total_Period_Value']
    )

    # === 4. ABC Group assignment ===
    def classify_abc(cum_pct: float, total_val: float) -> str:
        """Determines the ABC class."""
        if total_val == 0:
            return 'C'
        # Group A: Top 80% (0.0 - 0.80)
        if cum_pct <= a_threshold:
            return 'A'
        # Group B: Next 15% (0.80 - 0.95)
        elif cum_pct <= b_threshold:
            return 'B'
        # Group C: Bottom 5% (0.95 - 1.00)
        else:
            return 'C'

    # Apply classification
    agg_df['ABC_Group'] = agg_df.apply(
        lambda x: classify_abc(x['Cumulative_Percent'], x['Total_Period_Value']),
        axis=1
    )

    # === 5. Merge results back into original DataFrame ===
    # Join the ABC group back to the original DataFrame
    result_df = pd.merge(
        df,
        agg_df[['Period', 'SKU', 'ABC_Group']],
        on=['Period', 'SKU'],
        how='left'
    )

    # === 6. Validation (check if all rows received a group) ===
    if result_df['ABC_Group'].isna().any():
        # Add debugging details
        missing_count = result_df['ABC_Group'].isna().sum()
        raise RuntimeError(f"ABC group assignment failed for {missing_count} rows. Check input data.")

    print("ABC classification successfully completed.")
    return result_df


def save_local_file(data: pd.DataFrame, name: str) -> None:
    """
    Saves DataFrame to an Excel file in the output directory.
    Uses openpyxl or xlsxwriter engine.
    """
    dump_file_name = f"{name}.xlsx"
    data_dump = os.path.join(OUTPUT_PATH, dump_file_name)

    print(f"\nSaving data to file: {data_dump}")

    # Determine which engine is available
    engine_used = "openpyxl"
    try:
        import xlsxwriter
        engine_used = "xlsxwriter"
    except ImportError:
        pass  # openpyxl is usually available by default with pandas

    try:
        writer = pd.ExcelWriter(data_dump, engine=engine_used)
        data.to_excel(writer, sheet_name="ABC_Result", index=False)
        writer.close()
        print(f"✅ Data successfully saved using {engine_used.upper()}.")
    except Exception as e:
        print(f"❌ Error saving file: {e}")
        print("Please ensure the file is not open in another program.")


# --- MAIN SCRIPT LOGIC ---
if __name__ == "__main__":

    print("--- Starting Automated ABC Analysis ---")

    # 1. Load data
    # Note: change file names if they differ
    stock_df, cogs_df = load_data(stock_file='Stock.xlsx', cogs_file='COGS.xlsx')

    # 2. Transform data
    abc_input_df = transform_data(stock_df, cogs_df)

    # 3. Execute ABC classification function
    try:
        df_result = assign_abc_groups(abc_input_df)
    except Exception as e:
        print(f"\n❌ Critical error in ABC analysis: {e}")
        sys.exit(1)

    # 4. Save output file
    save_local_file(df_result, "abc_analysis_output")

    print("\n--- Script completed successfully! ---")