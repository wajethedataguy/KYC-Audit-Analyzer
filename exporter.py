import os
from datetime import datetime
import pandas as pd


def save_excel_output(
    logic_name: str,
    df: pd.DataFrame,
    output_dir: str = ".",
    include_timestamp: bool = True,
    verbose: bool = True,
    sheet_name: str = "Sheet1",
) -> str:
    """
    Saves a DataFrame to an Excel file with a standardized filename.

    Parameters:
    - logic_name: Name of the logic module (used in filename)
    - df: DataFrame to export
    - output_dir: Directory to save the file (default is current directory)
    - include_timestamp: Whether to append timestamp to filename
    - verbose: Whether to print status messages
    - sheet_name: Name of the Excel sheet (default is 'Sheet1')

    Returns:
    - Full path to the saved file (or empty string if skipped)
    """
    if df is None or df.empty:
        if verbose:
            print(f"⚠️ No data to export for {logic_name}. Skipping Excel save.")
        return ""

    # Ensure output directory exists
    os.makedirs(output_dir, exist_ok=True)

    # Build filename
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S") if include_timestamp else ""
    filename = (
        f"{logic_name}_{timestamp}.xlsx" if include_timestamp else f"{logic_name}.xlsx"
    )
    filepath = os.path.join(output_dir, filename)

    try:
        with pd.ExcelWriter(filepath, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name)
        if verbose:
            print(f"📁 Saved output to: {filepath}")
        return filepath
    except Exception as e:
        if verbose:
            print(f"❌ Failed to save Excel for {logic_name}: {e}")
        return ""
