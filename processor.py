import pandas as pd


def clean_dataframe(df):
    """
    Cleans the DataFrame by:
    - Dropping rows that are completely empty
    - Stripping whitespace from column names
    """
    df.dropna(how="all", inplace=True)
    df.columns = [col.strip() for col in df.columns]
    return df


def export_dataframe(df, file_path):
    """
    Exports the DataFrame to the specified file path.
    Supports .csv and .xlsx formats.
    """
    if file_path.endswith(".csv"):
        df.to_csv(file_path, index=False)
    elif file_path.endswith(".xlsx"):
        df.to_excel(file_path, index=False)
    else:
        raise ValueError("Unsupported file format")
