import pandas as pd


def load_file(path):
    try:
        if path.lower().endswith(".csv"):
            # Read CSV exactly as-is (do NOT auto-convert N/A to NaN)
            df = pd.read_csv(
                path,
                dtype=str,
                low_memory=False,
                keep_default_na=False,  # ✅ CRITICAL
                na_filter=False         # ✅ extra safety
            )
        else:
            # Read first sheet for Excel, preserve text as-is
            xls = pd.ExcelFile(path)
            df = pd.read_excel(
                xls,
                sheet_name=xls.sheet_names[0],
                dtype=str,
                keep_default_na=False   # ✅ CRITICAL
            )

        # Clean up column names
        df.columns = df.columns.str.strip()
        return df

    except Exception as e:
        print(f"❌ Failed to load file: {path}\n   Reason: {type(e).__name__}: {e}")
        return pd.DataFrame()
