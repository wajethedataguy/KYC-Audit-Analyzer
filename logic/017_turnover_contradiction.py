import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


def get_empty_output(columns: list) -> pd.DataFrame:
    """
    Returns a single-row DataFrame filled with pd.NA for the given columns.
    Ensures audit-safe output even when no data matches.
    """
    return pd.DataFrame([{col: pd.NA for col in columns}])


@register_logic(
    name="017_turnover_contradiction",
    description="Flags customers where narrative turnover is low but numeric value is positive or field contains invalid characters.",
    category="Compliance & Screening",
)
def logic_017_turnover_contradiction(
    dataframes: dict, mode: str = "preview"
) -> pd.DataFrame:
    try:
        # 🔍 Locate merged file
        merged_key = next(
            (k for k in dataframes if "merged_file.xlsx" in k.lower()), None
        )
        if merged_key is None:
            raise ValueError("Merged file not found in input dataframes.")
        df = dataframes[merged_key].copy()

        # 🔧 Normalize column names
        df.columns = (
            df.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )

        # ✅ Required columns
        required_columns = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "PURPOSE_OF_ACCOUNT",
            "ACCOUNT_TURNOVER",
            "ACCOUNT_TURNOVERGT1050M",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        missing = [col for col in required_columns if col not in df.columns]
        if missing:
            raise ValueError(f"Missing required columns: {missing}")

        # 🔧 Clean identifiers
        def clean_id(val):
            try:
                num = float(val)
                return str(int(num)) if num.is_integer() else str(val).strip()
            except:
                return str(val).strip()

        df["CUSTOMER_NUMBER"] = df["CUSTOMER_NUMBER"].fillna("").apply(clean_id)
        df["ACCOUNT_NUMBER"] = df["ACCOUNT_NUMBER"].fillna("").apply(clean_id)

        # 🔍 Normalize narrative field
        df["NARRATIVE_TURNOVER"] = (
            df["ACCOUNT_TURNOVER"].astype(str).str.lower().str.strip()
        )
        # Narrative NOT in high bands
        df["NARRATIVE_LOW"] = ~df["NARRATIVE_TURNOVER"].isin(["above 10m", "above 50m"])

        # 🔍 Numeric field as float (coerce invalid to NaN)
        df["NUMERIC_TURNOVER"] = pd.to_numeric(
            df["ACCOUNT_TURNOVERGT1050M"], errors="coerce"
        )

        # 🔍 Raw string field for pattern checks
        raw = df["ACCOUNT_TURNOVERGT1050M"].astype(str)
        stripped = raw.str.strip()

        # Blank / null-like values → should NOT be treated as garbage
        blank_or_null = stripped.eq("") | stripped.str.lower().isin(["nan", "null"])

        # Pure numeric strings (optional decimal) → valid numeric, not "garbage"
        is_numeric_like = stripped.str.match(r"^[0-9]+(\.[0-9]+)?$")

        # ✅ We want: any characters other than digits/dot, and not blank/null
        nonnumeric_chars = ~blank_or_null & ~is_numeric_like

        # ✅ Positive numeric turnover
        numeric_positive = df["NUMERIC_TURNOVER"] > 0

        # 🧠 Final contradiction logic:
        df_filtered = df[
            df["NARRATIVE_LOW"] & (numeric_positive | nonnumeric_chars)
        ].copy()

        df_filtered["MISMATCH_REASON"] = (
            "Narrative is low but numeric value is positive or field contains invalid characters"
        )

        # 📤 Prepare output
        output_columns = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "PURPOSE_OF_ACCOUNT",
            "ACCOUNT_TURNOVER",
            "ACCOUNT_TURNOVERGT1050M",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        output = df_filtered[output_columns].reset_index(drop=True)

        # 🧯 Handle empty output
        if output.empty:
            print("✅ No turnover contradictions found.")
            output = get_empty_output(output_columns)

        # 📁 Controlled export
        if mode == "full" and not output.isna().all(axis=1).iloc[0]:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_path = f"turnover_contradiction_{timestamp}.xlsx"
            output.to_excel(file_path, index=False)
            print(f"📁 Saved output to: {file_path}")
        elif mode == "full":
            print("✅ No turnover contradictions found. Excel file not created.")

        return output

    except Exception as e:
        print(f"❌ Exception occurred: {e}")
        fallback_columns = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "PURPOSE_OF_ACCOUNT",
            "ACCOUNT_TURNOVER",
            "ACCOUNT_TURNOVERGT1050M",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        return get_empty_output(fallback_columns)
