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
    name="016_monthly_turnover_mismatch",
    description="Flags customers where 'Expected_Monthly_Turnover_Individual' says Above 10M/50M but GT column is missing.",
    category="Compliance & Screening",
)
def logic_016_monthly_turnover_mismatch(
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
            "EXPECTED_MONTHLY_TURNOVER_INDIVIDUAL",
            "EXPECTED_MONTHLY_TURNOVER_GT",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        missing = [col for col in required_columns if col not in df.columns]
        if missing:
            raise ValueError(f"Missing required columns: {missing}")

        # 🔍 Normalize narratives
        df["INDIVIDUAL_NARRATIVE"] = (
            df["EXPECTED_MONTHLY_TURNOVER_INDIVIDUAL"]
            .astype(str)
            .str.lower()
            .str.strip()
        )
        df["GT_NARRATIVE"] = (
            df["EXPECTED_MONTHLY_TURNOVER_GT"].astype(str).str.lower().str.strip()
        )

        # 🧠 Flag mismatch: Individual says Above 10M/50M but GT is blank/missing
        mismatch_mask = df["INDIVIDUAL_NARRATIVE"].isin(["above 10m", "above 50m"]) & (
            df["GT_NARRATIVE"].isin(["", "nan", "none"])
            | df["EXPECTED_MONTHLY_TURNOVER_GT"].isna()
        )

        df_filtered = df[mismatch_mask].copy()

        # 📤 Prepare output
        output_columns = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "PURPOSE_OF_ACCOUNT",
            "EXPECTED_MONTHLY_TURNOVER_INDIVIDUAL",
            "EXPECTED_MONTHLY_TURNOVER_GT",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        output = df_filtered[output_columns].dropna(how="all").reset_index(drop=True)

        # 🧯 Handle empty output
        if output.empty:
            print("✅ No mismatches found.")
            output = get_empty_output(output_columns)

        return output

    except Exception as e:
        print(f"❌ Exception occurred: {e}")
        fallback_columns = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "PURPOSE_OF_ACCOUNT",
            "EXPECTED_MONTHLY_TURNOVER_INDIVIDUAL",
            "EXPECTED_MONTHLY_TURNOVER_GT",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        return get_empty_output(fallback_columns)
