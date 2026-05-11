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
    name="020_monthly_credit_turnover_contradiction",
    description="Flags customers where declared monthly credit turnover is 'Above 10M' or 'Above 50M' but numeric GT value is below or equal to threshold.",
    category="Compliance & Screening",
)
def logic_020_monthly_credit_turnover_contradiction(
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

        # 🔧 Clean identifiers
        def clean_id(val):
            try:
                num = float(val)
                return str(int(num)) if num.is_integer() else str(val).strip()
            except:
                return str(val).strip()

        df["CUSTOMER_NUMBER"] = df["CUSTOMER_NUMBER"].fillna("").apply(clean_id)
        df["ACCOUNT_NUMBER"] = df["ACCOUNT_NUMBER"].fillna("").apply(clean_id)

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

        # 🔍 Convert GT numeric values safely
        df["NUMERIC_GT"] = pd.to_numeric(
            df["EXPECTED_MONTHLY_TURNOVER_GT"], errors="coerce"
        )

        # 🧠 Apply contradiction logic
        contradiction_mask = (
            (df["INDIVIDUAL_NARRATIVE"] == "above 10m") & (df["NUMERIC_GT"] <= 10000000)
        ) | (
            (df["INDIVIDUAL_NARRATIVE"] == "above 50m") & (df["NUMERIC_GT"] <= 50000000)
        )

        df_filtered = df[contradiction_mask].copy()

        # 📤 Prepare output
        output_columns_raw = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "PURPOSE_OF_ACCOUNT",
            "EXPECTED_MONTHLY_TURNOVER_INDIVIDUAL",
            "EXPECTED_MONTHLY_TURNOVER_GT",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]

        output_columns_final = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "Purpose_of_Account",
            "Expected_Monthly_Turnover_Individual",
            "Expected_Monthly_Turnover_GT",
            "ProductDesc",
            "Sector_Description",
        ]

        output = df_filtered[output_columns_raw].copy()
        output.columns = output_columns_final
        output = output.reset_index(drop=True)

        # 🧯 Handle empty output
        if output.empty:
            print("✅ No monthly credit turnover contradictions found.")
            output = get_empty_output(output_columns_final)

        # 📁 Controlled export
        if mode == "full" and not output.isna().all(axis=1).iloc[0]:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_path = f"monthly_credit_turnover_contradiction_{timestamp}.xlsx"
            output.to_excel(file_path, index=False)
            print(f"📁 Saved output to: {file_path}")
        elif mode == "full":
            print(
                "✅ No monthly credit turnover contradictions found. Excel file not created."
            )

        return output

    except Exception as e:
        print(f"❌ Exception occurred: {e}")
        fallback_columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "Purpose_of_Account",
            "Expected_Monthly_Turnover_Individual",
            "Expected_Monthly_Turnover_GT",
            "ProductDesc",
            "Sector_Description",
        ]
        return get_empty_output(fallback_columns)
