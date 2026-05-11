import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


def get_empty_output(columns: list) -> pd.DataFrame:
    """
    Returns a single-row DataFrame filled with pd.NA for the given columns.
    Used when no contradictions are found or an error occurs.
    """
    return pd.DataFrame([{col: pd.NA for col in columns}])


@register_logic(
    name="030_missing_expected_credit_txns",
    description="Flags customers where Expected Monthly Credit Transactions is missing in KYC profile.",
    category="Compliance & Screening",
)
def logic_030_missing_expected_credit_txns(
    dataframes: dict, mode="preview"
) -> pd.DataFrame:
    try:
        # 🔍 Locate the merged file
        merged_key = next(
            (k for k in dataframes if "merged_file.xlsx" in k.lower()), None
        )
        if not merged_key:
            raise ValueError("Merged file not found in uploaded dataframes.")

        df = dataframes[merged_key].copy()

        # 🔧 Normalize column names
        df.columns = (
            df.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )

        # ✅ Required columns
        required_cols = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "EXPECTED_MONTHLY_CREDIT_TRANSACTIONS",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            raise ValueError(f"Missing required columns: {missing_cols}")

        # 🔢 Clean numeric field
        df["EXPECTED_CREDIT_TXNS_CLEAN"] = pd.to_numeric(
            df["EXPECTED_MONTHLY_CREDIT_TRANSACTIONS"], errors="coerce"
        )

        # 🧠 Apply missing logic: flag rows where value is missing or NaN
        df_filtered = df[df["EXPECTED_CREDIT_TXNS_CLEAN"].isna()].copy()

        # 📤 Prepare output
        output_columns_raw = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "EXPECTED_MONTHLY_CREDIT_TRANSACTIONS",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        output_columns_final = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "Expected_Monthly_Credit_Transactions",
            "ProductDesc",
            "Sector_Description",
        ]

        output = df_filtered[output_columns_raw].copy()
        output.columns = output_columns_final
        output = output.reset_index(drop=True)

        # 🧯 Handle empty output
        if output.empty:
            print("✅ No missing expected monthly credit transactions found.")
            return get_empty_output(output_columns_final)

        output = output.astype(object).where(pd.notna(output), None)

        # 📁 Optional export
        if mode == "full" and not output.isna().all(axis=1).iloc[0]:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_path = f"missing_expected_credit_txns_{timestamp}.xlsx"
            output.to_excel(file_path, index=False)
            print(f"✅ Logic 030 output saved at {file_path}")
        elif mode == "full":
            print("✅ No mismatches found. Excel file not created.")

        return output

    except Exception as e:
        print(f"❌ Error in logic_030_missing_expected_credit_txns: {e}")
        fallback_columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "Expected_Monthly_Credit_Transactions",
            "ProductDesc",
            "Sector_Description",
        ]
        return get_empty_output(fallback_columns)
