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
    name="008_pension_purpose_missing",
    description="Flags pensioner accounts where purpose is not declared as 'Pension'.",
    category="Purpose & Occupation Filter",
)
def logic_008_pension_purpose_missing(dataframes: dict, mode: str = "preview"):
    try:
        # 🔍 Locate merged file
        kyc_key = next((k for k in dataframes if "merged_file.xlsx" in k.lower()), None)
        if kyc_key is None:
            raise ValueError("Merged file not found in input dataframes.")

        df = dataframes[kyc_key].copy()
        if isinstance(df, tuple):
            df = next(iter(df[0].values())) if df[0] else pd.DataFrame()
        if df is None or df.empty:
            raise ValueError("Merged file is empty.")

        # 🔧 Normalize column names
        df.columns = (
            df.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )

        # ✅ Required columns
        required_columns = [
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "CUSTOMER_NUMBER",
            "PURPOSE_OF_ACCOUNT",
            "BAF_PEN_ACCT",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        missing = [col for col in required_columns if col not in df.columns]
        if missing:
            raise ValueError(f"Missing required columns: {missing}")

        # 🔧 Normalize fields
        df["BAF_PEN_ACCT"] = df["BAF_PEN_ACCT"].astype(str).str.upper().str.strip()
        df["PURPOSE_OF_ACCOUNT"] = (
            df["PURPOSE_OF_ACCOUNT"].astype(str).str.upper().str.strip()
        )

        # 🧠 Apply filtering logic
        df_filtered = df[
            (df["BAF_PEN_ACCT"] == "Y")
            & ~df["PURPOSE_OF_ACCOUNT"].str.contains("PENSION")
        ].copy()

        # 🔧 Clean account and customer numbers
        def clean_number(x):
            try:
                return str(int(float(x)))
            except:
                return str(x).strip()

        df_filtered["ACCOUNT_NUMBER"] = df_filtered["ACCOUNT_NUMBER"].apply(
            clean_number
        )
        df_filtered["CUSTOMER_NUMBER"] = df_filtered["CUSTOMER_NUMBER"].apply(
            clean_number
        )

        df_filtered = df_filtered[
            df_filtered["ACCOUNT_NUMBER"].notna()
            & (df_filtered["ACCOUNT_NUMBER"] != "")
        ]

        # 📤 Prepare output
        output = (
            df_filtered[
                [
                    "CUSTOMER_NUMBER",
                    "ACCOUNT_NUMBER",
                    "TITLEOFACCOUNT",
                    "PURPOSE_OF_ACCOUNT",
                    "PRODUCTDESC",
                    "SECTOR_DESCRIPTION",
                ]
            ]
            .drop_duplicates()
            .reset_index(drop=True)
        )

        # 🧼 Rename columns for clarity
        output.columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "Purpose_of_Account",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]

        # 🧯 Handle empty output
        if output.empty:
            print("✅ No pension purpose mismatches found.")
            output = get_empty_output(output.columns.tolist())

        # 📁 Controlled export
        if mode == "full" and not output.isna().all(axis=1).iloc[0]:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_path = f"pension_purpose_missing_{timestamp}.xlsx"
            output.to_excel(file_path, index=False)
            print(f"📁 Saved output to: {file_path}")
        elif mode == "full":
            print("✅ No pension purpose mismatches found. Excel file not created.")

        return output

    except Exception as e:
        print(f"❌ Exception occurred: {e}")
        fallback_columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "Purpose_of_Account",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        return get_empty_output(fallback_columns)
