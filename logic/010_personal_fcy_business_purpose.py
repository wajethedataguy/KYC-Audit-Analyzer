import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


def get_empty_output(columns: list) -> pd.DataFrame:
    return pd.DataFrame([{col: pd.NA for col in columns}])


@register_logic(
    name="010_personal_fcy_business_purpose",
    description="Flags personal foreign currency accounts with business-related purpose, which is non-compliant.",
    category="Purpose & Occupation Filter",
)
def logic_010_personal_fcy_business_purpose(
    dataframes: dict, mode: str = "preview"
) -> pd.DataFrame:
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

        required_columns = [
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "CCY",
            "CUSTOMER_NUMBER",
            "PURPOSE_OF_ACCOUNT",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
            "CUSTSECTORCODE",
        ]
        missing = [col for col in required_columns if col not in df.columns]
        if missing:
            raise ValueError(f"Missing required columns: {missing}")

        # 🔧 Normalize relevant fields
        # Normalize CUSTSECTORCODE so 1000.0 → 1000
        df["CUSTSECTORCODE"] = df["CUSTSECTORCODE"].apply(
            lambda x: (
                str(int(float(str(x).strip())))
                if str(x).strip() not in ["", "nan", "None"]
                else ""
            )
        )
        df["CCY"] = df["CCY"].astype(str).str.upper().str.strip()
        df["PURPOSE_OF_ACCOUNT"] = (
            df["PURPOSE_OF_ACCOUNT"].astype(str).str.upper().str.strip()
        )

        # ✅ Strict check: flag any purpose containing "BUS"
        is_business_like = df["PURPOSE_OF_ACCOUNT"].str.contains("BUS", na=False)

        df_filtered = df[
            (df["CUSTSECTORCODE"] == "1000")  # Personal sector
            & (df["CCY"] != "PKR")  # Foreign currency
            & is_business_like  # Purpose contains BUS
        ].copy()

        def clean_number(x):
            try:
                return str(int(float(str(x).strip())))
            except Exception:
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

        # 📤 Final output columns
        output_columns_final = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "CCY",
            "Purpose_of_Account",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]

        output = df_filtered[
            [
                "CUSTOMER_NUMBER",
                "ACCOUNT_NUMBER",
                "TITLEOFACCOUNT",
                "CCY",
                "PURPOSE_OF_ACCOUNT",
                "PRODUCTDESC",
                "SECTOR_DESCRIPTION",
            ]
        ].copy()

        # Rename to match desired casing
        output.columns = output_columns_final
        output = output.reset_index(drop=True)

        if output.empty:
            print("✅ No personal FCY business-purpose mismatches found.")
            output = get_empty_output(output_columns_final)

        # 🧼 Clean pd.NA for Excel export
        output = output.astype(object).where(pd.notna(output), None)

        if mode == "full" and not output.isna().all(axis=1).iloc[0]:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_path = f"personal_fcy_business_purpose_{timestamp}.xlsx"
            output.to_excel(file_path, index=False)
            print(f"📁 Saved output to: {file_path}")
        elif mode == "full":
            print("✅ No mismatches found. Excel file not created.")

        return output

    except Exception as e:
        print(f"❌ Exception occurred: {e}")
        fallback_columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "CCY",
            "Purpose_of_Account",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        return get_empty_output(fallback_columns)
