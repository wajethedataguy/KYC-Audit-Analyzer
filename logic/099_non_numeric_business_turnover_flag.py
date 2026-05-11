import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="099_non_numeric_business_turnover_flag",
    description="Flags business customers whose turnover field contains non-numeric values in KYC profile.",
    category="CDD & EDD Review",
)
def logic_099_non_numeric_business_turnover_flag(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    try:
        # 🔍 Load merged KYC file
        kyc_key = next((k for k in dataframes if "merged_file.xlsx" in k.lower()), None)
        df_kyc = dataframes.get(kyc_key)
        if isinstance(df_kyc, tuple):
            df_kyc = next(iter(df_kyc[0].values())) if df_kyc[0] else pd.DataFrame()
        if df_kyc is None or df_kyc.empty:
            raise ValueError("merged_file.xlsx not found or empty.")

        # 🔧 Normalize column names
        df_kyc.columns = (
            df_kyc.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )

        # ✅ Required columns
        required_cols = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "CUSTOMERFULLNAME",
            "CNIC_NUMBER",
            "CUSTOMER_OCCUPATION",
            "BUSINESSTURNOVER",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        missing = [col for col in required_cols if col not in df_kyc.columns]
        if missing:
            raise ValueError(f"Missing required columns: {missing}")

        # 🔧 Normalize fields
        df_kyc["CUSTOMER_OCCUPATION_CLEAN"] = (
            df_kyc["CUSTOMER_OCCUPATION"].astype(str).str.strip().str.lower()
        )
        df_kyc["BUSINESSTURNOVER_CLEAN"] = (
            df_kyc["BUSINESSTURNOVER"].astype(str).str.strip()
        )

        # 🧠 Define numeric check function
        def is_non_numeric(value):
            try:
                float(str(value).replace(",", "").replace("٫", ".").strip())
                return False
            except:
                return True

        # 🧠 Apply contradiction logic
        contradiction_mask = (
            df_kyc["CUSTOMER_OCCUPATION_CLEAN"].str.contains("business", na=False)
            & df_kyc["BUSINESSTURNOVER_CLEAN"].notna()
            & (df_kyc["BUSINESSTURNOVER_CLEAN"] != "")
            & df_kyc["BUSINESSTURNOVER_CLEAN"].apply(is_non_numeric)
        )

        # 📤 Final output (select raw cols, then rename)
        output = (
            df_kyc[contradiction_mask][required_cols]
            .drop_duplicates()
            .reset_index(drop=True)
        )

        output.columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "CustomerFullName",
            "CNIC_Number",
            "Customer_Occupation",
            "BusinessTurnover",
            "ProductDesc",
            "Sector_Description",
        ]

        # 🧯 Handle empty output
        if output.empty:
            output = pd.DataFrame(
                [
                    {
                        "Customer_Number": pd.NA,
                        "Account_Number": pd.NA,
                        "TitleOfAccount": pd.NA,
                        "CustomerFullName": pd.NA,
                        "CNIC_Number": pd.NA,
                        "Customer_Occupation": pd.NA,
                        "BusinessTurnover": pd.NA,
                        "ProductDesc": pd.NA,
                        "Sector_Description": pd.NA,
                    }
                ]
            )

        # 📁 Optional export
        if mode == "full":
            file_path = f"non_numeric_business_turnover_flag_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            output.to_excel(file_path, index=False)
        elif mode == "ui":
            # Keep all columns for UI mode as well
            output = output[
                [
                    "Customer_Number",
                    "Account_Number",
                    "TitleOfAccount",
                    "CustomerFullName",
                    "CNIC_Number",
                    "Customer_Occupation",
                    "BusinessTurnover",
                    "ProductDesc",
                    "Sector_Description",
                ]
            ]

        return output

    except Exception as e:
        print(f"❌ Error in logic_099_non_numeric_business_turnover_flag: {e}")
        return pd.DataFrame(
            [
                {
                    "Customer_Number": pd.NA,
                    "Account_Number": pd.NA,
                    "TitleOfAccount": pd.NA,
                    "CustomerFullName": pd.NA,
                    "CNIC_Number": pd.NA,
                    "Customer_Occupation": pd.NA,
                    "BusinessTurnover": pd.NA,
                    "ProductDesc": pd.NA,
                    "Sector_Description": pd.NA,
                }
            ]
        )
