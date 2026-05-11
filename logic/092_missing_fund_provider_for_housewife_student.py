import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="092_missing_fund_provider_for_housewife_student",
    description="Flags personal accounts of housewives/students missing fund provider details in core banking system.",
    category="CDD & EDD Review",
)
def logic_092_missing_fund_provider_for_housewife_student(
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
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
            "CUSTSECTORCODE",
            "CUSTOMER_OCCUPATION",
            "BAF_PEN_ACCT",
            "FUNDS_PROVIDER_ID_NUMBER",
        ]
        missing = [col for col in required_cols if col not in df_kyc.columns]
        if missing:
            raise ValueError(f"Missing required columns: {missing}")

        # 🔧 Normalize fields
        df_kyc["CUSTSECTORCODE_CLEAN"] = pd.to_numeric(
            df_kyc["CUSTSECTORCODE"], errors="coerce"
        )
        df_kyc["CUSTOMER_OCCUPATION_CLEAN"] = (
            df_kyc["CUSTOMER_OCCUPATION"].astype(str).str.strip().str.lower()
        )
        df_kyc["BAF_PEN_ACCT_CLEAN"] = (
            df_kyc["BAF_PEN_ACCT"].astype(str).str.strip().str.upper()
        )
        df_kyc["FUNDS_PROVIDER_ID_NUM_CLEAN"] = (
            df_kyc["FUNDS_PROVIDER_ID_NUMBER"].astype(str).str.strip().str.lower()
        )

        # 🧠 Define target occupations
        target_occupations = {"house wife", "student"}

        # 🧠 Apply contradiction logic
        contradiction_mask = (
            (df_kyc["CUSTSECTORCODE_CLEAN"].isin([1000, 1000.0]))
            & (df_kyc["CUSTOMER_OCCUPATION_CLEAN"].isin(target_occupations))
            & (df_kyc["BAF_PEN_ACCT_CLEAN"] != "Y")
            & (
                df_kyc["FUNDS_PROVIDER_ID_NUM_CLEAN"].isna()
                | (df_kyc["FUNDS_PROVIDER_ID_NUM_CLEAN"] == "")
                | (df_kyc["FUNDS_PROVIDER_ID_NUM_CLEAN"] == "nan")
            )
        )

        # 📤 Final output (select only the desired columns)
        output = (
            df_kyc[contradiction_mask][
                [
                    "CUSTOMER_NUMBER",
                    "ACCOUNT_NUMBER",
                    "TITLEOFACCOUNT",
                    "PRODUCTDESC",
                    "SECTOR_DESCRIPTION",
                ]
            ]
            .drop_duplicates()
            .reset_index(drop=True)
        )

        # Rename columns
        output.columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
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
                        "ProductDesc": pd.NA,
                        "Sector_Description": pd.NA,
                    }
                ]
            )

        # 📁 Optional export
        if mode == "full":
            file_path = f"missing_fund_provider_housewife_student_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            output.to_excel(file_path, index=False)
        elif mode == "ui":
            output = output[
                [
                    "Customer_Number",
                    "Account_Number",
                    "TitleOfAccount",
                    "ProductDesc",
                    "Sector_Description",
                ]
            ]

        return output

    except Exception as e:
        print(f"❌ Error in logic_092_missing_fund_provider_for_housewife_student: {e}")
        return pd.DataFrame(
            [
                {
                    "Customer_Number": pd.NA,
                    "Account_Number": pd.NA,
                    "TitleOfAccount": pd.NA,
                    "ProductDesc": pd.NA,
                    "Sector_Description": pd.NA,
                }
            ]
        )
