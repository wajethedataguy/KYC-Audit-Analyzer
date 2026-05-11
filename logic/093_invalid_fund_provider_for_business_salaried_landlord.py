import pandas as pd
import re
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="093_invalid_fund_provider_for_business_salaried_landlord",
    description="Flags accounts of business/salaried/landlord customers where fund provider details were incorrectly fed.",
    category="CDD & EDD Review",
)
def logic_093_invalid_fund_provider_for_business_salaried_landlord(
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
            "CUSTSECTORCODE",
            "CUSTOMER_OCCUPATION",
            "BAF_PEN_ACCT",
            "FUNDS_PROVIDER_ID_NUMBER",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
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

        # 🔧 Clean CNIC-like values (digits only, handles floats and scientific notation)
        def clean_cnic_like(value):
            try:
                if isinstance(value, float) or re.match(
                    r"^\d+\.?\d*e[\+\-]?\d+$", str(value).lower()
                ):
                    value = int(float(value))
                digits = re.sub(r"[^\d]", "", str(value).strip())
                return digits if digits.isdigit() else ""
            except:
                return ""

        df_kyc["CNIC_NUMBER_CLEAN"] = df_kyc["CNIC_NUMBER"].apply(clean_cnic_like)
        df_kyc["FUNDS_PROVIDER_ID_NUM_CLEAN"] = df_kyc[
            "FUNDS_PROVIDER_ID_NUMBER"
        ].apply(clean_cnic_like)

        # 🧠 Define target occupations and sectors
        target_occupations = {"business", "salaried", "staff", "landlord"}
        target_sectors = {1000, 1005, 1100}

        # 🔍 Apply contradiction logic
        contradiction_mask = (
            df_kyc["CUSTSECTORCODE_CLEAN"].isin(target_sectors)
            & df_kyc["CUSTOMER_OCCUPATION_CLEAN"].isin(target_occupations)
            & (df_kyc["BAF_PEN_ACCT_CLEAN"] != "Y")
            & df_kyc["FUNDS_PROVIDER_ID_NUM_CLEAN"].notna()
            & (df_kyc["FUNDS_PROVIDER_ID_NUM_CLEAN"] != "")
            & (df_kyc["FUNDS_PROVIDER_ID_NUM_CLEAN"] != "nan")
            & (df_kyc["FUNDS_PROVIDER_ID_NUM_CLEAN"] != df_kyc["CNIC_NUMBER_CLEAN"])
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
            "CustSectorCode",
            "Customer_Occupation",
            "BAF_PEN_ACCT",
            "Funds_Provider_ID_Number",
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
                        "CustSectorCode": pd.NA,
                        "Customer_Occupation": pd.NA,
                        "BAF_PEN_ACCT": pd.NA,
                        "Funds_Provider_ID_Number": pd.NA,
                        "ProductDesc": pd.NA,
                        "Sector_Description": pd.NA,
                    }
                ]
            )

        # 📁 Optional export
        if mode == "full":
            file_path = f"invalid_fund_provider_business_salaried_landlord_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
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
                    "CustSectorCode",
                    "Customer_Occupation",
                    "BAF_PEN_ACCT",
                    "Funds_Provider_ID_Number",
                    "ProductDesc",
                    "Sector_Description",
                ]
            ]

        return output

    except Exception as e:
        print(
            f"❌ Error in logic_093_invalid_fund_provider_for_business_salaried_landlord: {e}"
        )
        return pd.DataFrame(
            [
                {
                    "Customer_Number": pd.NA,
                    "Account_Number": pd.NA,
                    "TitleOfAccount": pd.NA,
                    "CustomerFullName": pd.NA,
                    "CNIC_Number": pd.NA,
                    "CustSectorCode": pd.NA,
                    "Customer_Occupation": pd.NA,
                    "BAF_PEN_ACCT": pd.NA,
                    "Funds_Provider_ID_Number": pd.NA,
                    "ProductDesc": pd.NA,
                    "Sector_Description": pd.NA,
                }
            ]
        )
