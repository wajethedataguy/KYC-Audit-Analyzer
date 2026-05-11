import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="068_pep_status_blank_in_kyc_profile",
    description="Flags customers whose ApprovalObtainedForPEP field is blank or null in KYC profile.",
    category="Compliance & Screening",
)
def logic_068_pep_status_blank_in_kyc_profile(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    try:
        # 🔍 Locate merged KYC file
        kyc_key = next((k for k in dataframes if "merged_file.xlsx" in k.lower()), None)
        if not kyc_key:
            raise ValueError("KYC profile file not found.")
        df_main = dataframes[kyc_key].copy()

        # 🔧 Normalize column names
        df_main.columns = (
            df_main.columns.str.strip()
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
            "CUSTSECTORCODE",
            "CUSTOMER_OCCUPATION",
            "APPROVALOBTAINEDFORPEP",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        missing = [col for col in required_cols if col not in df_main.columns]
        if missing:
            raise ValueError(f"Merged file missing required columns: {missing}")

        df = df_main[required_cols].copy()

        # 🔍 Normalize ApprovalObtainedForPEP field
        df["APPROVALOBTAINEDFORPEP_STR"] = (
            df["APPROVALOBTAINEDFORPEP"]
            .astype(str)
            .str.replace(r"[\s\u200b\xa0]+", "", regex=True)
            .str.strip()
        )

        # 🔍 Normalize Account_Number field
        df["ACCOUNT_NUMBER_STR"] = (
            df["ACCOUNT_NUMBER"]
            .fillna("")
            .astype(str)
            .str.replace(r"[\s\u200b\xa0]+", "", regex=True)
            .str.strip()
        )

        # 🧠 Apply contradiction logic
        df_filtered = df[
            (
                df["APPROVALOBTAINEDFORPEP"].isna()
                | (df["APPROVALOBTAINEDFORPEP_STR"] == "")
            )
            & (df["ACCOUNT_NUMBER"].notna())
            & (df["ACCOUNT_NUMBER_STR"] != "")
        ].copy()

        # 📤 Prepare output (select raw cols, then rename)
        output = df_filtered[
            [
                "CUSTOMER_NUMBER",
                "ACCOUNT_NUMBER",
                "TITLEOFACCOUNT",
                "CUSTOMERFULLNAME",
                "CUSTSECTORCODE",
                "CUSTOMER_OCCUPATION",
                "APPROVALOBTAINEDFORPEP",
                "PRODUCTDESC",
                "SECTOR_DESCRIPTION",
            ]
        ].copy()

        output.columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "CustomerFullName",
            "CustSectorCode",
            "Customer_Occupation",
            "ApprovalObtainedForPEP",
            "ProductDesc",
            "Sector_Description",
        ]
        output = output.reset_index(drop=True)

        # 🧯 Handle empty output
        if output.empty:
            output = pd.DataFrame(
                [
                    {
                        "Customer_Number": pd.NA,
                        "Account_Number": pd.NA,
                        "TitleOfAccount": pd.NA,
                        "CustomerFullName": pd.NA,
                        "CustSectorCode": pd.NA,
                        "Customer_Occupation": pd.NA,
                        "ApprovalObtainedForPEP": pd.NA,
                        "ProductDesc": pd.NA,
                        "Sector_Description": pd.NA,
                    }
                ]
            )

        # 📁 Optional export
        if mode == "full":
            file_path = (
                f"pep_status_blank_in_kyc_profile_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            )
            output.to_excel(file_path, index=False)

        return output

    except Exception as e:
        print(f"❌ Exception occurred: {e}")
        return pd.DataFrame(
            [
                {
                    "Customer_Number": pd.NA,
                    "Account_Number": pd.NA,
                    "TitleOfAccount": pd.NA,
                    "CustomerFullName": pd.NA,
                    "CustSectorCode": pd.NA,
                    "Customer_Occupation": pd.NA,
                    "ApprovalObtainedForPEP": pd.NA,
                    "ProductDesc": pd.NA,
                    "Sector_Description": pd.NA,
                }
            ]
        )
