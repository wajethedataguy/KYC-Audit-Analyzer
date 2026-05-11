import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="096_non_numeric_monthly_turnover_flag",
    description="Flags KYC records where non-numeric expected monthly credit turnover was fed in KYC profile.",
    category="CDD & EDD Review",
)
def logic_096_non_numeric_monthly_turnover_flag(
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
            "ACCOUNT_MONTHLY_TURNOVERGT1050M",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        missing = [col for col in required_cols if col not in df_kyc.columns]
        if missing:
            raise ValueError(f"Missing required columns: {missing}")

        # 🔧 Clean numeric fields
        df_kyc["ACCOUNT_MONTHLY_TURNOVERGT1050M_CLEAN"] = pd.to_numeric(
            df_kyc["ACCOUNT_MONTHLY_TURNOVERGT1050M"], errors="coerce"
        )
 
        # 🧠 Apply contradiction logic: flag if any of these fields are non-numeric but not blank
#        contradiction_mask = (
#            (
#                df_kyc["ACCOUNT_MONTHLY_TURNOVERGT1050M_CLEAN"].isna()
#                & df_kyc["ACCOUNT_MONTHLY_TURNOVERGT1050M"].notna()
#            )
#        )
        orig = (
            df_kyc["ACCOUNT_MONTHLY_TURNOVERGT1050M"]
            .astype(str)
            .str.strip()
            .str.lower()
        )

        contradiction_mask = (
            df_kyc["ACCOUNT_MONTHLY_TURNOVERGT1050M_CLEAN"].isna()
            & ~orig.isin({"", "nan", "none", "na", "n/a", "null"})
        )


        # 📤 Final output
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
            "Account_Monthly_TurnoverGT1050M",
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
                        "Account_Monthly_TurnoverGT1050M": pd.NA,
                        "ProductDesc": pd.NA,
                        "Sector_Description": pd.NA,
                    }
                ]
            )

        # 📁 Optional export
        if mode == "full":
            file_path = (
                f"non_numeric_monthly_turnover_flag_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            )
            output.to_excel(file_path, index=False)
        elif mode == "ui":
            output = output[
                [
                    "Customer_Number",
                    "Account_Number",
                    "TitleOfAccount",
                    "CustomerFullName",
                    "CNIC_Number",
                    "Account_Monthly_TurnoverGT1050M",
                    "ProductDesc",
                    "Sector_Description",
                ]
            ]

        return output

    except Exception as e:
        print(f"❌ Error in logic_096_non_numeric_monthly_turnover_flag: {e}")
        return pd.DataFrame(
            [
                {
                    "Customer_Number": pd.NA,
                    "Account_Number": pd.NA,
                    "TitleOfAccount": pd.NA,
                    "CustomerFullName": pd.NA,
                    "CNIC_Number": pd.NA,
                    "Account_Monthly_TurnoverGT1050M": pd.NA,
                    "ProductDesc": pd.NA,
                    "Sector_Description": pd.NA,
                }
            ]
        )
