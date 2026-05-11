import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="077_blank_kyc_review_date",
    description="Flags customers where KYC review date is blank in their KYC profile.",
    category="Compliance & Screening",
)
def logic_077_blank_kyc_review_date(
    dataframes: dict, mode: str = "full"
) -> pd.DataFrame:
    try:
        # 🔍 Locate merged KYC file
        kyc_key = next((k for k in dataframes if "merged_file.xlsx" in k.lower()), None)
        if not kyc_key:
            raise ValueError("Merged KYC file not found.")
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
            "CNIC_NUMBER",
            "ACCOUNT_OPEN_DT",
            "KYCRISK",
            "KYC_UPDATE_DATE",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        missing = [col for col in required_cols if col not in df_main.columns]
        if missing:
            raise ValueError(f"Missing required columns: {missing}")

        df = df_main[required_cols].copy()

        # 🧠 Apply contradiction logic
        mask = df["ACCOUNT_NUMBER"].notna() & (
            df["KYC_UPDATE_DATE"].isna()
            | (df["KYC_UPDATE_DATE"].astype(str).str.strip() == "")
        )
        df_filtered = df.loc[mask].copy()

        # 📤 Prepare output (select raw cols, then rename)
        output = df_filtered[
            [
                "CUSTOMER_NUMBER",
                "ACCOUNT_NUMBER",
                "TITLEOFACCOUNT",
                "CUSTOMERFULLNAME",
                "CNIC_NUMBER",
                "ACCOUNT_OPEN_DT",
                "KYCRISK",
                "KYC_UPDATE_DATE",
                "PRODUCTDESC",
                "SECTOR_DESCRIPTION",
            ]
        ].copy()

        output.columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "CustomerFullName",
            "CNIC_Number",
            "Account_Open_Date",
            "KYCRisk",
            "KYC_Update_Date",
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
                        "CNIC_Number": pd.NA,
                        "Account_Open_Date": pd.NA,
                        "KYCRisk": pd.NA,
                        "KYC_Update_Date": pd.NA,
                        "ProductDesc": pd.NA,
                        "Sector_Description": pd.NA,
                    }
                ]
            )

        # 📁 Optional export
        if mode == "full" and not output.empty:
            file_path = f"blank_kyc_review_date_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            output.to_excel(file_path, index=False)
        elif mode == "ui" and not output.empty:
            output = output[
                [
                    "Customer_Number",
                    "Account_Number",
                    "TitleOfAccount",
                    "CustomerFullName",
                    "CNIC_Number",
                    "Account_Open_Date",
                    "KYCRisk",
                    "KYC_Update_Date",
                    "ProductDesc",
                    "Sector_Description",
                ]
            ]

        return output

    except Exception as e:
        print(f"❌ Error in logic_077: {e}")
        return pd.DataFrame(
            [
                {
                    "Customer_Number": pd.NA,
                    "Account_Number": pd.NA,
                    "TitleOfAccount": pd.NA,
                    "CustomerFullName": pd.NA,
                    "CNIC_Number": pd.NA,
                    "Account_Open_Date": pd.NA,
                    "KYCRisk": pd.NA,
                    "KYC_Update_Date": pd.NA,
                    "ProductDesc": pd.NA,
                    "Sector_Description": pd.NA,
                }
            ]
        )
