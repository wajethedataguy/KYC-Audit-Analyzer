import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="080_resigned_bank_staff_profile_not_updated",
    description="Flags resigned bank staff whose KYC profile still reflects outdated employer or occupation info, restricted to Individuals sector.",
    category="Compliance & Screening",
)
def logic_080_resigned_bank_staff_profile_not_updated(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    try:
        # 🔍 Locate merged KYC file
        kyc_key = next((k for k in dataframes if "merged_file.xlsx" in k.lower()), None)
        if not kyc_key:
            raise ValueError("Merged KYC file not found.")

        df_main = dataframes[kyc_key]
        if isinstance(df_main, tuple):
            df_main = next(iter(df_main[0].values())) if df_main[0] else pd.DataFrame()

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
            "NAMEOFEMPLOYER",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        missing = [col for col in required_cols if col not in df_main.columns]
        if missing:
            raise ValueError(f"Missing required columns: {missing}")

        df = df_main[required_cols].copy()

        # 🔍 Normalize CustSectorCode and filter out current bank staff
        df["CUSTSECTORCODE_CLEAN"] = (
            df["CUSTSECTORCODE"]
            .astype(str)
            .str.replace(".0", "", regex=False)
            .str.strip()
        )
        df = df[df["CUSTSECTORCODE_CLEAN"] != "1005"]

        # 🔧 Normalize employer name for robust matching
        df["EMPLOYER_NORM"] = (
            df["NAMEOFEMPLOYER"]
            .astype(str)
            .str.upper()
            .str.replace(" ", "", regex=False)  # remove spaces
            .str.replace("-", "", regex=False)  # remove hyphens
            .str.replace("_", "", regex=False)  # remove underscores
            .str.strip()
        )

        # 🧠 Apply contradiction logic + restrict to Individuals sector
        contradiction_mask = (df["EMPLOYER_NORM"].str.contains("ALFALAH", na=False)) & (
            df["SECTOR_DESCRIPTION"].astype(str).str.strip().str.upper()
            == "INDIVIDUALS"
        )

        # 📤 Prepare output (select raw cols, then rename)
        output = df[contradiction_mask][
            [
                "CUSTOMER_NUMBER",
                "ACCOUNT_NUMBER",
                "TITLEOFACCOUNT",
                "CUSTOMERFULLNAME",
                "CUSTSECTORCODE",
                "CUSTOMER_OCCUPATION",
                "NAMEOFEMPLOYER",
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
            "NameOfEmployer",
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
                        "NameOfEmployer": pd.NA,
                        "ProductDesc": pd.NA,
                        "Sector_Description": pd.NA,
                    }
                ]
            )

        # 📁 Optional export
        if mode == "full":
            file_path = f"resigned_bank_staff_profile_not_updated_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            output.to_excel(file_path, index=False)
        elif mode == "ui":
            output = output[
                [
                    "Customer_Number",
                    "Account_Number",
                    "TitleOfAccount",
                    "CustomerFullName",
                    "CustSectorCode",
                    "Customer_Occupation",
                    "NameOfEmployer",
                    "ProductDesc",
                    "Sector_Description",
                ]
            ]

        return output

    except Exception as e:
        print(f"❌ Error in logic_080_resigned_bank_staff_profile_not_updated: {e}")
        return pd.DataFrame(
            [
                {
                    "Customer_Number": pd.NA,
                    "Account_Number": pd.NA,
                    "TitleOfAccount": pd.NA,
                    "CustomerFullName": pd.NA,
                    "CustSectorCode": pd.NA,
                    "Customer_Occupation": pd.NA,
                    "NameOfEmployer": pd.NA,
                    "ProductDesc": pd.NA,
                    "Sector_Description": pd.NA,
                }
            ]
        )
