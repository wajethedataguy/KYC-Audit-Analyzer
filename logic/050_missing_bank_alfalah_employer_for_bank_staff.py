import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="050_missing_bank_alfalah_employer_for_bank_staff",
    description="Flags salaried bank staff whose NameOfEmployer does not mention Bank Alfalah.",
    category="Compliance & Screening",
)
def logic_050_missing_bank_alfalah_employer_for_bank_staff(
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
            "NAMEOFEMPLOYER",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        missing = [col for col in required_cols if col not in df_main.columns]
        if missing:
            raise ValueError(f"Merged file missing required columns: {missing}")

        df = df_main[required_cols].copy()

        # 🔍 Normalize fields
        df["CUSTSECTORCODE"] = pd.to_numeric(df["CUSTSECTORCODE"], errors="coerce")
        df["CUSTOMER_OCCUPATION_CLEAN"] = (
            df["CUSTOMER_OCCUPATION"].astype(str).str.strip().str.upper()
        )
        df["NAMEOFEMPLOYER_CLEAN"] = (
            df["NAMEOFEMPLOYER"].astype(str).str.strip().str.lower()
        )

        # 🧠 Define Bank Alfalah keywords
        bank_keywords = [
            "falah",
            "bafl",
            "bank alfalah",
            "alfalah bank",
            "bank alfalah ltd",
            "bank alfalah limited",
        ]

        # 🔍 Flag rows where employer does NOT mention Bank Alfalah
        df["IS_MISSING_BANK_EMPLOYER"] = df["NAMEOFEMPLOYER_CLEAN"].apply(
            lambda val: not any(kw in val for kw in bank_keywords)
        )
        df["IS_SALARIED"] = df["CUSTOMER_OCCUPATION_CLEAN"] == "SALARIED"
        df["IS_BANK_SECTOR"] = df["CUSTSECTORCODE"] == 1005

        # 🧠 Final filter
        flagged = df[
            df["IS_MISSING_BANK_EMPLOYER"] & df["IS_SALARIED"] & df["IS_BANK_SECTOR"]
        ].copy()

        # 📤 Prepare output (select raw cols, then rename)
        output = flagged[
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
            file_path = f"missing_bank_alfalah_employer_for_bank_staff_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
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
                    "NameOfEmployer": pd.NA,
                    "ProductDesc": pd.NA,
                    "Sector_Description": pd.NA,
                }
            ]
        )
