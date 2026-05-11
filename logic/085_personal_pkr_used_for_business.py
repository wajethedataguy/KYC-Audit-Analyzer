import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="085_personal_pkr_used_for_business",
    description="Flags personal PKR accounts used for business/commercial purposes with annual credit turnover exceeding Rs. 450M.",
    category="CDD & EDD Review",
)
def logic_085_personal_pkr_used_for_business(
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
            "CCY",
            "TOTAL_CREDIT_TOTAL",
        ]
        missing = [col for col in required_cols if col not in df_kyc.columns]
        if missing:
            raise ValueError(f"Missing required columns: {missing}")

        # 🔧 Normalize account numbers
        df_kyc["ACCOUNT_NUMBER_RAW"] = df_kyc["ACCOUNT_NUMBER"]
        df_kyc["ACCOUNT_NUMBER"] = (
            pd.to_numeric(df_kyc["ACCOUNT_NUMBER_RAW"], errors="coerce")
            .astype("Int64")
            .astype(str)
        )

        # 🔍 Normalize fields
        df_kyc["CUSTSECTORCODE_CLEAN"] = pd.to_numeric(
            df_kyc["CUSTSECTORCODE"], errors="coerce"
        )
        df_kyc["CCY_CLEAN"] = df_kyc["CCY"].astype(str).str.strip().str.upper()
        df_kyc["CUSTOMER_OCCUPATION_CLEAN"] = (
            df_kyc["CUSTOMER_OCCUPATION"].astype(str).str.strip().str.lower()
        )
        df_kyc["TOTAL_CREDIT_CLEAN"] = pd.to_numeric(
            df_kyc["TOTAL_CREDIT_TOTAL"], errors="coerce"
        )

        # 🧠 Apply contradiction logic
        contradiction_mask = (
            (df_kyc["CUSTSECTORCODE_CLEAN"].isin([1000, 1000.0]))
            & (df_kyc["CCY_CLEAN"] == "PKR")
            & (df_kyc["CUSTOMER_OCCUPATION_CLEAN"] == "business")
            & (df_kyc["TOTAL_CREDIT_CLEAN"] > 450_000_000)
        )

        # 📤 Final output (select only required cols, then rename)
        output_cols = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
            "TOTAL_CREDIT_TOTAL",
        ]
        output = (
            df_kyc[contradiction_mask][output_cols]
            .drop_duplicates()
            .reset_index(drop=True)
        )

        output.columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "ProductDesc",
            "Sector_Description",
            "Total_Credit_Turnover",
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
                        "Total_Credit_Turnover": pd.NA,
                    }
                ]
            )

        # 📁 Optional export
        if mode == "full":
            file_path = f"personal_pkr_misuse_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            output.to_excel(file_path, index=False)
        elif mode == "ui":
            output = output[
                [
                    "Customer_Number",
                    "Account_Number",
                    "TitleOfAccount",
                    "ProductDesc",
                    "Sector_Description",
                    "Total_Credit_Turnover",
                ]
            ]

        return output

    except Exception as e:
        print(f"❌ Error in logic_085_personal_pkr_used_for_business: {e}")
        return pd.DataFrame(
            [
                {
                    "Customer_Number": pd.NA,
                    "Account_Number": pd.NA,
                    "TitleOfAccount": pd.NA,
                    "ProductDesc": pd.NA,
                    "Sector_Description": pd.NA,
                    "Total_Credit_Turnover": pd.NA,
                }
            ]
        )
