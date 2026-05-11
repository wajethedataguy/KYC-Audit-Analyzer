import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="060_business_turnover_missing_for_business_profiles",
    description="Flags business customers where BusinessTurnover is not a valid integer (excluding 0).",
    category="Compliance & Screening",
)
def logic_060_business_turnover_missing_for_business_profiles(
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
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "CUSTOMER_NUMBER",
            "CUSTOMERFULLNAME",
            "CUSTSECTORCODE",
            "CUSTOMER_OCCUPATION",
            "BUSINESSTURNOVER",
            "SOURCEOFINCOME",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        missing = [col for col in required_cols if col not in df_main.columns]
        if missing:
            raise ValueError(f"Merged file missing required columns: {missing}")

        df = df_main[required_cols].copy()

        # 🔁 Normalize occupation
        df["CUSTOMER_OCCUPATION_NORM"] = (
            df["CUSTOMER_OCCUPATION"].astype(str).str.strip().str.upper()
        )

        # 🔁 Try to convert BusinessTurnover to numeric
        df["BUSINESSTURNOVER_NUM"] = pd.to_numeric(
            df["BUSINESSTURNOVER"], errors="coerce"
        )

        # 🧠 Filter: occupation is BUSINESS and turnover is not a valid integer (NaN) or text, excluding 0
        df_filtered = df[
            (df["CUSTOMER_OCCUPATION_NORM"] == "BUSINESS")
            & (
                df["BUSINESSTURNOVER_NUM"].isna()  # not numeric
                | (df["BUSINESSTURNOVER_NUM"] == 0)  # explicitly 0
            )
        ].copy()

        # 📤 Prepare output
        output = df_filtered[
            [
                "CUSTOMER_NUMBER",
                "ACCOUNT_NUMBER",
                "TITLEOFACCOUNT",
                "PRODUCTDESC",
                "SECTOR_DESCRIPTION",
                "BUSINESSTURNOVER",  # include raw turnover so you can see actual values
            ]
        ].copy()

        output.columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "ProductDesc",
            "SECTOR_DESCRIPTION",
            "BusinessTurnover",
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
                        "ProductDesc": pd.NA,
                        "SECTOR_DESCRIPTION": pd.NA,
                        "BusinessTurnover": pd.NA,
                    }
                ]
            )

        # 📁 Optional export
        if mode == "full":
            file_path = f"business_turnover_missing_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
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
                    "ProductDesc": pd.NA,
                    "SECTOR_DESCRIPTION": pd.NA,
                    "BusinessTurnover": pd.NA,
                }
            ]
        )
