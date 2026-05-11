import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="044_missing_nature_of_business",
    description="Flags business customers (sector codes 1000–1122) whose NatureOfBusiness is missing or vague.",
    category="Compliance & Screening",
)
def logic_044_missing_nature_of_business(dataframes: dict, mode="full") -> pd.DataFrame:
    try:
        # 🔍 Identify KYC file
        kyc_key = next((k for k in dataframes if "merged_file.xlsx" in k.lower()), None)
        if not kyc_key:
            raise ValueError("KYC profile file not found.")

        df = dataframes[kyc_key].copy()
        df.columns = (
            df.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )

        required_cols = [
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "CUSTOMER_NUMBER",
            "CUSTOMERFULLNAME",
            "CUSTSECTORCODE",
            "CUSTOMER_OCCUPATION",
            "NATUREOFBUSINESS",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        if missing := [col for col in required_cols if col not in df.columns]:
            raise ValueError(f"Missing required columns: {missing}")

        # 🔍 Load NILL_COMBINATIONS reference (strict, file-driven)
        raw_nill_values = set()
        for df_nill in dataframes.values():
            if isinstance(df_nill, dict):
                for sheet in df_nill.values():
                    if (
                        isinstance(sheet, pd.DataFrame)
                        and "NILL_COMBINATIONS" in sheet.columns
                    ):
                        raw_nill_values.update(
                            sheet["NILL_COMBINATIONS"].dropna().astype(str).str.strip()
                        )
            elif (
                isinstance(df_nill, pd.DataFrame)
                and "NILL_COMBINATIONS" in df_nill.columns
            ):
                raw_nill_values.update(
                    df_nill["NILL_COMBINATIONS"].dropna().astype(str).str.strip()
                )

        nill_set = {str(val).strip().upper() for val in raw_nill_values}

        # 🔧 Normalize fields
        df["CUSTSECTORCODE_CLEAN"] = (
            df["CUSTSECTORCODE"]
            .astype(str)
            .str.extract(r"(\d+)")[0]
            .fillna("")
            .str.strip()
        )
        df["CUSTOMER_OCCUPATION_CLEAN"] = (
            df["CUSTOMER_OCCUPATION"].fillna("").astype(str).str.lower().str.strip()
        )
        df["NATUREOFBUSINESS_CLEAN"] = (
            df["NATUREOFBUSINESS"].fillna("").astype(str).str.upper().str.strip()
        )

        # 🎯 Business segment filter
        allowed_sector_codes = {"1000", "1100", "1110", "1111", "1120", "1121", "1122"}
        df["IS_ALLOWED_SECTOR"] = df["CUSTSECTORCODE_CLEAN"].isin(allowed_sector_codes)
        df["IS_BUSINESS_OCCUPATION"] = df["CUSTOMER_OCCUPATION_CLEAN"] == "business"

        # ❗ Missing / vague NatureOfBusiness
        df["IS_NATURE_MISSING"] = (df["NATUREOFBUSINESS_CLEAN"] == "") | (
            df["NATUREOFBUSINESS_CLEAN"].isin(nill_set)
        )

        # Filter flagged population
        df_flagged = df[
            df["IS_ALLOWED_SECTOR"]
            & df["IS_BUSINESS_OCCUPATION"]
            & df["IS_NATURE_MISSING"]
        ].copy()

        # 📤 Prepare output (Customer_Number always first)
        output = df_flagged[
            [
                "CUSTOMER_NUMBER",
                "ACCOUNT_NUMBER",
                "TITLEOFACCOUNT",
                "CUSTOMERFULLNAME",
                "CUSTSECTORCODE",
                "CUSTOMER_OCCUPATION",
                "NATUREOFBUSINESS",
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
            "NatureOfBusiness",
            "ProductDesc",
            "Sector_Description",
        ]

        # 🔄 Mask NILL_COMBINATIONS in the output
        nature_clean = (
            output["NatureOfBusiness"].fillna("").astype(str).str.upper().str.strip()
        )
        output["NatureOfBusiness"] = nature_clean.where(
            ~nature_clean.isin(nill_set) & (nature_clean != ""), pd.NA
        )

        output = output.reset_index(drop=True)

        # 🧯 Fallback row if no matches
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
                        "NatureOfBusiness": pd.NA,
                        "ProductDesc": pd.NA,
                        "Sector_Description": pd.NA,
                    }
                ]
            )

        # 📁 Optional export
        if mode == "full":
            file_path = (
                f"missing_nature_of_business_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            )
            output.to_excel(file_path, index=False)

        return output

    except Exception:
        return pd.DataFrame(
            [
                {
                    "Customer_Number": pd.NA,
                    "Account_Number": pd.NA,
                    "TitleOfAccount": pd.NA,
                    "CustomerFullName": pd.NA,
                    "CustSectorCode": pd.NA,
                    "Customer_Occupation": pd.NA,
                    "NatureOfBusiness": pd.NA,
                    "ProductDesc": pd.NA,
                    "Sector_Description": pd.NA,
                }
            ]
        )
