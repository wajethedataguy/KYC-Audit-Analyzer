import pandas as pd
import re
from datetime import datetime
from KYC_Viewer.utils import register_logic


def normalize(val: str) -> str:
    """Normalize text values: lowercase, strip spaces, remove non-alphanumeric characters."""
    return re.sub(r"[^\w]+", "", str(val).strip().lower())


@register_logic(
    name="106_missing_incorporation_date_for_business_entities",
    description="Flags business entities where Date of Incorporation is missing or fed as placeholder.",
    category="CDD & EDD Review",
)
def logic_106_missing_incorporation_date_for_business_entities(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    try:
        # 🔍 Load KYC file
        kyc_key = next((k for k in dataframes if "merged_file.xlsx" in k.lower()), None)
        df_kyc = dataframes.get(kyc_key)
        if isinstance(df_kyc, tuple):
            df_kyc = next(iter(df_kyc[0].values())) if df_kyc[0] else pd.DataFrame()
        if df_kyc is None or df_kyc.empty:
            raise ValueError("KYC file not found or empty.")

        # 🔧 Normalize column names
        df_kyc.columns = (
            df_kyc.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )

        required_cols = [
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "CUSTOMER_NUMBER",
            "CNIC_NUMBER",
            "CUSTSECTORCODE",
            "CUSTOMER_OCCUPATION",
            "DATEOFINCORPORATIONOFBUSINESS",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        missing = [col for col in required_cols if col not in df_kyc.columns]
        if missing:
            raise ValueError(f"Missing required columns: {missing}")

        # 🔍 Load NILL_COMBINATIONS reference dynamically
        raw_nill_values = set()
        for df_nill in dataframes.values():
            if isinstance(df_nill, dict):
                for sheet in df_nill.values():
                    sheet.columns = (
                        sheet.columns.str.strip()
                        .str.upper()
                        .str.replace(" ", "_")
                        .str.replace("-", "_")
                    )
                    if "NILL_COMBINATIONS" in sheet.columns:
                        raw_nill_values.update(
                            sheet["NILL_COMBINATIONS"]
                            .dropna()
                            .astype(str)
                            .str.strip()
                            .str.lower()
                        )
            elif isinstance(df_nill, pd.DataFrame):
                df_nill.columns = (
                    df_nill.columns.str.strip()
                    .str.upper()
                    .str.replace(" ", "_")
                    .str.replace("-", "_")
                )
                if "NILL_COMBINATIONS" in df_nill.columns:
                    raw_nill_values.update(
                        df_nill["NILL_COMBINATIONS"]
                        .dropna()
                        .astype(str)
                        .str.strip()
                        .str.lower()
                    )

        nill_values = set(map(normalize, raw_nill_values))

        # 🔧 Normalize key fields
        df_kyc["ACCOUNT_NUMBER"] = (
            pd.to_numeric(df_kyc["ACCOUNT_NUMBER"], errors="coerce")
            .astype("Int64")
            .astype(str)
        )
        df_kyc["CNIC_NUMBER"] = (
            df_kyc["CNIC_NUMBER"].astype(str).str.replace(r"[^\d]", "", regex=True)
        )

        # 🔧 Normalize KYC fields
        df_kyc["DATEOFINCORP_CLEAN"] = (
            df_kyc["DATEOFINCORPORATIONOFBUSINESS"].astype(str).map(normalize)
        )
        df_kyc["CUSTOMER_OCCUPATION_CLEAN"] = (
            df_kyc["CUSTOMER_OCCUPATION"].astype(str).map(normalize)
        )
        df_kyc["CUSTSECTORCODE_NUM"] = pd.to_numeric(
            df_kyc["CUSTSECTORCODE"], errors="coerce"
        )

        # 🧠 Apply contradiction logic
        contradiction_mask = (
            df_kyc["DATEOFINCORP_CLEAN"].isin(nill_values)
            & (df_kyc["CUSTSECTORCODE_NUM"] > 1006)
            & (df_kyc["CUSTOMER_OCCUPATION_CLEAN"] != "others")
        )

        df_flagged = df_kyc[contradiction_mask].copy()

        # 📊 Diagnostics
        print("✅ Total KYC records:", len(df_kyc))
        print("✅ Final contradictions flagged:", len(df_flagged))

        # 📤 Final output
        output_cols = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "CNIC_Number",
            "CustSectorCode",
            "Customer_Occupation",
            "DateOfIncorporationOfBusiness",
            "ProductDesc",
            "Sector_Description",
        ]
        output = (
            df_flagged[
                [
                    "CUSTOMER_NUMBER",
                    "ACCOUNT_NUMBER",
                    "TITLEOFACCOUNT",
                    "CNIC_NUMBER",
                    "CUSTSECTORCODE",
                    "CUSTOMER_OCCUPATION",
                    "DATEOFINCORPORATIONOFBUSINESS",
                    "PRODUCTDESC",
                    "SECTOR_DESCRIPTION",
                ]
            ]
            .drop_duplicates()
            .reset_index(drop=True)
        )

        output.columns = output_cols

        # 🧯 Handle empty output
        if output.empty:
            output = pd.DataFrame([{col: pd.NA for col in output_cols}])

        # 📁 Optional export
        if mode == "full":
            file_path = f"missing_incorporation_date_for_business_entities_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            output.to_excel(file_path, index=False)
        elif mode == "ui":
            output = output[
                [
                    "Account_Number",
                    "TitleOfAccount",
                    "Customer_Number",
                    "CNIC_Number",
                    "CustSectorCode",
                    "Customer_Occupation",
                    "DateOfIncorporationOfBusiness",
                    "ProductDesc",
                    "Sector_Description",
                ]
            ]

        return output

    except Exception as e:
        print(
            f"❌ Error in logic_106_missing_incorporation_date_for_business_entities: {e}"
        )
        return pd.DataFrame(
            [
                {
                    "Account_Number": pd.NA,
                    "TitleOfAccount": pd.NA,
                    "Customer_Number": pd.NA,
                    "CNIC_Number": pd.NA,
                    "CustSectorCode": pd.NA,
                    "Customer_Occupation": pd.NA,
                    "DateOfIncorporationOfBusiness": pd.NA,
                    "ProductDesc": pd.NA,
                    "Sector_Description": pd.NA,
                }
            ]
        )
