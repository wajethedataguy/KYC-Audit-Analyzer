import pandas as pd
import re
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="103_ownership_status_for_non_business_flag",
    description="Flags KYC records where ownership status was fed for non-business customers.",
    category="CDD & EDD Review",
)
def logic_103_ownership_status_for_non_business_flag(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    try:
        # 🔍 Load KYC file (always merged_file.xlsx)
        kyc_key = next((k for k in dataframes if "merged_file.xlsx" in k.lower()), None)
        df_kyc = dataframes.get(kyc_key)
        if isinstance(df_kyc, tuple):
            df_kyc = next(iter(df_kyc[0].values())) if df_kyc[0] else pd.DataFrame()
        if df_kyc is None or df_kyc.empty:
            raise ValueError("KYC file not found or empty")

        # 🔧 Normalize column names
        df_kyc.columns = (
            df_kyc.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )

        # 🔧 Column mapping
        col_map = {
            "ACCOUNT_NUMBER": "Account_Number",
            "TITLEOFACCOUNT": "TitleOfAccount",
            "CUSTOMER_NUMBER": "Customer_Number",
            "CUSTOMERFULLNAME": "CustomerFullName",
            "CUSTSECTORCODE": "CustSectorCode",
            "CUSTOMER_OCCUPATION": "Customer_Occupation",
            "STATUSOFOWNERSHIP": "StatusOfOwnership",
            "PRODUCTDESC": "ProductDesc",
            "SECTOR_DESCRIPTION": "Sector_Description",
        }
        df_kyc = df_kyc.rename(
            columns={k: v for k, v in col_map.items() if k in df_kyc.columns}
        )

        required_cols = list(col_map.values())
        missing = [c for c in required_cols if c not in df_kyc.columns]
        if missing:
            raise ValueError(f"Missing required columns in KYC file: {missing}")

        # 🔍 Load NILL_COMBINATIONS reference from uploaded files
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

        def normalize(val):
            return re.sub(r"[^\w]+", "", str(val).strip().lower())

        nill_values = set(map(normalize, raw_nill_values))

        # 🔧 Normalize KYC fields
        df_kyc["Customer_Occupation_Clean"] = (
            df_kyc["Customer_Occupation"].astype(str).map(normalize)
        )
        df_kyc["StatusOfOwnership_Raw"] = (
            df_kyc["StatusOfOwnership"].astype(str).str.strip()
        )
        df_kyc["StatusOfOwnership_Clean"] = df_kyc["StatusOfOwnership_Raw"].map(
            normalize
        )

        # 🧠 Define business occupations
        business_occupations = {"business", "landlord"}

        # 🧠 Apply contradiction logic
        contradiction_mask = (
            ~df_kyc["Customer_Occupation_Clean"].isin(business_occupations)
            & df_kyc["Customer_Occupation_Clean"].notna()
            & (~df_kyc["Customer_Occupation_Clean"].isin(nill_values))
            & df_kyc["StatusOfOwnership_Raw"].notna()
            & (df_kyc["StatusOfOwnership_Raw"] != "")
            & (~df_kyc["StatusOfOwnership_Clean"].isin(nill_values))
        )

        # 📤 Final output
        output_cols = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "CustomerFullName",
            "CustSectorCode",
            "Customer_Occupation",
            "StatusOfOwnership",
            "ProductDesc",
            "Sector_Description",
        ]
        output = (
            df_kyc[contradiction_mask][output_cols]
            .drop_duplicates()
            .reset_index(drop=True)
        )

        # 🧯 Handle empty output
        if output.empty:
            output = pd.DataFrame([{col: pd.NA for col in output_cols}])

        # 📁 Optional export
        if mode == "full":
            file_path = f"ownership_status_for_non_business_flag_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            output.to_excel(file_path, index=False)
        elif mode == "ui":
            output = output[output_cols]

        return output

    except Exception as e:
        print(f"❌ Error in logic_103_ownership_status_for_non_business_flag: {e}")
        return pd.DataFrame(
            [
                {
                    "Customer_Number": pd.NA,
                    "Account_Number": pd.NA,
                    "TitleOfAccount": pd.NA,
                    "CustomerFullName": pd.NA,
                    "CustSectorCode": pd.NA,
                    "Customer_Occupation": pd.NA,
                    "StatusOfOwnership": pd.NA,
                    "ProductDesc": pd.NA,
                    "Sector_Description": pd.NA,
                }
            ]
        )
