import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="091_company_account_wrongly_flagged_as_sole_proprietor",
    description="Flags company accounts wrongly categorized as sole proprietorship despite FBR record showing partnership.",
    category="CDD & EDD Review",
)
def logic_091_company_account_wrongly_flagged_as_sole_proprietor(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    try:
        # Helper: robust CNIC cleaner
        def clean_cnic(series: pd.Series) -> pd.Series:
            return (
                series.astype(str)
                .str.strip()
                .str.replace(r"\.0$", "", regex=True)  # remove trailing .0 if float-ish
                .str.replace(r"[^\d]", "", regex=True)  # keep digits only
            )

        # 🔍 Load merged KYC file
        kyc_key = next((k for k in dataframes if "merged_file.xlsx" in k.lower()), None)
        df_kyc = dataframes.get(kyc_key)
        if isinstance(df_kyc, tuple):
            df_kyc = next(iter(df_kyc[0].values())) if df_kyc[0] else pd.DataFrame()
        if df_kyc is None or df_kyc.empty:
            raise ValueError("merged_file.xlsx not found or empty.")

        # 🔧 Normalize KYC column names
        df_kyc.columns = (
            df_kyc.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )

        # 🔧 Normalize KYC fields
        df_kyc["CNIC_NUMBER_CLEAN"] = clean_cnic(df_kyc["CNIC_NUMBER"])
        df_kyc["SECTOR_DESCRIPTION_CLEAN"] = (
            df_kyc["SECTOR_DESCRIPTION"].astype(str).str.strip().str.lower()
        )

        # 🔍 Load FBR ATL Details
        atl_key = next((k for k in dataframes if "fbr" in k.lower()), None)
        if not atl_key:
            raise ValueError("FBR_ATL_Details file not found.")
        df_atl = dataframes.get(atl_key)
        if isinstance(df_atl, tuple):
            df_atl = next(iter(df_atl[0].values())) if df_atl[0] else pd.DataFrame()

        # 🔧 Normalize ATL column names
        df_atl.columns = (
            df_atl.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )

        # 🔧 Normalize ATL fields
        df_atl["CUSTOMER_CNIC"] = clean_cnic(df_atl["CUSTOMER_CNIC"])
        df_atl["DESCRIPTION_CLEAN"] = (
            df_atl["DESCRIPTION"].astype(str).str.strip().str.lower()
        )
        df_atl["BUSINESS_ASPER_ATL_CLEAN"] = (
            df_atl["BUSINESS_ASPER_ATL"].astype(str).str.strip()
        )

        # 🔗 Join ATL info to KYC
        df_joined = pd.merge(
            df_kyc,
            df_atl[["CUSTOMER_CNIC", "DESCRIPTION_CLEAN", "BUSINESS_ASPER_ATL_CLEAN"]],
            left_on="CNIC_NUMBER_CLEAN",
            right_on="CUSTOMER_CNIC",
            how="left",
        )

        # 🧠 Apply contradiction logic
        contradiction_mask = df_joined["SECTOR_DESCRIPTION_CLEAN"].str.contains(
            "sole", na=False
        ) & df_joined["DESCRIPTION_CLEAN"].str.contains("partnership", na=False)

        # 📤 Final output
        output_cols = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "CNIC_NUMBER",
            "SECTOR_DESCRIPTION",
            "DESCRIPTION_CLEAN",
            "BUSINESS_ASPER_ATL_CLEAN",
            "PRODUCTDESC",
        ]
        output = (
            df_joined[contradiction_mask][output_cols]
            .drop_duplicates()
            .sort_values(by="CNIC_NUMBER")
            .reset_index(drop=True)
        )

        # Rename for clarity
        output = output.rename(
            columns={
                "CUSTOMER_NUMBER": "Customer_Number",
                "ACCOUNT_NUMBER": "Account_Number",
                "TITLEOFACCOUNT": "TitleOfAccount",
                "CNIC_NUMBER": "CNIC_Number",
                "SECTOR_DESCRIPTION": "Sector_Description_KYC",
                "DESCRIPTION_CLEAN": "Status_FBR",
                "BUSINESS_ASPER_ATL_CLEAN": "Business_asPer_ATL",
                "PRODUCTDESC": "ProductDesc",
            }
        )

        # 🧯 Handle empty output
        if output.empty:
            output = pd.DataFrame([{col: pd.NA for col in output.columns}])

        # 📁 Optional export
        if mode == "full":
            file_path = (
                f"company_account_wrongly_flagged_as_sole_proprietor_"
                f"{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            )
            output.to_excel(file_path, index=False)
            print(f"✅ Logic 091 output saved at {file_path}")
        elif mode == "ui":
            output = output[
                [
                    "Customer_Number",
                    "Account_Number",
                    "TitleOfAccount",
                    "CNIC_Number",
                    "Sector_Description_KYC",
                    "Status_FBR",
                    "Business_asPer_ATL",
                    "ProductDesc",
                ]
            ]

        return output

    except Exception as e:
        print(
            f"❌ Error in logic_091_company_account_wrongly_flagged_as_sole_proprietor: {e}"
        )
        return pd.DataFrame(
            [
                {
                    "Customer_Number": pd.NA,
                    "Account_Number": pd.NA,
                    "TitleOfAccount": pd.NA,
                    "CNIC_Number": pd.NA,
                    "Sector_Description_KYC": pd.NA,
                    "Status_FBR": pd.NA,
                    "Business_asPer_ATL": pd.NA,
                    "ProductDesc": pd.NA,
                }
            ]
        )
