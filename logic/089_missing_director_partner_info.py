import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="089_missing_director_partner_info",
    description="Flags corporate customers missing director/partner names or CNICs in KYC profile, violating documentation standards.",
    category="CDD & EDD Review",
)
def logic_089_missing_director_partner_info(
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
            "CUSTSECTORCODE",
            "CUSTOMER_OCCUPATION",
            "PARTNERDIRECTORNAME",
            "PARTNERDIRECTORIDS",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        missing = [col for col in required_cols if col not in df_kyc.columns]
        if missing:
            raise ValueError(f"Missing required columns: {missing}")

        # 🔍 Normalize fields
        df_kyc["CUSTSECTORCODE_CLEAN"] = pd.to_numeric(
            df_kyc["CUSTSECTORCODE"], errors="coerce"
        )
        df_kyc["PARTNERDIRECTORNAME_CLEAN"] = (
            df_kyc["PARTNERDIRECTORNAME"].astype(str).str.strip().str.lower()
        )
        df_kyc["PARTNERDIRECTORIDS_CLEAN"] = (
            df_kyc["PARTNERDIRECTORIDS"].astype(str).str.strip().str.lower()
        )

        # 🧠 Apply contradiction logic:
        # Corporate customers (sector code > 1100) missing either names or CNICs
        contradiction_mask = (df_kyc["CUSTSECTORCODE_CLEAN"] > 1100) & (
            df_kyc["PARTNERDIRECTORNAME_CLEAN"].isna()
            | (df_kyc["PARTNERDIRECTORNAME_CLEAN"] == "")
            | (df_kyc["PARTNERDIRECTORNAME_CLEAN"] == "nan")
            | df_kyc["PARTNERDIRECTORIDS_CLEAN"].isna()
            | (df_kyc["PARTNERDIRECTORIDS_CLEAN"] == "")
            | (df_kyc["PARTNERDIRECTORIDS_CLEAN"] == "nan")
        )

        # 📤 Final output (select raw cols, then rename)
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
            "CustSectorCode",
            "Customer_Occupation",
            "PartnerDirectorName",
            "PartnerDirectorIDs",
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
                        "CustSectorCode": pd.NA,
                        "Customer_Occupation": pd.NA,
                        "PartnerDirectorName": pd.NA,
                        "PartnerDirectorIDs": pd.NA,
                        "ProductDesc": pd.NA,
                        "Sector_Description": pd.NA,
                    }
                ]
            )

        # 📁 Optional export
        if mode == "full":
            file_path = (
                f"missing_director_partner_info_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            )
            output.to_excel(file_path, index=False)
            print(f"✅ Logic 089 output saved at {file_path}")
        elif mode == "ui":
            output = output[
                [
                    "Customer_Number",
                    "Account_Number",
                    "TitleOfAccount",
                    "CustomerFullName",
                    "CNIC_Number",
                    "CustSectorCode",
                    "Customer_Occupation",
                    "PartnerDirectorName",
                    "PartnerDirectorIDs",
                    "ProductDesc",
                    "Sector_Description",
                ]
            ]

        return output

    except Exception as e:
        print(f"❌ Error in logic_089_missing_director_partner_info: {e}")
        return pd.DataFrame(
            [
                {
                    "Customer_Number": pd.NA,
                    "Account_Number": pd.NA,
                    "TitleOfAccount": pd.NA,
                    "CustomerFullName": pd.NA,
                    "CNIC_Number": pd.NA,
                    "CustSectorCode": pd.NA,
                    "Customer_Occupation": pd.NA,
                    "PartnerDirectorName": pd.NA,
                    "PartnerDirectorIDs": pd.NA,
                    "ProductDesc": pd.NA,
                    "Sector_Description": pd.NA,
                }
            ]
        )
