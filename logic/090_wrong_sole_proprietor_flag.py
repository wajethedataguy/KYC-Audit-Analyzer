import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="090_wrong_sole_proprietor_flag",
    description="Flags business individuals wrongly marked as Sole Proprietor in KYC despite being Partner/Owner per FBR record.",
    category="CDD & EDD Review",
)
def logic_090_wrong_sole_proprietor_flag(dataframes: dict, mode="full") -> pd.DataFrame:
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

        # 🔧 Normalize KYC fields
        df_kyc["CUSTSECTORCODE_CLEAN"] = pd.to_numeric(
            df_kyc["CUSTSECTORCODE"], errors="coerce"
        )
        df_kyc["CUSTOMER_OCCUPATION_CLEAN"] = (
            df_kyc["CUSTOMER_OCCUPATION"].astype(str).str.strip().str.lower()
        )
        df_kyc["STATUSOFOWNERSHIP_CLEAN"] = (
            df_kyc["STATUSOFOWNERSHIP"].astype(str).str.strip().str.lower()
        )
        df_kyc["NAMEOFBUSINESS_CLEAN"] = (
            df_kyc["NAMEOFBUSINESS"].astype(str).str.strip()
        )
        df_kyc["CNIC_NUMBER_CLEAN"] = (
            df_kyc["CNIC_NUMBER"].astype(str).str.replace(r"[^\d]", "", regex=True)
        )

        # 🔍 Load FBR ATL Details
        atl_key = next((k for k in dataframes if "fbr" in k.lower()), None)
        if not atl_key:
            raise ValueError("FBR_ATL_Details file not found.")
        df_atl = dataframes.get(atl_key)
        if isinstance(df_atl, tuple):
            df_atl = next(iter(df_atl[0].values())) if df_atl[0] else pd.DataFrame()

        # 🔧 Normalize ATL fields
        df_atl.columns = (
            df_atl.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )
        df_atl["CUSTOMER_CNIC"] = (
            df_atl["CUSTOMER_CNIC"].astype(str).str.replace(r"[^\d]", "", regex=True)
        )
        df_atl["BUSINESS_ASPER_ATL"] = (
            df_atl["BUSINESS_ASPER_ATL"].astype(str).str.strip()
        )

        # 🔗 Join ATL info
        df_joined = pd.merge(
            df_kyc,
            df_atl[["CUSTOMER_CNIC", "BUSINESS_ASPER_ATL"]],
            left_on="CNIC_NUMBER_CLEAN",
            right_on="CUSTOMER_CNIC",
            how="left",
        )

        # 🧠 Apply contradiction logic
        partner_keywords = ["p/o", "partner", "co-owner", "joint owner"]
        df_joined["BUSINESS_ASPER_ATL_LOWER"] = (
            df_joined["BUSINESS_ASPER_ATL"].astype(str).str.lower()
        )

        # Regex match for noisy sole proprietor variants
        df_joined["IS_SOLE_PROPRIETOR"] = df_joined[
            "STATUSOFOWNERSHIP_CLEAN"
        ].str.contains(r"\b(?:prop|sole)\b", case=False, regex=True)

        contradiction_mask = (
            (df_joined["CUSTSECTORCODE_CLEAN"] == 1000)
            & (df_joined["CUSTOMER_OCCUPATION_CLEAN"] == "business")
            & df_joined["IS_SOLE_PROPRIETOR"]
            & (df_joined["NAMEOFBUSINESS_CLEAN"] != "")
            & df_joined["BUSINESS_ASPER_ATL_LOWER"].apply(
                lambda x: any(keyword in x for keyword in partner_keywords)
            )
        )

        # 📤 Final output
        output_cols = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "CNIC_NUMBER",
            "NAMEOFBUSINESS",
            "BUSINESS_ASPER_ATL",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        output = (
            df_joined[contradiction_mask][output_cols]
            .drop_duplicates()
            .sort_values(by="CNIC_NUMBER")
            .reset_index(drop=True)
        )

        # 🧯 Handle empty output
        if output.empty:
            output = pd.DataFrame([{col: pd.NA for col in output_cols}])

        # 📁 Optional export
        if mode == "full":
            file_path = (
                f"wrong_sole_proprietor_flag_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            )
            output.to_excel(file_path, index=False)
        elif mode == "ui":
            output = output[
                [
                    "CUSTOMER_NUMBER",
                    "ACCOUNT_NUMBER",
                    "TITLEOFACCOUNT",
                    "CNIC_NUMBER",
                    "NAMEOFBUSINESS",
                    "BUSINESS_ASPER_ATL",
                    "PRODUCTDESC",
                    "SECTOR_DESCRIPTION",
                ]
            ]

        return output

    except Exception as e:
        print(f"❌ Error in logic_090_wrong_sole_proprietor_flag: {e}")
        return pd.DataFrame(
            [
                {
                    col: pd.NA
                    for col in [
                        "CUSTOMER_NUMBER",
                        "ACCOUNT_NUMBER",
                        "TITLEOFACCOUNT",
                        "CNIC_NUMBER",
                        "NAMEOFBUSINESS",
                        "BUSINESS_ASPER_ATL",
                        "PRODUCTDESC",
                        "SECTOR_DESCRIPTION",
                    ]
                }
            ]
        )
