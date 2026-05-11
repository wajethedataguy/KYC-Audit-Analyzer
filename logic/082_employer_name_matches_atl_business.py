import pandas as pd
from datetime import datetime
from difflib import SequenceMatcher
from KYC_Viewer.utils import register_logic


def normalize_name(val: str) -> str:
    """Normalize names for comparison: uppercase, remove M/S, punctuation, trim."""
    return str(val).upper().replace("M/S", "").replace(",", "").replace(".", "").strip()


def normalize_compact(val: str) -> str:
    """Stronger normalization for substring checks: remove spaces after normalize_name."""
    return normalize_name(val).replace(" ", "")


def tokenize(val: str) -> set:
    """Tokenize normalized text into unique words."""
    return set(normalize_name(val).split())


def similar(a: str, b: str) -> float:
    """Return similarity ratio between two normalized strings."""
    return SequenceMatcher(None, a, b).ratio()


@register_logic(
    name="082_employer_name_matches_atl_business",
    description="Flags salaried individuals whose employer name matches ATL business name, indicating possible concealment.",
    category="CDD & EDD Review",
)
def logic_082_employer_name_matches_atl_business(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    try:
        # 🔍 Locate files
        kyc_key = next((k for k in dataframes if "merged_file.xlsx" in k.lower()), None)
        atl_key = next((k for k in dataframes if "fbr" in k.lower() and "atl" in k.lower()), None)
#        atl_key = next((k for k in dataframes if "fbr_atl" in k.lower()), None)
        if not kyc_key or not atl_key:
            raise ValueError("Required files not found: merged_file.xlsx or FBR_ATL")

        df_kyc = dataframes[kyc_key]
        df_atl = dataframes[atl_key]
        if isinstance(df_kyc, tuple):
            df_kyc = next(iter(df_kyc[0].values())) if df_kyc[0] else pd.DataFrame()
        if isinstance(df_atl, tuple):
            df_atl = next(iter(df_atl[0].values())) if df_atl[0] else pd.DataFrame()

        # 🔧 Normalize column names
        df_kyc.columns = (
            df_kyc.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )
        df_atl.columns = (
            df_atl.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )

        # ✅ Required columns
        kyc_cols = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "CUSTOMERFULLNAME",
            "CNIC_NUMBER",
            "CUSTSECTORCODE",
            "CUSTOMER_OCCUPATION",
            "NAMEOFEMPLOYER",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        atl_cols = ["CUSTOMER_CNIC", "BUSINESS_ASPER_ATL"]
        if any(col not in df_kyc.columns for col in kyc_cols):
            raise ValueError("Missing required columns in KYC file.")
        if any(col not in df_atl.columns for col in atl_cols):
            raise ValueError("Missing required columns in ATL file.")

        df_kyc = df_kyc[kyc_cols].copy()
        df_atl = df_atl[atl_cols].copy()

        # 🔍 Normalize and join
        df_kyc["CNIC_NUMBER"] = df_kyc["CNIC_NUMBER"].astype(str).str.strip()
        df_atl["CUSTOMER_CNIC"] = df_atl["CUSTOMER_CNIC"].astype(str).str.strip()
        df_atl["BUSINESS_ASPER_ATL"] = (
            df_atl["BUSINESS_ASPER_ATL"].fillna("").astype(str).str.strip()
        )

        df_kyc_filtered = df_kyc[df_kyc["CNIC_NUMBER"].isin(df_atl["CUSTOMER_CNIC"])]
        df_merged = pd.merge(
            df_kyc_filtered,
            df_atl,
            left_on="CNIC_NUMBER",
            right_on="CUSTOMER_CNIC",
            how="left",
        )

        # 🔍 Normalize fields
        df_merged["CUSTSECTORCODE_CLEAN"] = (
            df_merged["CUSTSECTORCODE"]
            .astype(str)
            .str.replace(".0", "", regex=False)
            .str.strip()
        )
        df_merged["EMPLOYER_NORM"] = df_merged["NAMEOFEMPLOYER"].apply(normalize_name)
        df_merged["BUSINESS_NORM"] = df_merged["BUSINESS_ASPER_ATL"].apply(
            normalize_name
        )
        df_merged["EMPLOYER_COMPACT"] = df_merged["NAMEOFEMPLOYER"].apply(
            normalize_compact
        )
        df_merged["BUSINESS_COMPACT"] = df_merged["BUSINESS_ASPER_ATL"].apply(
            normalize_compact
        )
        df_merged["EMPLOYER_TOKENS"] = df_merged["NAMEOFEMPLOYER"].apply(tokenize)
        df_merged["BUSINESS_TOKENS"] = df_merged["BUSINESS_ASPER_ATL"].apply(tokenize)

        # 🧠 Hybrid matching: fuzzy + substring + token overlap
        THRESH = 0.85

        def hybrid_match(row) -> bool:
            emp = row["EMPLOYER_NORM"]
            biz = row["BUSINESS_NORM"]
            emp_c = row["EMPLOYER_COMPACT"]
            biz_c = row["BUSINESS_COMPACT"]
            emp_t = row["EMPLOYER_TOKENS"]
            biz_t = row["BUSINESS_TOKENS"]

            if not emp or not biz:
                return False

            if similar(emp, biz) >= THRESH:
                return True
            if emp_c in biz_c or biz_c in emp_c:
                return True
            if emp_t and biz_t:
                overlap = len(emp_t & biz_t)
                if overlap / max(len(emp_t), 1) >= 0.8:
                    return True
            return False

        df_merged["EMPLOYER_MATCHES_FBR"] = df_merged.apply(hybrid_match, axis=1)

        # 🎯 Target salaried individuals (occupation = Salaried) with employer-business match
        contradiction_mask = (
            df_merged["CUSTOMER_OCCUPATION"].astype(str).str.strip().str.upper()
            == "SALARIED"
        ) & (df_merged["EMPLOYER_MATCHES_FBR"])

        df_merged["CONTRADICTION_REASON"] = ""
        df_merged.loc[contradiction_mask, "CONTRADICTION_REASON"] = (
            "Employer name in KYC matches ATL business name (hybrid match) — possible concealment"
        )

        # 📤 Final output
        output_cols = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "CUSTOMERFULLNAME",
            "CNIC_NUMBER",
            "CUSTSECTORCODE",
            "CUSTOMER_OCCUPATION",
            "NAMEOFEMPLOYER",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
            "BUSINESS_ASPER_ATL",
            "CONTRADICTION_REASON",
        ]
        output = (
            df_merged[contradiction_mask][output_cols]
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
            "NameOfEmployer",
            "ProductDesc",
            "Sector_Description",
            "Business_AsPer_ATL",
            "Contradiction_Reason",
        ]

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
                        "NameOfEmployer": pd.NA,
                        "ProductDesc": pd.NA,
                        "Sector_Description": pd.NA,
                        "Business_AsPer_ATL": pd.NA,
                        "Contradiction_Reason": pd.NA,
                    }
                ]
            )

        if mode == "full":
            file_path = f"employer_name_matches_atl_business_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            output.to_excel(file_path, index=False)
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
                    "NameOfEmployer",
                    "Business_AsPer_ATL",
                    "ProductDesc",
                    "Sector_Description",
                ]
            ]

        return output

    except Exception as e:
        print(f"❌ Error in logic_082_employer_name_matches_atl_business: {e}")
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
                    "NameOfEmployer": pd.NA,
                    "Business_AsPer_ATL": pd.NA,
                    "ProductDesc": pd.NA,
                    "Sector_Description": pd.NA,
                }
            ]
        )
