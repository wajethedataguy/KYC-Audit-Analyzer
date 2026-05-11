import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


def normalize_name(val: str) -> str:
    """Normalize names for comparison: uppercase, remove M/S, punctuation, spaces."""
    return (
        str(val)
        .upper()
        .replace("M/S", "")
        .replace(",", "")
        .replace(".", "")
        .replace(" ", "")
        .strip()
    )


@register_logic(
    name="081_business_info_missing_for_atl_registered_individuals",
    description="Flags ATL-registered individuals whose KYC profiles lack valid business info in employer field, indicating weak CDD/EDD.",
    category="CDD & EDD Review",
)
def logic_081_business_info_missing_for_atl_registered_individuals(
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
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "CUSTOMER_NUMBER",
            "CUSTOMERFULLNAME",
            "CNIC_NUMBER",
            "CUSTSECTORCODE",
            "CUSTOMER_OCCUPATION",
            "NAMEOFEMPLOYER",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        atl_cols = ["ACCOUNT_NO", "BUSINESS_ASPER_ATL"]
        if any(col not in df_kyc.columns for col in kyc_cols):
            raise ValueError("Missing required columns in KYC file.")
        if any(col not in df_atl.columns for col in atl_cols):
            raise ValueError("Missing required columns in ATL file.")

        df_kyc = df_kyc[kyc_cols].copy()
        df_atl = df_atl[atl_cols].copy()

        # 🔍 Normalize and join on account number
        df_kyc["ACCOUNT_NUMBER"] = df_kyc["ACCOUNT_NUMBER"].astype(str).str.strip()
        df_atl["ACCOUNT_NO"] = df_atl["ACCOUNT_NO"].astype(str).str.strip()
        df_atl["BUSINESS_ASPER_ATL"] = (
            df_atl["BUSINESS_ASPER_ATL"].fillna("").astype(str).str.strip()
        )

        df_kyc_filtered = df_kyc[df_kyc["ACCOUNT_NUMBER"].isin(df_atl["ACCOUNT_NO"])]
        df_merged = pd.merge(
            df_kyc_filtered,
            df_atl,
            left_on="ACCOUNT_NUMBER",
            right_on="ACCOUNT_NO",
            how="left",
        )

        # 🔍 Normalize fields
        df_merged["CUSTSECTORCODE_CLEAN"] = pd.to_numeric(
            df_merged["CUSTSECTORCODE"], errors="coerce"
        )
        df_merged["CUSTOMER_OCCUPATION_CLEAN"] = (
            df_merged["CUSTOMER_OCCUPATION"].astype(str).str.strip().str.upper()
        )
        df_merged["TITLEOFACCOUNT_CLEAN"] = (
            df_merged["TITLEOFACCOUNT"].astype(str).str.strip().str.upper()
        )
        df_merged["BUSINESS_ASPER_ATL_CLEAN"] = (
            df_merged["BUSINESS_ASPER_ATL"].astype(str).str.strip().str.upper()
        )
        df_merged["EMPLOYER_NORM"] = df_merged["NAMEOFEMPLOYER"].apply(normalize_name)
        df_merged["BUSINESS_NORM"] = df_merged["BUSINESS_ASPER_ATL"].apply(
            normalize_name
        )
        df_merged["TITLEOFACCOUNT_NORM"] = df_merged["TITLEOFACCOUNT"].apply(
            normalize_name
        )

        # 🚫 Exclude invalid placeholders only (keep valid occupations like Salaried)
        df_merged["CUSTOMER_OCCUPATION_CLEAN"] = df_merged[
            "CUSTOMER_OCCUPATION_CLEAN"
        ].replace({"0": pd.NA, "NONE": pd.NA, "NAN": pd.NA})
        df_merged = df_merged[df_merged["CUSTOMER_OCCUPATION_CLEAN"].notna()]

        # 🚫 Exclude only specific occupations
        excluded_occupations = {"BUSINESS", "OTHERS", "LANDLORD"}
        df_merged = df_merged[
            ~df_merged["CUSTOMER_OCCUPATION_CLEAN"].isin(excluded_occupations)
        ]

        # 🚫 Drop rows where ATL business info is missing
        df_merged = df_merged[df_merged["BUSINESS_ASPER_ATL_CLEAN"].notna()]
        df_merged = df_merged[df_merged["BUSINESS_ASPER_ATL_CLEAN"] != ""]

        # 🧠 Row-wise check: flag if employer name is not exactly equal to ATL business
        df_merged["BUSINESS_NOT_IN_EMPLOYER"] = df_merged.apply(
            lambda r: pd.notna(r["BUSINESS_NORM"])
            and r["BUSINESS_NORM"] != r["EMPLOYER_NORM"],
            axis=1,
        )

        # 🧠 Apply contradiction logic
        contradiction_mask = (
            (df_merged["CUSTSECTORCODE_CLEAN"] <= 1100)
            & (df_merged["BUSINESS_ASPER_ATL_CLEAN"] != "")
            & (df_merged["BUSINESS_NOT_IN_EMPLOYER"])
            & (df_merged["BUSINESS_NORM"] != df_merged["TITLEOFACCOUNT_NORM"])
        )

        # 📤 Final output
        output_cols = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "CUSTOMER_OCCUPATION",
            "CNIC_NUMBER",
            "NAMEOFEMPLOYER",
            "BUSINESS_ASPER_ATL",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        output = (
            df_merged[contradiction_mask][output_cols]
            .dropna(subset=["BUSINESS_ASPER_ATL"])  # ✅ drop NaN ATL values
            .drop_duplicates()
            .reset_index(drop=True)
        )

        # 🧯 Handle empty output
        if output.empty:
            output = pd.DataFrame([{col: pd.NA for col in output_cols}])

        # 📁 Optional export
        if mode == "full":
            file_path = (
                f"business_info_missing_for_atl_registered_"
                f"{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            )
            output.to_excel(file_path, index=False)
        elif mode == "ui":
            output = output[
                [
                    "CUSTOMER_NUMBER",
                    "ACCOUNT_NUMBER",
                    "TITLEOFACCOUNT",
                    "CUSTOMER_OCCUPATION",
                    "CNIC_NUMBER",
                    "NAMEOFEMPLOYER",
                    "BUSINESS_ASPER_ATL",
                    "PRODUCTDESC",
                    "SECTOR_DESCRIPTION",
                ]
            ]

        return output

    except Exception as e:
        print(
            f"❌ Error in logic_081_business_info_missing_for_atl_registered_individuals: {e}"
        )
        return pd.DataFrame(
            [
                {
                    col: pd.NA
                    for col in [
                        "CUSTOMER_NUMBER",
                        "ACCOUNT_NUMBER",
                        "TITLEOFACCOUNT",
                        "CUSTOMER_OCCUPATION",
                        "CNIC_NUMBER",
                        "NAMEOFEMPLOYER",
                        "BUSINESS_ASPER_ATL",
                        "PRODUCTDESC",
                        "SECTOR_DESCRIPTION",
                    ]
                }
            ]
        )
