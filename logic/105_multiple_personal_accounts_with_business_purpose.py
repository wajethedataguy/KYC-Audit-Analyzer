import pandas as pd
import re
from datetime import datetime
from KYC_Viewer.utils import register_logic


def normalize(val: str) -> str:
    """Normalize text values: lowercase, strip spaces, remove non-alphanumeric characters."""
    return re.sub(r"[^\w]+", "", str(val).strip().lower())


def _coerce_sector_1000_mask(series: pd.Series) -> pd.Series:
    """Robust match for sector code 1000."""
    s = series.astype(str).str.strip()
    s = s.str.replace(r"[^\d\.\-]", "", regex=True)
    num = pd.to_numeric(s, errors="coerce")
    return num.eq(1000)


@register_logic(
    name="105_multiple_personal_accounts_with_business_purpose",
    description="Flags CNICs with multiple personal accounts used for business purposes, violating compliance instructions.",
    category="CDD & EDD Review",
)
def logic_105_multiple_personal_accounts_with_business_purpose(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    try:
        # Load KYC file
        kyc_key = next((k for k in dataframes if "merged_file.xlsx" in k.lower()), None)
        df_kyc = dataframes.get(kyc_key)
        if isinstance(df_kyc, tuple):
            df_kyc = next(iter(df_kyc[0].values())) if df_kyc[0] else pd.DataFrame()
        if df_kyc is None or df_kyc.empty:
            raise ValueError("KYC file not found or empty.")

        # Normalize column names
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
            "PURPOSE_OF_ACCOUNT",
            "CCY",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        missing = [col for col in required_cols if col not in df_kyc.columns]
        if missing:
            raise ValueError(f"Missing required columns: {missing}")

        # Sector filter: 1000
        mask_1000 = _coerce_sector_1000_mask(df_kyc["CUSTSECTORCODE"])
        df_kyc = df_kyc.loc[mask_1000].copy()

        if df_kyc.empty:
            print("ℹ️ No rows matched CUSTSECTORCODE == 1000 after coercion.")
            cols = [
                "Customer_Number",
                "Account_Number",
                "TitleOfAccount",
                "Customer_Occupation",
                "Purpose_Of_Account",
                "CNIC_Number",
                "ProductDesc",
                "Sector_Description",
            ]
            empty_out = pd.DataFrame([{c: pd.NA for c in cols}])
            if mode == "full":
                file_path = f"multiple_personal_accounts_with_business_purpose_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
                empty_out.to_excel(file_path, index=False)
            return empty_out

        # Normalize key fields
        df_kyc["ACCOUNT_NUMBER"] = (
            pd.to_numeric(df_kyc["ACCOUNT_NUMBER"], errors="coerce")
            .astype("Int64")
            .astype(str)
        )
        df_kyc["CNIC_NUMBER"] = (
            df_kyc["CNIC_NUMBER"].astype(str).str.replace(r"[^\d]", "", regex=True)
        )

        # Normalize text fields
        df_kyc["PURPOSE_CLEAN"] = (
            df_kyc["PURPOSE_OF_ACCOUNT"].astype(str).map(normalize)
        )
        df_kyc["OCCUPATION_CLEAN"] = (
            df_kyc["CUSTOMER_OCCUPATION"].astype(str).map(normalize)
        )

        # Business-purpose proxy
        df_business_purpose = df_kyc[
            df_kyc["CUSTOMER_OCCUPATION"].str.contains("business", case=False, na=False)
            & df_kyc["PURPOSE_OF_ACCOUNT"].str.contains(
                "business|buis|busin|buss", case=False, na=False, regex=True
            )
        ].copy()

        # Group by CNIC and flag those with multiple personal accounts
        flagged_groups = (
            df_business_purpose.groupby("CNIC_NUMBER")
            .filter(lambda g: g["ACCOUNT_NUMBER"].nunique() > 1)
            .copy()
        )

        # Final output (select uppercase cols, then rename)
        output = (
            flagged_groups[
                [
                    "CUSTOMER_NUMBER",
                    "ACCOUNT_NUMBER",
                    "TITLEOFACCOUNT",
                    "CUSTOMER_OCCUPATION",
                    "PURPOSE_OF_ACCOUNT",
                    "CNIC_NUMBER",
                    "PRODUCTDESC",
                    "SECTOR_DESCRIPTION",
                ]
            ]
            .drop_duplicates()
            .reset_index(drop=True)
        )

        output.columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "Customer_Occupation",
            "Purpose_Of_Account",
            "CNIC_Number",
            "ProductDesc",
            "Sector_Description",
        ]

        if output.empty:
            output = pd.DataFrame([{col: pd.NA for col in output.columns}])

        if mode == "full":
            file_path = f"multiple_personal_accounts_with_business_purpose_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            output.to_excel(file_path, index=False)

        return output

    except Exception as e:
        print(
            f"❌ Error in logic_105_multiple_personal_accounts_with_business_purpose: {e}"
        )
        return pd.DataFrame(
            [
                {
                    "Customer_Number": pd.NA,
                    "Account_Number": pd.NA,
                    "TitleOfAccount": pd.NA,
                    "Customer_Occupation": pd.NA,
                    "Purpose_Of_Account": pd.NA,
                    "CNIC_Number": pd.NA,
                    "ProductDesc": pd.NA,
                    "Sector_Description": pd.NA,
                }
            ]
        )
