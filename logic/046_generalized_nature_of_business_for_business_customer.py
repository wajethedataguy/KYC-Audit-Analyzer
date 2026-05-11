import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


def normalize_text(value: str) -> str:
    return (
        str(value)
        .strip()
        .upper()
        .replace(".", "")
        .replace("/", "")
        .replace("&", "AND")
        .replace("-", "")
        .replace("’", "")
        .replace("'", "")
        .replace("\u00a0", "")  # non-breaking space
        .replace(" ", "")
    )


@register_logic(
    name="046_generalized_nature_of_business_for_business_customer",
    description="Flags business customers whose NatureOfBusiness matches vague terms from InExplicitOccu plus expanded normalized terms.",
    category="Compliance & Screening",
)
def logic_046_generalized_nature_of_business_for_business_customer(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    try:
        # Locate merged KYC file
        kyc_key = next((k for k in dataframes if "merged_file.xlsx" in k.lower()), None)
        if not kyc_key:
            raise ValueError("KYC profile file not found.")
        df_main = dataframes[kyc_key].copy()

        # Normalize column names
        df_main.columns = (
            df_main.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )

        # Required columns
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
        missing = [col for col in required_cols if col not in df_main.columns]
        if missing:
            raise ValueError(f"Merged file missing required columns: {missing}")

        df = df_main[required_cols].copy()

        # Force occupation to Business if CustSectorCode >= 1100
        df["CUSTSECTORCODE_NUM"] = pd.to_numeric(df["CUSTSECTORCODE"], errors="coerce")
        df.loc[df["CUSTSECTORCODE_NUM"] >= 1100, "CUSTOMER_OCCUPATION"] = "Business"

        # Normalize fields
        df["CUSTOMER_OCCUPATION_CLEAN"] = (
            df["CUSTOMER_OCCUPATION"].astype(str).str.strip().str.upper()
        )
        df["NATUREOFBUSINESS_CLEAN"] = (
            df["NATUREOFBUSINESS"].astype(str).str.strip().str.upper()
        )
        df["NATUREOFBUSINESS_NORMALIZED"] = (
            df["NATUREOFBUSINESS"].astype(str).apply(normalize_text)
        )

        # Load reference terms (raw)
        ref_terms_raw = set()
        for df_ref in dataframes.values():
            if isinstance(df_ref, dict):
                for sheet_df in df_ref.values():
                    sheet_df.columns = (
                        sheet_df.columns.str.strip()
                        .str.upper()
                        .str.replace(" ", "_")
                        .str.replace("-", "_")
                    )
                    if "IN_EXPLICIT_OCCUPATION" in sheet_df.columns:
                        ref_terms_raw.update(
                            sheet_df["IN_EXPLICIT_OCCUPATION"]
                            .dropna()
                            .astype(str)
                            .str.strip()
                            .str.upper()
                            .unique()
                        )
            elif isinstance(df_ref, pd.DataFrame):
                df_ref.columns = (
                    df_ref.columns.str.strip()
                    .str.upper()
                    .str.replace(" ", "_")
                    .str.replace("-", "_")
                )
                if "IN_EXPLICIT_OCCUPATION" in df_ref.columns:
                    ref_terms_raw.update(
                        df_ref["IN_EXPLICIT_OCCUPATION"]
                        .dropna()
                        .astype(str)
                        .str.strip()
                        .str.upper()
                        .unique()
                    )

        # Add expanded vague/generalized terms explicitly (raw)
        extra_terms_raw = {
            "WHOLESALER",
            "WHOLE SALER",
            "WHOLE SELLER",
            "WHOLESELLER",
            "RETAILER",
            "OTHER",
            "PROCESSING",
            "SERVICE PROVIDER",
            "RATILER",
            "ANY OTHER BUSINESS",
            "MANUFACTURER",
            "TRADER",
            "TRADING",
            "IMPORT AND EXPORT",
            "IMPORT & EXPORT",
            "IMPORT",
            "EXPORT",
            "OTHERS",
            "SHOP KEEPER",
            "WHOLESALLER",
            "WHOLE SALLER",
            "WHOLE SELER",
            "WHOLESELER",
            "IMPORT OR EXPORT",
            "IMPORTS",
            "EXPORTS",
            "IMPORTS AND EXPORTS",
            "IMPORTS OR EXPORTS",
            "IMPORTS & EXPORTS",
            "MANUFACTURING",
            "SELF EMPLOYED",
            "SHOP",
            "SHOWROOM",
            "SHOW ROOM",
        }
        ref_terms_raw.update(extra_terms_raw)

        # Normalize reference terms for robust matching
        ref_terms_normalized = {normalize_text(term) for term in ref_terms_raw}

        # Apply filtering logic
        mask_business = df["CUSTOMER_OCCUPATION_CLEAN"] == "BUSINESS"
        mask_nature_provided = df["NATUREOFBUSINESS"].notna() & (
            df["NATUREOFBUSINESS"].astype(str).str.strip() != ""
        )
        mask_nature_vague = df["NATUREOFBUSINESS_NORMALIZED"].isin(ref_terms_normalized)

        df_flagged = df[mask_business & mask_nature_provided & mask_nature_vague].copy()

        # Prepare output
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
        output = output.reset_index(drop=True)

        # Empty output fallback
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

        # Export if needed
        if mode == "full":
            file_path = f"generalized_nature_of_business_for_business_customer_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
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
                    "CustomerFullName": pd.NA,
                    "CustSectorCode": pd.NA,
                    "Customer_Occupation": pd.NA,
                    "NatureOfBusiness": pd.NA,
                    "ProductDesc": pd.NA,
                    "Sector_Description": pd.NA,
                }
            ]
        )
