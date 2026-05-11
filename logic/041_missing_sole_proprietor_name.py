import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="041_missing_sole_proprietor_name",
    description="Flags sole proprietor customers (sector code 1100, occupation containing 'Business') whose SoleProprietorName is missing or invalid.",
    category="Compliance & Screening",
)
def logic_041_missing_sole_proprietor_name(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    try:
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
            "SOLEPROPRIETORNAME",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        if missing := [col for col in required_cols if col not in df.columns]:
            raise ValueError(f"Missing required columns: {missing}")

        # 🔍 Load NILL_COMBINATIONS for SoleProprietorName checks
        raw_nill_values = set()
        for df_nill in dataframes.values():
            if isinstance(df_nill, dict):
                for sheet in df_nill.values():
                    if not isinstance(sheet, pd.DataFrame):
                        continue
                    sheet.columns = (
                        sheet.columns.str.strip()
                        .str.upper()
                        .str.replace(" ", "_")
                        .str.replace("-", "_")
                    )
                    if "NILL_COMBINATIONS" in sheet.columns:
                        raw_nill_values.update(
                            sheet["NILL_COMBINATIONS"].dropna().astype(str).str.strip()
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
                        df_nill["NILL_COMBINATIONS"].dropna().astype(str).str.strip()
                    )

        nill_set = {str(v).strip().lower() for v in raw_nill_values}

        # Normalize fields
        df["CUSTSECTORCODE_CLEAN"] = (
            df["CUSTSECTORCODE"].astype(str).str.strip().str.lower()
        )
        df["CUSTOMER_OCCUPATION_CLEAN"] = (
            df["CUSTOMER_OCCUPATION"].fillna("").astype(str).str.lower().str.strip()
        )
        df["SOLEPROPRIETORNAME_CLEAN"] = (
            df["SOLEPROPRIETORNAME"].fillna("").astype(str).str.lower().str.strip()
        )
        df["CUSTOMERFULLNAME_CLEAN"] = (
            df["CUSTOMERFULLNAME"].fillna("").astype(str).str.lower().str.strip()
        )
        df["TITLEOFACCOUNT_CLEAN"] = (
            df["TITLEOFACCOUNT"].fillna("").astype(str).str.lower().str.strip()
        )

        generic_names = {
            "owner",
            "proprietor",
            "sole proprietor",
            "n/a",
            "none",
            "not provided",
            "null",
            "na",
            "",
        }

        def is_invalid_name(row):
            val = row["SOLEPROPRIETORNAME_CLEAN"]
            full_name = row["CUSTOMERFULLNAME_CLEAN"]
            title = row["TITLEOFACCOUNT_CLEAN"]

            # Missing/null
            if pd.isna(row["SOLEPROPRIETORNAME"]):
                return True

            # ✅ NEW CONDITION: if SoleProprietorName equals TitleOfAccount or CustomerFullName => INVALID
            # (ignore blanks)
            if val != "" and (val == full_name or val == title):
                return True

            # Existing invalid checks
            if val in nill_set:
                return True
            if val in generic_names:
                return True
            if val.isdigit() or val.isspace():
                return True

            return False

        # ✅ Accept both 1100 and 1100.0
        df["IS_SECTOR_1100"] = df["CUSTSECTORCODE_CLEAN"].isin(["1100", "1100.0"])
        df["IS_BUSINESS_OCCUPATION"] = df["CUSTOMER_OCCUPATION_CLEAN"].str.contains(
            "business", na=False
        )
        df["IS_INVALID_NAME"] = df.apply(is_invalid_name, axis=1)

        df_flagged = df[
            df["IS_SECTOR_1100"] & df["IS_BUSINESS_OCCUPATION"] & df["IS_INVALID_NAME"]
        ].copy()

        output = df_flagged[
            [
                "ACCOUNT_NUMBER",
                "TITLEOFACCOUNT",
                "CUSTOMER_NUMBER",
                "CUSTOMERFULLNAME",
                "CUSTSECTORCODE",
                "CUSTOMER_OCCUPATION",
                "SOLEPROPRIETORNAME",
                "PRODUCTDESC",
                "SECTOR_DESCRIPTION",
            ]
        ].copy()

        output.columns = [
            "Account_Number",
            "TitleOfAccount",
            "Customer_Number",
            "CustomerFullName",
            "CustSectorCode",
            "Customer_Occupation",
            "SoleProprietorName",
            "ProductDesc",
            "Sector_Description",
        ]
        output = output.reset_index(drop=True)

        if output.empty:
            output = pd.DataFrame(
                [
                    {
                        "Account_Number": pd.NA,
                        "TitleOfAccount": pd.NA,
                        "Customer_Number": pd.NA,
                        "CustomerFullName": pd.NA,
                        "CustSectorCode": pd.NA,
                        "Customer_Occupation": pd.NA,
                        "SoleProprietorName": pd.NA,
                        "ProductDesc": pd.NA,
                        "Sector_Description": pd.NA,
                    }
                ]
            )

        if mode == "full":
            file_path = (
                f"missing_sole_proprietor_name_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            )
            output.to_excel(file_path, index=False)

        return output

    except Exception:
        return pd.DataFrame(
            [
                {
                    "Account_Number": pd.NA,
                    "TitleOfAccount": pd.NA,
                    "Customer_Number": pd.NA,
                    "CustomerFullName": pd.NA,
                    "CustSectorCode": pd.NA,
                    "Customer_Occupation": pd.NA,
                    "SoleProprietorName": pd.NA,
                    "ProductDesc": pd.NA,
                    "Sector_Description": pd.NA,
                }
            ]
        )
