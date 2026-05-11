import pandas as pd
import re
from difflib import SequenceMatcher
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="042_missing_name_of_business",
    description="Flags business individuals (sector code 1000, occupation 'Business') whose NameOfBusiness is missing, generic, or similar to TitleOfAccount.",
    category="Compliance & Screening",
)
def logic_042_missing_name_of_business(dataframes: dict, mode="full") -> pd.DataFrame:
    try:
        # 🔍 Identify KYC file
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
            "NAMEOFBUSINESS",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        missing = [col for col in required_cols if col not in df.columns]
        if missing:
            raise ValueError(f"Missing required columns: {missing}")

        # 🔍 Load NILL_COMBINATIONS reference (file-driven)
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

        nill_set = {str(val).strip().upper() for val in raw_nill_values}
        nill_set.update({"NA"})  # safety

        # 🔧 Normalize fields
        df["CUSTSECTORCODE_CLEAN"] = (
            pd.to_numeric(df["CUSTSECTORCODE"], errors="coerce")
            .fillna(0)
            .astype(int)
            .astype(str)
        )
        df["CUSTOMER_OCCUPATION_CLEAN"] = (
            df["CUSTOMER_OCCUPATION"].fillna("").astype(str).str.lower().str.strip()
        )
        df["NAMEOFBUSINESS_CLEAN"] = (
            df["NAMEOFBUSINESS"].fillna("").astype(str).str.strip().str.upper()
        )
        df["TITLEOFACCOUNT_CLEAN"] = (
            df["TITLEOFACCOUNT"].fillna("").astype(str).str.strip().str.upper()
        )
        df["CUSTOMERFULLNAME_CLEAN"] = (
            df["CUSTOMERFULLNAME"].fillna("").astype(str).str.strip().str.upper()
        )

        special_invalid_names = {"MAHEEB ULLAH"}

        def is_invalid_name(row):
            name = row["NAMEOFBUSINESS_CLEAN"]
            title = row["TITLEOFACCOUNT_CLEAN"]
            full_name = row["CUSTOMERFULLNAME_CLEAN"]
            return (
                name == ""
                or name in nill_set
                or name in special_invalid_names
                or name == title
                or name == full_name
                or name.isdigit()
                or len(re.sub(r"[A-Z]", "", name)) == len(name)
            )

        def is_similar(a, b, threshold=0.85):
            a_norm = re.sub(r"[^\w]", "", str(a)).upper()
            b_norm = re.sub(r"[^\w]", "", str(b)).upper()
            if not a_norm or not b_norm:
                return False
            return SequenceMatcher(None, a_norm, b_norm).ratio() >= threshold

        def is_title_extension(row):
            name = row["NAMEOFBUSINESS_CLEAN"]
            title = row["TITLEOFACCOUNT_CLEAN"]
            if not name or not title:
                return False
            name_norm = re.sub(r"[^\w]", "", name)
            title_norm = re.sub(r"[^\w]", "", title)
            if not name_norm.startswith(title_norm):
                return False
            extra_len = len(name_norm) - len(title_norm)
            return extra_len > 0 and extra_len <= 15

        df["IS_SECTOR_1000"] = df["CUSTSECTORCODE_CLEAN"] == "1000"
        df["IS_BUSINESS_OCCUPATION"] = df["CUSTOMER_OCCUPATION_CLEAN"].str.contains(
            "business", case=False, na=False
        )
        df["IS_NAME_INVALID"] = df.apply(is_invalid_name, axis=1)
        df["IS_NAME_SIMILAR_TO_TITLE"] = df.apply(
            lambda row: is_similar(
                row["NAMEOFBUSINESS_CLEAN"], row["TITLEOFACCOUNT_CLEAN"]
            ),
            axis=1,
        )
        df["IS_TITLE_EXTENSION"] = df.apply(is_title_extension, axis=1)

        df_flagged = df[
            df["IS_SECTOR_1000"]
            & df["IS_BUSINESS_OCCUPATION"]
            & (
                df["IS_NAME_INVALID"]
                | (df["IS_NAME_SIMILAR_TO_TITLE"] & ~df["IS_TITLE_EXTENSION"])
            )
        ].copy()

        output = df_flagged[
            [
                "CUSTOMER_NUMBER",
                "ACCOUNT_NUMBER",
                "TITLEOFACCOUNT",
                "CUSTOMERFULLNAME",
                "CUSTSECTORCODE",
                "CUSTOMER_OCCUPATION",
                "NAMEOFBUSINESS",
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
            "NameOfBusiness",
            "ProductDesc",
            "Sector_Description",
        ]
        output = output.reset_index(drop=True)

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
                        "NameOfBusiness": pd.NA,
                        "ProductDesc": pd.NA,
                        "Sector_Description": pd.NA,
                    }
                ]
            )

        if mode == "full":
            file_path = f"missing_name_of_business_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            output.to_excel(file_path, index=False)

        return output

    except Exception:
        return pd.DataFrame(
            [
                {
                    "Customer_Number": pd.NA,
                    "Account_Number": pd.NA,
                    "TitleOfAccount": pd.NA,
                    "CustomerFullName": pd.NA,
                    "CustSectorCode": pd.NA,
                    "Customer_Occupation": pd.NA,
                    "NameOfBusiness": pd.NA,
                    "ProductDesc": pd.NA,
                    "Sector_Description": pd.NA,
                }
            ]
        )
