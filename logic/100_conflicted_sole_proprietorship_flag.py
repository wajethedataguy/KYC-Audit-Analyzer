import pandas as pd
import re
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="100_conflicted_sole_proprietorship_flag",
    description="Flags accounts where multiple individuals claim the same business name as sole proprietorship, violating ownership rules.",
    category="CDD & EDD Review",
)
def logic_100_conflicted_sole_proprietorship_flag(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    try:
        # Load merged KYC file
        kyc_key = next((k for k in dataframes if "merged_file.xlsx" in k.lower()), None)
        df = dataframes.get(kyc_key)
        if isinstance(df, tuple):
            df = next(iter(df[0].values())) if df[0] else pd.DataFrame()
        if df is None or df.empty:
            raise ValueError("merged_file.xlsx not found or empty.")

        # Normalize columns
        df.columns = (
            df.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
            .str.replace(r"[^\w]", "_", regex=True)
        )

        # ----- BASIC CLEANING -----
        df["CUSTOMER_OCCUPATION_CLEAN"] = (
            df["CUSTOMER_OCCUPATION"].fillna("").astype(str).str.strip().str.lower()
        )
        df["CCY_CLEAN"] = df["CCY"].fillna("").astype(str).str.strip().str.upper()

        # Sector normalization
        sector_raw = df["CUSTSECTORCODE"].fillna("").astype(str).str.strip()
        df["CUSTSECTORCODE_CLEAN"] = sector_raw.str.replace(
            r"\.0$", "", regex=True
        ).str.lstrip("0")
        sector = df["CUSTSECTORCODE_CLEAN"]

        # Business name: unify NameOfBusiness and TitleOfAccount
        name_from_business = df["NAMEOFBUSINESS"].fillna("").astype(str).str.strip()
        title_from_account = df["TITLEOFACCOUNT"].fillna("").astype(str).str.strip()
        df["BUSINESS_NAME_RAW"] = name_from_business
        df.loc[df["BUSINESS_NAME_RAW"] == "", "BUSINESS_NAME_RAW"] = title_from_account
        df.loc[sector == "1100", "BUSINESS_NAME_RAW"] = title_from_account
        df["BUSINESS_NAME_CLEAN"] = (
            df["BUSINESS_NAME_RAW"]
            .astype(str)
            .str.lower()
            .str.replace(r"[^\w\s]", "", regex=True)
            .str.replace(r"\s+", " ", regex=True)
            .str.strip()
        )

        # CNIC normalization
        df["CNIC_RAW"] = df["CNIC_NUMBER"].fillna("").astype(str).str.strip()
        df["CNIC_CLEAN"] = (
            df["CNIC_RAW"].str.replace(r"[^\d]", "", regex=True).str.zfill(13)
        )

        # Ownership normalization
        def normalize_ownership(value):
            if pd.isna(value) or str(value).strip() == "":
                return "blank"  # ✅ allow blank values
            text = str(value).lower().strip()
            text = re.sub(r"[^\w\s]", "", text)
            if re.search(r"(own|pro|sole)", text):
                return "sole_proprietor"
            return "none"

        df["OWNERSHIP_CLEAN"] = df["STATUSOFOWNERSHIP"].apply(normalize_ownership)

        # ----- FILTER POPULATION -----
        df_filtered = df[
            sector.isin(["1000", "1100"])
            & (df["CUSTOMER_OCCUPATION_CLEAN"] == "business")
            & (df["CCY_CLEAN"] == "PKR")
            & df["BUSINESS_NAME_CLEAN"].notna()
            & (df["BUSINESS_NAME_CLEAN"] != "")
            & (df["OWNERSHIP_CLEAN"].isin(["sole_proprietor", "blank"]))
        ].copy()

        # ----- CONTRADICTION DETECTION -----
        contradiction_rows = []
        for name, group in df_filtered.groupby("BUSINESS_NAME_CLEAN"):
            if group["CNIC_CLEAN"].nunique() > 1:
                contradiction_rows.append(group)

        if contradiction_rows:
            # ✅ Only drop exact duplicate rows, not CNIC+Business pairs
            df_flagged = pd.concat(contradiction_rows).drop_duplicates()
            df_flagged["NAMEOFBUSINESS"] = (
                df_flagged["BUSINESS_NAME_RAW"].fillna("").astype(str).str.strip()
            )
        else:
            df_flagged = pd.DataFrame()

        # 📤 Final output
        output_cols = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "CNIC_NUMBER",
            "CUSTOMERFULLNAME",
            "NAMEOFBUSINESS",
            "STATUSOFOWNERSHIP",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        output = (
            df_flagged[output_cols].reset_index(drop=True)
            if not df_flagged.empty
            else pd.DataFrame([{col: pd.NA for col in output_cols}])
        )

        # Export / return
        if mode == "full":
            file_path = f"conflicted_sole_proprietorship_flag_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            output.to_excel(file_path, index=False)
        else:
            output = output[output_cols]

        return output

    except Exception as e:
        print(f"❌ Error in logic_100_conflicted_sole_proprietorship_flag: {e}")
        return pd.DataFrame(
            [
                {
                    "Customer_Number": pd.NA,
                    "Account_Number": pd.NA,
                    "TitleOfAccount": pd.NA,
                    "CNIC_Number": pd.NA,
                    "CustomerFullName": pd.NA,
                    "NameOfBusiness": pd.NA,
                    "StatusOfOwnership": pd.NA,
                    "ProductDesc": pd.NA,
                    "Sector_Description": pd.NA,
                }
            ]
        )
