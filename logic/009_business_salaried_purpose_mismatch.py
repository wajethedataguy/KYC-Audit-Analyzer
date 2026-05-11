import pandas as pd
import unicodedata
import re
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="009_business_salaried_profile_match",
    description=(
        "Surface exceptions where Business/Payroll products do NOT align with Purpose.\n"
        "Scope: Business codes (1150, 6808, 6803) and Payroll codes (1011, 6012, 6809).\n"
        "A) Business product: purpose must be business-like and must NOT be salaried/payroll-like.\n"
        "B) Payroll product: purpose must be salaried/payroll-like and must NOT be business-like."
    ),
    category="Purpose & Occupation Filter",
)
def logic_009_business_salaried_profile_match(
    dataframes: dict, mode: str = "full"
) -> pd.DataFrame:
    """
    Mismatch logic (purpose-based, product-scoped):
    - Business PRODUCTCODE in (1150/6808/6803):
        mismatch if PURPOSE is NOT business-like OR PURPOSE is salaried/payroll-like.
    - Payroll PRODUCTCODE in (1011/6012/6809):
        mismatch if PURPOSE is NOT salaried/payroll-like OR PURPOSE is business-like.
    De-duplicates by Account_Number.
    """
    try:
        # Locate merged file
        kyc_key = next((k for k in dataframes if "merged_file.xlsx" in k.lower()), None)
        if kyc_key is None:
            raise ValueError("Merged file not found in input dataframes.")

        df = dataframes[kyc_key]
        if isinstance(df, tuple):
            df = next(iter(df[0].values())) if df and df[0] else pd.DataFrame()
        df = df.copy()
        if df is None or df.empty:
            raise ValueError("Merged file is empty.")

        # Normalize column names
        df.columns = (
            df.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )

        # Required columns
        required = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
            "PURPOSE_OF_ACCOUNT",
            "PRODUCTCODE",
            "CUSTOMER_OCCUPATION",
        ]
        missing = [c for c in required if c not in df.columns]
        if missing:
            raise ValueError(f"Missing required columns: {missing}")

        # Helpers
        def normalize_text(text) -> str:
            s = "" if pd.isna(text) else str(text)
            s = unicodedata.normalize("NFKD", s)
            s = "".join(ch for ch in s if not unicodedata.combining(ch))
            s = s.upper().strip()
            s = re.sub(r"[\u200B-\u200D\uFEFF]", "", s)  # zero-width chars
            s = re.sub(r"\s+", " ", s)
            return s

        def normalize_product_code(code) -> str:
            if pd.isna(code):
                return ""
            s = str(code).strip()
            if s == "":
                return ""
            try:
                return str(int(float(s)))
            except Exception:
                digits = re.sub(r"\D", "", s)
                return digits if digits else s.upper()

        def to_clean_id(x) -> str:
            s = "" if pd.isna(x) else str(x).strip()
            if s == "":
                return ""
            try:
                return str(int(float(s)))
            except Exception:
                return s

        # Normalize fields
        for col in [
            "PURPOSE_OF_ACCOUNT",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
            "TITLEOFACCOUNT",
            "CUSTOMER_OCCUPATION",
        ]:
            df[col] = df[col].map(normalize_text)

        df["PRODUCTCODE"] = df["PRODUCTCODE"].map(normalize_product_code)
        df["ACCOUNT_NUMBER"] = df["ACCOUNT_NUMBER"].map(to_clean_id)
        df["CUSTOMER_NUMBER"] = df["CUSTOMER_NUMBER"].map(to_clean_id)

        # Strict product codes
        business_codes = {"1150", "6808", "6803"}
        payroll_codes = {"1011", "6012", "6809"}

        in_business_products = df["PRODUCTCODE"].isin(business_codes)
        in_payroll_products = df["PRODUCTCODE"].isin(payroll_codes)

        purpose = df["PURPOSE_OF_ACCOUNT"].fillna("").astype(str)

        # Business-like purpose (FIX: add BUISNES. + BUISNES)
        business_like_pattern = (
            r"(?:"
            r"\bBUSINESS\b|\bBUSNIESS\b|\bBUSIENSS\b|\bBUISNESS\b|\bBUSSINESS\b|"
            r"\bBUISNES\.\b|\bBUISNES\b|"  # ✅ FIX for 'BUISNES.'
            r"\bBUS\S+|BUINESS\b|BUISNES\.|BUSIENSS\.|BUSNIESS\."
            r")"
        )
        has_business_like = purpose.str.contains(business_like_pattern, na=False, regex=True)

        # Salaried/payroll-like purpose (with common misspellings)
        salaried_like_pattern = r"(?:\bSALARY\b|\bSALARIED\b|\bSAL\b|\bPAYROLL\b|\bSALRY\b|\bSALARLY\b|\bSALERY\b)"
        has_salaried_like = purpose.str.contains(salaried_like_pattern, na=False, regex=True)

        # Mismatch logic
        mask_business_mismatch = in_business_products & ((~has_business_like) | has_salaried_like)
        mask_payroll_mismatch = in_payroll_products & ((~has_salaried_like) | has_business_like)

        final_mask = mask_business_mismatch | mask_payroll_mismatch
        df_filtered = df.loc[final_mask].copy()

        # Require non-empty ACCOUNT_NUMBER
        df_filtered = df_filtered[df_filtered["ACCOUNT_NUMBER"].astype(str).str.strip().ne("")]

        # De-duplicate by Account_Number
        df_filtered = df_filtered.sort_values(["ACCOUNT_NUMBER"]).drop_duplicates(subset=["ACCOUNT_NUMBER"])

        # Prepare output
        output_cols = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
            "PURPOSE_OF_ACCOUNT",
            "CUSTOMER_OCCUPATION",
            "PRODUCTCODE",
        ]
        output = df_filtered[output_cols].reset_index(drop=True)
        output.columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "ProductDesc",
            "Sector_Description",
            "Purpose_of_Account",
            "Customer_Occupation",
            "ProductCode",
        ]

        # Optional export
        if mode.lower() == "full" and not output.empty:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_path = f"purpose_mismatch_business_salaried_scoped_{timestamp}.xlsx"
            output.to_excel(file_path, index=False)
            print(f"📁 Saved output to: {file_path}")
        elif mode.lower() == "full":
            print("✅ No matching records found. Excel file not created.")

        return output

    except Exception as e:
        print(f"❌ Exception occurred: {e}")
        fallback = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "ProductDesc",
            "Sector_Description",
            "Purpose_of_Account",
            "Customer_Occupation",
            "ProductCode",
        ]
        return pd.DataFrame([{col: pd.NA for col in fallback}])
