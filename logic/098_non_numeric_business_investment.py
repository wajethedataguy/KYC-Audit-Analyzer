import pandas as pd
import re
import unicodedata
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="098_non_numeric_business_investment_flag",
    description="Flags business customers whose investment field contains non-numeric values in KYC profile.",
    category="CDD & EDD Review",
)
def logic_098_non_numeric_business_investment_flag(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    try:
        # 🔍 Locate merged KYC file
        # file_loader.py already reads with keep_default_na=False,
        # so "N/A" and "NA" arrive here as real strings — use as-is.
        kyc_key = next((k for k in dataframes if "merged_file.xlsx" in k.lower()), None)
        if not kyc_key:
            raise ValueError("KYC profile file not found.")
        df_main = dataframes[kyc_key].copy()

        # 🔧 Normalize column names
        df_main.columns = (
            df_main.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )

        # ✅ Required columns
        required_cols = [
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "CUSTOMER_NUMBER",
            "CUSTOMERFULLNAME",
            "CNIC_NUMBER",
            "CUSTSECTORCODE",
            "CUSTOMER_OCCUPATION",
            "INVESTMENTINBUSINESS",
        ]
        missing = [col for col in required_cols if col not in df_main.columns]
        if missing:
            raise ValueError(f"Merged file missing required columns: {missing}")

        df = df_main[required_cols].copy()

        # 🔍 Load NILL_COMBINATIONS from other uploaded files
        raw_nill_values = set()
        for k, df_nill in dataframes.items():
            if k == kyc_key:
                continue
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
                            sheet["NILL_COMBINATIONS"]
                            .dropna()
                            .astype(str)
                            .str.strip()
                            .str.lower()
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
                        df_nill["NILL_COMBINATIONS"]
                        .dropna()
                        .astype(str)
                        .str.strip()
                        .str.lower()
                    )

        nill_set = {str(v).strip().lower() for v in raw_nill_values}

        # --- normalize tricky strings (NBSP / zero-width chars etc.)
        def norm_str(x) -> str:
            s = "" if pd.isna(x) else str(x)
            s = unicodedata.normalize("NFKC", s)
            s = s.replace("\u00A0", " ")
            s = re.sub(r"[\u200B-\u200D\uFEFF]", "", s)
            s = s.strip().lower()
            s = re.sub(r"\s+", " ", s)
            return s

        # ── Normalize fields ──────────────────────────────────────────────
        df["CUSTSECTORCODE_CLEAN"]      = pd.to_numeric(df["CUSTSECTORCODE"], errors="coerce")
        df["CUSTOMER_OCCUPATION_CLEAN"] = df["CUSTOMER_OCCUPATION"].map(norm_str)
        df["INVESTMENT_STR"]            = df["INVESTMENTINBUSINESS"].map(norm_str)

        # ── Numeric pattern ───────────────────────────────────────────────
        numeric_pattern = re.compile(r"^\-?[\d,]+(\.\d+)?$")

        def is_numeric_value(s: str) -> bool:
            if s in ("", "none", "null"):
                return True
            return bool(numeric_pattern.match(s.replace(",", "")))

        is_numeric = df["INVESTMENT_STR"].map(is_numeric_value)

        # ── Scope filter ──────────────────────────────────────────────────
        in_scope = (
            df["CUSTSECTORCODE_CLEAN"].isin([1000, 1100])
            & df["CUSTOMER_OCCUPATION_CLEAN"].str.contains("business", na=False)
        )

        # ── NILL suppression (N/A and NA always override) ─────────────────
        is_na_token = df["INVESTMENT_STR"].str.match(
            r"^#?n\s*[/\\\.\-]\s*a[\.!]?$", na=False
        )
        is_na_plain = df["INVESTMENT_STR"] == "na"
        in_nill     = df["INVESTMENT_STR"].isin(nill_set)

        # ── Final mask ────────────────────────────────────────────────────
        contradiction_mask = (
            in_scope
            & ~is_numeric
            & (~in_nill | is_na_token | is_na_plain)
        )

        # ── Build output ──────────────────────────────────────────────────
        df["INVESTMENT_RAW"] = df["INVESTMENT_STR"]
        output_cols = required_cols + ["INVESTMENT_RAW"]

        output = (
            df.loc[contradiction_mask, output_cols]
            .drop_duplicates()
            .reset_index(drop=True)
        )

        if output.empty:
            output = pd.DataFrame([{col: pd.NA for col in output_cols}])

        if mode == "full":
            file_path = f"non_numeric_business_investment_flag_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            output.to_excel(file_path, index=False)

        return output

    except Exception as e:
        print(f"❌ Error in logic_098_non_numeric_business_investment_flag: {e}")
        fallback_cols = [
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "CUSTOMER_NUMBER",
            "CUSTOMERFULLNAME",
            "CNIC_NUMBER",
            "CUSTSECTORCODE",
            "CUSTOMER_OCCUPATION",
            "INVESTMENTINBUSINESS",
        ]
        return pd.DataFrame([{col: pd.NA for col in fallback_cols}])