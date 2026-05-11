import os
import re
import numpy as np
import pandas as pd
import xlsxwriter
import csv
import warnings

warnings.filterwarnings("ignore", category=FutureWarning)


# Part-1#
# ─────────────────────────────────────────────────────────
# 🔍 Turnover Mapping (Power Query-aligned)
# ─────────────────────────────────────────────────────────
turnover_map = {
    "Below 1M": 0,
    "1M to 5M": 1_000_000,
    "5M to 10M": 5_000_000,
    "Below 10M": 0,
    "10M to 50M": 10_000_000,
}


def convert_turnover(value, fallback=None):
    """
    Convert text buckets like '1M to 5M' into numeric.
    Returns numeric value or np.nan if cannot parse.
    """
    s = "" if pd.isna(value) else str(value).strip()
    mapped = turnover_map.get(s)
    if mapped is not None:
        return mapped

    # fallback numeric
    if fallback is not None and not pd.isna(fallback):
        try:
            return float(fallback)
        except Exception:
            return np.nan

    # try direct numeric cast
    try:
        return float(re.sub(r"[^\d\.]", "", s)) if s else np.nan
    except Exception:
        return np.nan


# Part-2#
# ─────────────────────────────────────────────────────────
# 🧼 Column Cleaning
# ─────────────────────────────────────────────────────────
def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Strip and underscore column names; safely trim string cells only when appropriate.
    Keeps numeric / boolean / nullable types intact to avoid .str accessor errors.
    Special handling: skip IsMatched so it remains raw (cleaned later in merge step).
    """
    df = df.copy()

    # Standardize column names
    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.replace(" ", "_", regex=False)
        .str.replace("-", "_", regex=False)
    )

    # For columns that are object / string / categorical, convert to pandas "string"
    # dtype then strip and normalise literal "nan"/"None" into <NA>.
    for col in df.columns:
        ser = df[col]

        # 🚫 Special case: leave IsMatched untouched
        if col == "IsMatched":
            df[col] = ser
            continue

        # 🚫 Preserve raw text columns — do not cast to StringDtype as pandas
        # may silently convert "N/A" / "NA" to pd.NA in that dtype.
        if col == "InvestmentInBusiness":
            df[col] = ser
            continue

        if (
            pd.api.types.is_object_dtype(ser)
            or pd.api.types.is_string_dtype(ser)
            or pd.api.types.is_categorical_dtype(ser)
        ):
            try:
                s = ser.astype("string").str.strip()
                # Normalize textual placeholders to missing
                s = s.replace({"nan": pd.NA, "NaN": pd.NA, "None": pd.NA})
                df[col] = s
            except Exception:
                # If conversion unexpectedly fails, leave original column untouched
                df[col] = ser
        else:
            # numeric / boolean / datetime columns: do not change dtype here
            df[col] = ser

    return df


def export_to_excel(df: pd.DataFrame, output_path: str, id_columns: list = None):
    if id_columns is None:
        id_columns = []

    # Ensure all ID columns are stringified
    for col in id_columns:
        if col in df.columns:
            df[col] = df[col].astype(str)

    # Use ExcelWriter with formatting
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Merged")
        workbook = writer.book
        worksheet = writer.sheets["Merged"]

        # Apply text format to ID columns
        text_format = workbook.add_format({"num_format": "@", "align": "left"})
        for idx, col in enumerate(df.columns):
            if col in id_columns:
                worksheet.set_column(idx, idx, 25, text_format)


# Part-3#
# ─────────────────────────────────────────────────────────
# 🔍 Role Detection via Column Signatures (no filename reliance)
# ─────────────────────────────────────────────────────────
KYC_SIGNATURE = {"Account_Num", "KYC_Annual_Turnover", "Cust_Sector_Code"}
CRP_SIGNATURE = {
    "Account_No2",
    "A_C_Turnover_P_A___in_Amount__for_Individual_Trusts_Societies2",
}


def detect_file_roles(files_data):
    """
    Scan uploaded files (sheets) and return (kyc_df, crp_df, base_path).
    This version is more defensive:
      - cleans columns
      - tries primary signatures, then fallback heuristics
      - auto-swaps if roles appear reversed
      - provides helpful debug prints
    """
    kyc_df = None
    crp_df = None
    base_path = None

    for filename, (sheets_dict, path) in files_data.items():
        if base_path is None and isinstance(path, str) and len(path) > 0:
            base_path = os.path.dirname(os.path.abspath(path))
        for sheet_name, df in sheets_dict.items():
            df = clean_columns(df)
            cols = set(df.columns)

            # primary detect
            if kyc_df is None and KYC_SIGNATURE.issubset(cols):
                kyc_df = df
            if crp_df is None and CRP_SIGNATURE.issubset(cols):
                crp_df = df

    # Fallback heuristics if primary detection didn't find both
    if kyc_df is None or crp_df is None:
        for filename, (sheets_dict, path) in files_data.items():
            for sheet_name, df in sheets_dict.items():
                df = clean_columns(df)
                cols = set(df.columns)
                # Heuristic: if dataframe contains Account_Num and Customer_Num it's probably KYC
                if kyc_df is None and {"Account_Num", "Customer_Num"}.issubset(cols):
                    kyc_df = df
                # Heuristic: if dataframe contains Account_No2 and Customer_ID2 it's probably CRP
                if crp_df is None and {"Account_No2", "Customer_ID2"}.issubset(cols):
                    crp_df = df
                if kyc_df is not None and crp_df is not None:
                    break
            if kyc_df is not None and crp_df is not None:
                break

    # Final check: if still missing, raise with diagnostics
    if kyc_df is None or crp_df is None:
        kyc_cols = list(kyc_df.columns) if kyc_df is not None else []
        crp_cols = list(crp_df.columns) if crp_df is not None else []
        raise ValueError(
            "❌ Could not detect both KYC and CRP sheets based on signatures/heuristics.\n"
            f"KYC sample cols: {kyc_cols[:20]}\nCRP sample cols: {crp_cols[:20]}"
        )

    # Ensure roles are not swapped: KYC should have Account_Num, CRP should have Account_No2
    # If swapped, swap them back.
    kyc_cols = set(kyc_df.columns)
    crp_cols = set(crp_df.columns)
    swapped = False
    if "Account_Num" not in kyc_cols and "Account_Num" in crp_cols:
        if "Account_No2" not in crp_cols and "Account_No2" in kyc_cols:
            # both ambiguous; leave as-is
            pass
        else:
            # swap if clearly reversed
            kyc_df, crp_df = crp_df, kyc_df
            swapped = True

    if swapped:
        print(
            "⚠️ Auto-swap performed: detected CRP/KYC roles were reversed. Fixed automatically."
        )

    return kyc_df, crp_df, base_path


# Part-4#
# ─────────────────────────────────────────────────────────
# 🔢 Turnover Normalization Helpers (optional overall view)
# ─────────────────────────────────────────────────────────
def normalize_turnover_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Extract generic Account_Turnover* columns if present (for diagnostics/consistency).
    Converts numeric columns safely even with strings like '4.22019E+12'.
    """
    df = df.copy()

    def extract_turnover(row):
        for col in [
            "A_C_Turnover_P_A____in_Amount__for_Corporate",
            "A_C_Turnover_P_A___in_Amount__for_Individual_Trusts_Societies2",
            "Business_Turnover2",
            "KYC_Annual_Turnover",
        ]:
            val = row.get(col)
            if pd.notna(val) and str(val).strip() != "":
                return val
        return np.nan

    def extract_turnover_narrative(row):
        for col in [
            "A_C_Turnover_P_A___in_Range__for_Corporate",
            "A_C_Turnover_P_A___in_Range__for_Individual_Trusts_Societies2",
        ]:
            val = row.get(col)
            if pd.notna(val) and str(val).strip() != "":
                return val
        return np.nan

    df["Account_Turnover_in_Numbers"] = df.apply(extract_turnover, axis=1)
    df["Account_Turnover"] = df.apply(extract_turnover_narrative, axis=1)

    # Strip commas, spaces, and safely convert to numeric
    df["Account_Turnover_in_Numbers_clean"] = pd.to_numeric(
        df["Account_Turnover_in_Numbers"]
        .astype(str)
        .str.replace(r"[,\s]", "", regex=True),
        errors="coerce",
    )
    return df


# Part-5#
# ─────────────────────────────────────────────────────────
# 📅 Date Normalization (Power Query-aligned)
# ─────────────────────────────────────────────────────────
def normalize_date_columns(df: pd.DataFrame, format_map: dict = None) -> pd.DataFrame:
    """
    Normalize likely date columns in KYC/CRP DataFrames.
    Adds drift flags and parse status overlays for audit traceability.
    Ensures day-first parsing and fallback formats to align with Power Query behavior.
    """
    df = df.copy()
    if format_map is None:
        format_map = {}

    skip_cols = {
        "MANDATEE_NAME",
        "MANDATEE_NIC",
        "Ac_Last_Modif_Date",
        "DATE_LAST_CR_AUTO",
        "DATE_LAST_DR_AUTO",
        "ID_Val_Date",
        "DateOfIncorporationOfBusiness",
    }

    fallback_formats = ["%Y%m%d", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y"]

    def is_compound(val):
        return "?" in str(val) or len(str(val)) > 12

    def try_parse(val, primary_fmt=None):
        val = str(val).strip()
        if val in ["", "NaT", "nan", "None"]:
            return pd.NaT, "failed"

        # Detect compound or malformed values
        if "?" in val or len(val) > 12:
            return val, "skipped"

        try:
            if primary_fmt:
                parsed = pd.to_datetime(val, format=primary_fmt, errors="coerce")
                if pd.notna(parsed):
                    return parsed, "parsed"
            else:
                parsed = pd.to_datetime(val, dayfirst=True, errors="coerce")
                if pd.notna(parsed):
                    return parsed, "parsed"
        except:
            pass

        # Fallback formats
        fallback_formats = ["%Y%m%d", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y"]
        for fmt in fallback_formats:
            try:
                parsed = pd.to_datetime(val, format=fmt, errors="coerce")
                if pd.notna(parsed):
                    return parsed, "fallback"
            except:
                continue

        return pd.NaT, "failed"

    date_cols = [c for c in df.columns if "date" in c.lower() and c not in skip_cols]

    for col in date_cols:
        raw_series = df[col].astype(str).str.strip().replace({"nan": "", "NaT": ""})
        fmt = format_map.get(col)
        parsed = []
        drift_flag = []
        status = []

        for val in raw_series:
            if is_compound(val):
                parsed.append(val)
                drift_flag.append(True)
                status.append("skipped")
                continue

            parsed_val = try_parse(val, fmt)
            parsed.append(parsed_val)
            drift_flag.append(str(parsed_val) != val)
            status.append("parsed" if parsed_val != pd.NaT else "failed")

        df[col] = parsed
        df[f"{col}_Drift_Flag"] = drift_flag
        df[f"{col}_ParseStatus"] = status

    return df


# Part-6#
# ─────────────────────────────────────────────────────────
# 🧠 KYC Turnover Logic (Power Query aligned)
# ─────────────────────────────────────────────────────────
def apply_kyc_turnover_logic(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["Cust_Sector_Code"] = pd.to_numeric(
        df.get("Cust_Sector_Code", np.nan), errors="coerce"
    )

    def safe_numeric(value):
        """Convert anything numeric-like to float, else np.nan"""
        if pd.isna(value):
            return np.nan
        try:
            return float(str(value).replace(",", "").strip())
        except Exception:
            return np.nan

    def annual(row):
        base = (
            row.get("KYC_Annual_Turnover")
            if pd.notna(row["Cust_Sector_Code"]) and row["Cust_Sector_Code"] <= 1100
            else row.get("KYC_Ann_TO_Corporate")
        )
        fallback = (
            row.get("TO_Greater_Than_10M")
            if pd.notna(row["Cust_Sector_Code"]) and row["Cust_Sector_Code"] <= 1100
            else row.get("TO_Greater_Than_50M")
        )
        return convert_turnover(base, safe_numeric(fallback))

    def monthly(row):
        base = (
            row.get("MONTH_TOVER_RG")
            if pd.notna(row["Cust_Sector_Code"]) and row["Cust_Sector_Code"] <= 1100
            else row.get("MON_TOVER_CRG")
        )
        fallback = (
            row.get("EXP_MONTH_TOVER")
            if pd.notna(row["Cust_Sector_Code"]) and row["Cust_Sector_Code"] <= 1100
            else row.get("MON_TOVER_CORP")
        )
        return convert_turnover(base, safe_numeric(fallback))

    # Store base selections (traceability)
    df["A003KYC_BaseValue"] = df.apply(
        lambda row: (
            row.get("KYC_Annual_Turnover")
            if pd.notna(row["Cust_Sector_Code"]) and row["Cust_Sector_Code"] <= 1100
            else row.get("KYC_Ann_TO_Corporate")
        ),
        axis=1,
    )
    df["TurnoverGT1050M"] = df.apply(
        lambda row: (
            row.get("TO_Greater_Than_10M")
            if pd.notna(row["Cust_Sector_Code"]) and row["Cust_Sector_Code"] <= 1100
            else row.get("TO_Greater_Than_50M")
        ),
        axis=1,
    )

    # Numeric annual & monthly turnover
    df["A003_KYC_Turnover"] = df.apply(annual, axis=1)
    df["A003MonthTurnover"] = df.apply(
        lambda row: (
            row.get("MONTH_TOVER_RG")
            if pd.notna(row["Cust_Sector_Code"]) and row["Cust_Sector_Code"] <= 1100
            else row.get("MON_TOVER_CRG")
        ),
        axis=1,
    )
    df["KYCMonthlyTurnoverGT1050M"] = df.apply(
        lambda row: (
            row.get("EXP_MONTH_TOVER")
            if pd.notna(row["Cust_Sector_Code"]) and row["Cust_Sector_Code"] <= 1100
            else row.get("MON_TOVER_CORP")
        ),
        axis=1,
    )
    df["A003_Monthly_Credit_Turnover"] = df.apply(monthly, axis=1)

    return df


# Part-7#
# ─────────────────────────────────────────────────────────
# 🧠 CRP Turnover Logic (Power Query aligned)
# ─────────────────────────────────────────────────────────
def apply_crp_turnover_logic(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    def safe_numeric(value):
        """Convert anything numeric-like to float, else np.nan"""
        if pd.isna(value):
            return np.nan
        try:
            return float(str(value).replace(",", "").strip())
        except Exception:
            return np.nan

    # base text bucket & numeric override
    df["CRP_BaseValue"] = df.get(
        "A_C_Turnover_P_A___in_Range__for_Individual_Trusts_Societies2", np.nan
    )
    df["CRPTurnoverGT1050M"] = df.get(
        "A_C_Turnover_P_A___in_Amount__for_Individual_Trusts_Societies2", np.nan
    )

    df["CRP_Turnover"] = df.apply(
        lambda row: convert_turnover(
            row.get("CRP_BaseValue"), safe_numeric(row.get("CRPTurnoverGT1050M"))
        ),
        axis=1,
    )

    # monthly versions
    df["CRP_MonthlyTurnover"] = df.get(
        "Expected_Monthly_Turnover__In_Range____for_Individual_Trusts_Societies2",
        np.nan,
    )
    df["CRPMonthlyTurnoverGT1050M"] = df.get(
        "Expected_Monthly_Turnover__In_Amount____for_Individual_Trusts_Societies2",
        np.nan,
    )

    df["CRP_MonthlyTurnover_Number"] = df.apply(
        lambda row: convert_turnover(
            row.get("CRP_MonthlyTurnover"),
            safe_numeric(row.get("CRPMonthlyTurnoverGT1050M")),
        ),
        axis=1,
    )
    return df


# Part-8#

# ─────────────────────────────────────────────────────────
# 🔧 Utility Functions
# ─────────────────────────────────────────────────────────


def format_id_column(series):
    """Safely format numeric-like IDs to string, avoiding scientific notation and preserving alphanumeric values."""

    def safe_format(x):
        try:
            return str(int(float(x))) if pd.notnull(x) and str(x).strip() != "" else ""
        except:
            return str(x).strip() if pd.notnull(x) else ""

    return series.apply(safe_format)


# ─────────────────────────────────────────────────────────
# 🔧 Safe CNIC Normalization (Handles Mixed Types)
# ─────────────────────────────────────────────────────────


def normalize_cnic(series: pd.Series) -> pd.Series:
    """Safely normalize CNIC values to 13-digit numeric strings."""
    series = series.copy()
    # convert only stringifiable values
    series = series.astype("string")
    series = (
        series.str.replace(r"[^0-9]", "", regex=True)
        .str.strip()
        .str.zfill(13)
        .where(series.notna(), np.nan)
    )
    return series


# ─────────────────────────────────────────────────────────
# 🔗 Merge Logic with Smart Deduplication (CRP prioritized)
# ─────────────────────────────────────────────────────────


def merge_datasets(
    crp: pd.DataFrame, kyc: pd.DataFrame, final_cols: list, conditional_fields: dict
) -> pd.DataFrame:
    crp = crp.copy()
    kyc = kyc.copy()

    # 🔗 Outer merge: CRP as left, KYC as right
    merged = pd.merge(
        crp,
        kyc,
        on=["Account_Number", "Customer_Number"],
        how="outer",
        suffixes=("_CRP", "_KYC"),
        indicator=True,
    )

    # ✅ Drop unmatched CRP records (left_only)
    merged = merged[merged["_merge"].isin(["both", "right_only"])].copy()

    # 🧠 Apply contradiction-safe conditional logic
    merged = apply_conditional_merge(merged, conditional_fields)

    # 🧹 Drop redundant _CRP/_KYC columns
    merged = drop_redundant_columns(merged, conditional_fields)

    # 🧾 Add merge source flag
    merged["Merge_Source"] = merged["_merge"].map(
        {"both": "Matched", "right_only": "KYC Only"}
    )

    # 🆔 Insert record ID
    merged.insert(0, "Record_ID", range(1, len(merged) + 1))

    # 🧼 Drop _merge column
    merged.drop(columns=["_merge"], inplace=True, errors="ignore")

    # 🗂️ Reorder columns if final_cols is valid
    if set(final_cols).issubset(merged.columns):
        merged = merged[final_cols]

    print(f"✅ Final merged dataset: {len(merged)} records")
    return merged


# ─────────────────────────────────────────────────────────
# 🧩 Boolean Checker (safe)
# ─────────────────────────────────────────────────────────


def debug_check_booleans(df: pd.DataFrame):
    """Safely verify boolean columns without .str errors."""
    if not isinstance(df, pd.DataFrame):
        print("⚠️ Input is not a DataFrame.")
        return

    for col in df.columns:
        if pd.api.types.is_bool_dtype(df[col]) or pd.api.types.is_dtype_equal(
            df[col].dtype, "boolean"
        ):
            invalid_mask = df[col].astype("string").isin(["", "nan", "None"])
            if invalid_mask.any():
                print(
                    f"⚠️ Column '{col}' has {invalid_mask.sum()} invalid boolean values."
                )


# ─────────────────────────────────────────────────────────
# 🧩 CONDITIONAL MERGE ENGINE
# ─────────────────────────────────────────────────────────
def apply_conditional_merge(df: pd.DataFrame, conditional_fields: dict) -> pd.DataFrame:
    df = df.copy()

    def is_missing(s: pd.Series) -> pd.Series:
        s_str = s.astype("string")
        return s.isna() | (s_str.str.strip() == "")

    for new_field, (kyc_field, crp_field) in conditional_fields.items():
        kyc_series = (
            df[kyc_field] if kyc_field in df.columns else pd.Series([pd.NA] * len(df))
        )
        crp_series = (
            df[crp_field] if crp_field in df.columns else pd.Series([pd.NA] * len(df))
        )

        df[new_field] = pd.NA

        both = df["_merge"] == "both"
        right_only = df["_merge"] == "right_only"

        # ✅ Matched: CRP preferred; fallback to KYC if CRP missing/blank
        crp_missing = is_missing(crp_series)
        df.loc[both, new_field] = crp_series.where(~crp_missing, kyc_series)[both]

        # ✅ Unmatched KYC: use KYC
        df.loc[right_only, new_field] = kyc_series[right_only]

    return df


# ─────────────────────────────────────────────────────────
# 🧼 SANITIZATION UTILITIES
# ─────────────────────────────────────────────────────────


def sanitize_booleans(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for col in df.columns:
        if pd.api.types.is_bool_dtype(df[col]) or pd.api.types.is_dtype_equal(
            df[col].dtype, "boolean"
        ):
            df[col] = df[col].replace("", pd.NA).fillna(False).astype(bool)
    return df


# ─────────────────────────────────────────────────────────
# 🧹 REMOVE DUPLICATES — KEEP HIGHEST QUALITY RECORD
# ─────────────────────────────────────────────────────────
def deduplicate_records(df: pd.DataFrame) -> pd.DataFrame:
    """
    Deduplicate rows based on logical keys (Account_Number + Customer_Number).
    Keeps the highest-quality record using the following priority:
      1️⃣ IsMatched == True
      2️⃣ Latest COB_Date
      3️⃣ Latest Account_Open_Dt
    Provides summary of duplicates found and dropped.
    """

    df = df.copy()

    # --- Normalize keys ---
    key_cols = []
    for key in ["Account_Number", "Customer_Number"]:
        if key in df.columns:
            df[key] = df[key].astype(str).str.strip().str.upper()
            key_cols.append(key)
        else:
            print(f"⚠️ Missing key column: {key}")

    if not key_cols:
        print("❌ No key columns found. Skipping deduplication.")
        return df

    # --- Convert date columns safely ---
    for col in ["COB_Date", "Account_Open_Dt"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    # --- Ensure IsMatched exists and valid ---
    if "IsMatched" not in df.columns:
        df["IsMatched"] = False
    else:
        df["IsMatched"] = df["IsMatched"].fillna(False).astype(bool)

    # --- Count duplicates before dropping ---
    total_before = len(df)
    dup_count = df.duplicated(subset=key_cols, keep=False).sum()

    if dup_count > 0:
        print(f"🔁 Found {dup_count} duplicate rows based on {key_cols}")

    # --- Sort with full priority ---
    df = df.sort_values(
        by=key_cols + ["IsMatched", "COB_Date", "Account_Open_Dt"],
        ascending=[True, True, False, False, False],
        na_position="last",
    )

    # --- Drop duplicates: keep the best (first) ---
    df = df.drop_duplicates(subset=key_cols, keep="first").reset_index(drop=True)

    total_after = len(df)
    dropped = total_before - total_after

    print(
        f"✅ Deduplication complete — {dropped} duplicate rows removed. Remaining: {total_after}"
    )

    # --- Optional sanity check for residuals ---
    dup_residual = df.duplicated(subset=key_cols, keep=False)
    if dup_residual.any():
        print(
            f"⚠️ Warning: {dup_residual.sum()} residual duplicates remain even after cleanup."
        )
        # Uncomment below to inspect
        # print(df.loc[dup_residual, key_cols + ['IsMatched', 'COB_Date']])

    return df


# ─────────────────────────────────────────────────────────
# 🧩 COLUMN CLEANUP
# ─────────────────────────────────────────────────────────


def drop_redundant_columns(df: pd.DataFrame, conditional_fields: dict) -> pd.DataFrame:
    df = df.copy()
    redundant = {
        k for pair in conditional_fields.values() for k in pair if k in df.columns
    }
    df.drop(columns=list(redundant), inplace=True, errors="ignore")
    return df


# ─────────────────────────────────────────────────────────
# 🚀 FINAL MERGE ENGINE PIPELINE
# ─────────────────────────────────────────────────────────


def merge_engine(df: pd.DataFrame, conditional_fields: dict) -> pd.DataFrame:
    df = sanitize_booleans(df)
    df = apply_conditional_merge(df, conditional_fields)
    debug_check_booleans(df)  # ✅ added back safely
    df = deduplicate_records(df)
    df = drop_redundant_columns(df, conditional_fields)
    return df


# Explicit Power Query Part-3 removals (columns fully dropped after unified merge)
PQ_EXPLICIT_REMOVALS = [
    "BusinessDate",
    "COB_DATE2",
    "BranchCode",
    "Br__Code___Name2",
    "Account_Num",
    "Account_No2",
    "Title_of_Account",
    "Account_Title2",
    "Purpose",
    "Purpose_of_Account2",
    "PFAMAPPROVAL",
    "Approval_Obtained2",
    "Account_Open_Date",
    "Date_Account_Opening2",
    "Product_Code",
    "Product_Type_of_Account2",
    "Currency",
    "Currency_Type_of_Account2",
    "MODEDEPOSITS",
    "Dominant_Mode_of_Deposit2",
    "MODEWITHDRAW",
    "Dominant_Mode_of_Withdrawal2",
    "SOURCE_OF_INCOME",
    "Source_Of__Income2",
    "UNSCLISTST",
    "Account_Screened_UNSC_List2",
    "MONTH_TOVER_RG",
    "Expected_Monthly_Turnover__In_Range____for_Individual_Trusts_Societies2",
    "POSTING_RESTRICT",
    "Posting_restrict2",
    "LOCKED_AMOUNT",
    "Locked_Amount2",
    "StatusCode",
    "Status_of_Account2",
    "Cust_Open_Date",
    "Date_of_customer_opening",
    "Customer_Num",
    "Customer_ID2",
    "Customer_Full_Name",
    "Customer_Name2",
    "ID_Number",
    "CNIC___ID_No2",
    "SBP_Child_Industry",
    "SBP_Industry_Recording_Code2",
    "KYC_Annual_Turnover",
    "A_C_Turnover_P_A___in_Range__for_Individual_Trusts_Societies2",
    "TO_Greater_Than_10M",
    "A_C_Turnover_P_A___in_Amount__for_Individual_Trusts_Societies2",
    "EXP_MONTH_TOVER",
    "Expected_Monthly_Turnover__In_Amount____for_Indvidual_Trusts_Societies2",
    "KYC_NO_TRANS",
    "Expected_Montly_Credit_Transactions2",
    "KYC_Risk",
    "Risk_Level2",
    "Cust_Sector_Code",
    "Entity_Type__Sector_2",
    "Cust_Occupation",
    "Occupation2",
    "Employed_Since",
    "Employed_Since2",
    "Nature_of_Business",
    "Nature_Purpose_of_Business_Organisation2",
    "Relation_with_Minor",
    "Relationship_with_Minor2",
    "Sole_Proprietor_Name",
    "Name_of_Sole_Proprietor2",
    "Business_Turnover",
    "Business_Turnover2",
    "Customer_Position",
    "Title_or_Position2",
    "Cust_Current_Salary",
    "Salary___Other_Income2",
    "Customer_Comments",
    "Customer_Profile2",
    "Name_of_Employer",
    "Name_Of_Employer2",
    "Investment_in_Business",
    "Investment_in_Business2",
    "Date_of_incorporation_of_business",
    "Date_of_incorporation_of_Business2",
    "Political_Figure",
    "Status_Of_PEP2",
    "HRAMAPPROVAL",
    "Approval_Obtained_for_PEP",
    "Employment_Status",
    "Status2",
    "Exp_Counter_Parties",
    "Expected_Type_of_Counter_Parties2",
    "Exp_International_Geography",
    "Expected_Int_Geographies_for_Trx2",
    "Exp_Local_Geography",
    "Expected_Local_Geographies_for_Trx2",
    "Name_of_Business",
    "Name_OF_Business2",
    "Status_of_Ownership",
    "Satus_of_Ownership2",
    "Partner_Director_Name",
    "Name_of_directors_partners_turstees",
    "Partner_Director_IDs",
    "ID_Number_of_Partners_Directors_Trustees",
    "A003KYC_BaseValue",
    "CRP_BaseValue",
    "TurnoverGT1050M",
    "CRPTurnoverGT1050M",
    "A003_KYC_Turnover",
    "CRP_Turnover",
    "A003MonthTurnover",
    "CRP_MonthlyTurnover",
    "KYCMonthlyTurnoverGT1050M",
    "CRPMonthlyTurnoverGT1050M",
    "A003_Monthly_Credit_Turnover",
    "CRP_MonthlyTurnover_Number",
]


def pq_explicit_cleanup(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    drop_list = [c for c in PQ_EXPLICIT_REMOVALS if c in df.columns]
    return df.drop(columns=drop_list, errors="ignore")


# Part-10#
# ─────────────────────────────────────────────────────────
# 🗂️ Final column ordering (keep known “final” first; append any new)
# ─────────────────────────────────────────────────────────
def reorder_columns_final(df: pd.DataFrame, preferred_order: list) -> pd.DataFrame:
    df = df.copy()
    existing_preferred = [c for c in preferred_order if c in df.columns]
    rest = [c for c in df.columns if c not in existing_preferred]
    return df[existing_preferred + rest]


# ─────────────────────────────────────────────────────────
# 🚀 Main Pipeline Entry Point
# ─────────────────────────────────────────────────────────
def process_uploaded_files(
    files_data: dict, save_filename: str = "merged_file.xlsx"
) -> pd.DataFrame:
    """
    files_data: dict[str_filename, (dict[sheet_name->DataFrame], str_full_file_path)]
    Returns merged DataFrame and saves it as CSV and Excel in the folder of the uploaded files.
    Output is formatted to match Power Query merged file behavior.
    """
    try:
        # 1) Detect KYC & CRP roles (sheet-level) and base path
        kyc, crp, base_path = detect_file_roles(files_data)
        crp["Account_Number"] = crp["Account_No2"].astype(str).str.strip().str.upper()
        crp["Customer_Number"] = crp["Customer_ID2"].astype(str).str.strip().str.upper()
        kyc["Account_Number"] = kyc["Account_Num"].astype(str).str.strip().str.upper()
        kyc["Customer_Number"] = kyc["Customer_Num"].astype(str).str.strip().str.upper()

        # 2) Build output paths
        if base_path is None:
            base_path = os.getcwd()
        csv_out_path = os.path.join(base_path, save_filename)
        xlsx_out_path = os.path.join(base_path, save_filename.replace(".csv", ".xlsx"))

        # 3) Skip if both CSV and Excel already exist
        if os.path.exists(csv_out_path) and os.path.exists(xlsx_out_path):
            print(
                f"✅ Merged files already exist at:\n- {csv_out_path}\n- {xlsx_out_path}\nSkipping re-merge."
            )
            return pd.read_csv(csv_out_path, dtype=str, low_memory=False)

        print("ℹ️ No existing merged files found. Proceeding with merge...")

        # 4) Apply business logic
        kyc = apply_kyc_turnover_logic(kyc)
        crp = apply_crp_turnover_logic(crp)

        # 6) Apply conditional merge logic
        conditional_fields = {
            "Customer_Number": ("Customer_Num", "Customer_ID2"),
            "Account_Number": ("Account_Num", "Account_No2"),
            "TitleOfAccount": ("Title_of_Account", "Account_Title2"),
            "Purpose_of_Account": ("Purpose", "Purpose_of_Account2"),
            "ProductCode": ("Product_Code", "Product_Type_of_Account2"),
            "SourceOfIncome": ("SOURCE_OF_INCOME", "Source_Of__Income2"),
            "Account_Turnover_in_Numbers": ("A003_KYC_Turnover", "CRP_Turnover"),
            "Account_Monthly_Turnover_in_Numbers": (
                "A003_Monthly_Credit_Turnover",
                "CRP_MonthlyTurnover_Number",
            ),
            "CNIC_Number": ("ID_Number", "CNIC___ID_No2"),
            "KYCRisk": ("KYC_Risk", "Risk_Level2"),
            "EmployedSince": ("Employed_Since", "Employed_Since2"),
            "NatureOfBusiness": (
                "Nature_of_Business",
                "Nature_Purpose_of_Business_Organisation2",
            ),
            "NameOfBusiness": ("Name_of_Business", "Name_OF_Business2"),
            "StatusOfOwnership": ("Status_of_Ownership", "Satus_of_Ownership2"),
            "BusinessTurnover": ("Business_Turnover", "Business_Turnover2"),
            "InvestmentInBusiness": (
                "Investment_in_Business",
                "Investment_in_Business2",
            ),
            "DateOfIncorporationOfBusiness": (
                "Date_of_incorporation_of_business",
                "Date_of_incorporation_of_Business2",
            ),
            "PoliticalFigure": ("Political_Figure", "Approval_Obtained_for_PEP"),
            "ApprovalObtainedForPEP": ("PEP_Related", "Status_Of_PEP2"),
            "ApprovalObtained": ("HRAMAPPROVAL", "Approval_Obtained2"),
            "EmploymentStatus": ("Employment_Status", "Status2"),
            "ExpCounterParties": (
                "Exp_Counter_Parties",
                "Expected_Type_of_Counter_Parties2",
            ),
            "ExpInternationalGeography": (
                "Exp_International_Geography",
                "Expected_Int_Geographies_for_Trx2",
            ),
            "ExpLocalGeography": (
                "Exp_Local_Geography",
                "Expected_Local_Geographies_for_Trx2",
            ),
            "CustomerFullName": ("Customer_Full_Name", "Customer_Name2"),
            "Customer_Occupation": ("Cust_Occupation", "Occupation2"),
            "CustSectorCode": ("Cust_Sector_Code", "Entity_Type__Sector_2"),
            "Account_Open_Dt": ("Account_Open_Date", "Date_Account_Opening2"),
            "COB_Date": ("BusinessDate", "COB_DATE2"),
            "Branch_Code": ("BranchCode", "Br__Code___Name2"),
            "Customer_Creation_Date": ("Cust_Open_Date", "Date_of_customer_opening"),
            "SoleProprietorName": ("Sole_Proprietor_Name", "Name_of_Sole_Proprietor2"),
            "CustomerPosition": ("Customer_Position", "Title_or_Position2"),
            "Salary_Other_Income": ("Cust_Current_Salary", "Salary___Other_Income2"),
            "CustomerProfile": ("Customer_Comments", "Customer_Profile2"),
            "NameOfEmployer": ("Name_of_Employer", "Name_Of_Employer2"),
            "RelationWithMinor": ("Relation_with_Minor", "Relationship_with_Minor2"),
            "Funds_Provider_ID_Number": (
                "Funds_Provider_ID_Num",
                "Funds_Provider_ID_Num",
            ),
            "PartnerDirectorName": (
                "Partner_Director_Name",
                "Name_of_directors_partners_turstees",
            ),
            "PartnerDirectorIDs": (
                "Partner_Director_IDs",
                "ID_Number_of_Partners_Directors_Trustees",
            ),
            "Status_Of_Account": ("StatusCode", "Status_of_Account2"),
            "UNSC_Screening": ("UNSCLISTST", "Account_Screened_UNSC_List2"),
            "Expected_Monthly_Turnover_Individual": (
                "MONTH_TOVER_RG",
                "Expected_Monthly_Turnover__In_Range____for_Individual_Trusts_Societies2",
            ),
            "Expected_Monthly_Turnover_GT": (
                "EXP_MONTH_TOVER",
                "Expected_Monthly_Turnover__In_Amount____for_Indvidual_Trusts_Societies2",
            ),
            "Expected_Month_Tover_CRG": (
                "MON_TOVER_CRG",
                "MON_TOVER_CRG",
            ),
            "Expected_Month_Tover_CORP": (
                "EXP_MONTH_TOVER",
                "EXP_MONTH_TOVER",
            ),
            "Posting_Restrictions": ("POSTING_RESTRICT", "Posting_restrict2"),
            "LockedAmount": ("LOCKED_AMOUNT", "Locked_Amount2"),
            "SBPChildIndustry": ("SBP_Child_Industry", "SBP_Industry_Recording_Code2"),
            "KYCAnnualTurnover": (
                "KYC_Annual_Turnover",
                "A_C_Turnover_P_A___in_Range__for_Individual_Trusts_Societies2",
            ),
            "TOGreaterThan10M": (
                "TO_Greater_Than_10M",
                "A_C_Turnover_P_A___in_Amount__for_Individual_Trusts_Societies2",
            ),
            "AreaManagerApproval": ("PFAMAPPROVAL", "Approval_Obtained_for_PEP"),
            "Account_Turnover": ("A003KYC_BaseValue", "CRP_BaseValue"),
            "Account_TurnoverGT1050M": ("TurnoverGT1050M", "CRPTurnoverGT1050M"),
            "Account_Monthly_Turnover": ("A003MonthTurnover", "CRP_MonthlyTurnover"),
            "Account_Monthly_TurnoverGT1050M": (
                "KYCMonthlyTurnoverGT1050M",
                "CRPMonthlyTurnoverGT1050M",
            ),
            "CCY": ("Currency", "Currency_Type_of_Account2"),
            "Dominant_Mode_of_Deposit": ("MODEDEPOSITS", "Dominant_Mode_of_Deposit2"),
            "Dominant_Mode_of_Withdrawal": (
                "MODEWITHDRAW",
                "Dominant_Mode_of_Withdrawal2",
            ),
            "Expected_Monthly_Credit_Transactions": (
                "KYC_NO_TRANS",
                "Expected_Montly_Credit_Transactions2",
            ),
        }

        # 5) Merge datasets
        all_columns = list(
            set(kyc.columns).union(set(crp.columns)).union({"IsMatched"})
        )
        merged = merge_datasets(
            crp, kyc, final_cols=all_columns, conditional_fields=conditional_fields
        )
        # ─────────────────────────────────────────────────────────
        # 🧠 Full Mapping & Forensic Normalization
        # ─────────────────────────────────────────────────────────

        # 🔎 Debug invalid boolean values
        debug_check_booleans(merged)

        # 🔧 Fix invalid boolean values
        merged = sanitize_booleans(merged)

        # Step: Cleanup explicit PQ columns
        merged = pq_explicit_cleanup(merged)

        # 2) Standardize column names
        merged.columns = (
            merged.columns.str.strip().str.replace(" ", "_").str.replace("-", "_")
        )

        # 3) Normalize IsMatched flag
        if "IsMatched" in merged.columns:
            merged["IsMatched"] = merged["IsMatched"].apply(
                lambda x: "TRUE" if x else "FALSE"
            )

        # 4) Forensic capture of raw date fields
        raw_capture_cols = [
            "KYC_Update_Date",
            "Cust_DOB",
            "DATE_LAST_CR_CUST",
            "DATE_LAST_DR_CUST",
            "SURRENDER_DATE",
            "Ac_Last_Modif_Date",
            "ID_Val_Date",
            "Date_of_incorporation_of_business",
            "Date_Account_Opening2",
            "Date_of_customer_opening",
            "COB_DATE2",
            "BusinessDate",
        ]
        for col in raw_capture_cols:
            if col in merged.columns:
                merged[f"{col}_RAW"] = merged[col].astype(str)

        # 5) Normalize date formats
        format_map = {
            "Cust_Open_Date": "%m/%d/%Y",
            "Account_Open_Date": "%m/%d/%Y",
            "SURRENDER_DATE": "%d-%m-%Y",  # changed from %Y%m%d
            "DATE_LAST_CR_CUST": "%d-%m-%Y",
            "DATE_LAST_DR_CUST": "%d-%m-%Y",
            "Date_Account_Opening2": "%d-%m-%Y",
            "Date_of_customer_opening": "%m/%d/%Y %H:%M",
            "KYC_Update_Date": "%d-%m-%Y",
            "Cust_DOB": "%d-%m-%Y",
            "BusinessDate": "%m/%d/%Y",
            "Date_of_incorporation_of_business": "%d-%m-%Y",
            "Date_of_incorporation_of_Business2": "%d-%m-%Y",
            "COB_DATE2": "%d-%m-%Y",
            "ID_Val_Date": "%d-%m-%Y",
            "Ac_Last_Modif_Date": "%d-%m-%Y",
        }

        def normalize_date_columns(df: pd.DataFrame, format_map: dict) -> pd.DataFrame:
            # Columns to skip because they contain codes, not dates
            skip_cols = ["KYC_Update_Date"]  # you can add other similar columns here

            for col, fmt in format_map.items():
                if col in df.columns:
                    if col in skip_cols:
                        continue  # skip these columns entirely

                    def parse_date(val):
                        val = str(val).strip()
                        if val in ["", "NaT", "nan", "None"]:
                            return ""
                        try:
                            parsed = pd.to_datetime(val, format=fmt, errors="coerce")
                            return (
                                parsed.strftime("%Y-%m-%d")
                                if not pd.isnull(parsed)
                                else ""
                            )
                        except Exception:
                            return val

                    df[col] = df[col].apply(parse_date)
            return df

        # 6) Apply drift detection overlay
        def apply_drift_overlay(df: pd.DataFrame, raw_cols: list) -> pd.DataFrame:
            for col in raw_cols:
                raw_col = f"{col}_RAW"
                if col in df.columns and raw_col in df.columns:
                    # ✅ Ensure both are strings before using .str
                    df[raw_col] = df[raw_col].astype("string").fillna("")
                    df[col] = df[col].astype("string").fillna("")

                    df[f"{col}_Drift_Flag"] = (
                        df[raw_col].str.strip().fillna("") != ""
                    ) & (df[col].str.strip().fillna("") == "")
            return df

        # Columns whose original text values must be preserved exactly as-is.
        # These are skipped in the generic fillna/string-cast loops below so that
        # values like "N/A" and "NA" are not silently converted to empty strings.
        PRESERVE_AS_IS = {"InvestmentInBusiness"}

        # Ensure all non-numeric columns are safe strings
        for col in merged.columns:
            if col in PRESERVE_AS_IS:
                # Cast to plain str so downstream code can always call .str methods,
                # but do NOT fillna — keep "N/A" / "NA" intact.
                merged[col] = merged[col].astype(str).str.strip()
                merged[col] = merged[col].replace({"nan": "", "<NA>": "", "NaT": ""})
                continue
            if (
                not pd.api.types.is_numeric_dtype(merged[col])
                and merged[col].dtype != bool
            ):
                merged[col] = merged[col].astype("string").fillna("")

        merged = apply_drift_overlay(merged, raw_capture_cols)

        # 7) Fill missing values and enforce string type, but preserve booleans
        for col in merged.columns:
            if col in PRESERVE_AS_IS:
                continue  # already handled above — do not overwrite
            if merged[col].dtype == "boolean" or merged[col].dtype == bool:
                merged[col] = merged[col].fillna(False)  # keep True/False clean
            elif pd.api.types.is_numeric_dtype(merged[col]):
                merged[col] = merged[col].fillna(0)  # or leave NaN if needed
            else:
                merged[col] = (
                    merged[col].fillna("").replace({pd.NaT: ""}).astype(str).str.strip()
                )

        merged.columns.name = None  # Prevent Excel "Index" header

        # 8) Final safe ID formatting
        id_columns = [
            "Account_Number",
            "Customer_Number",
            "CNIC_Number",
            "PartnerDirectorIDs",
            "BAF_PEN_ACCT",
            "Guradian_ID_Number2",
            "NTN",
        ]

        def apply_id_cleanup(df: pd.DataFrame, id_cols: list) -> pd.DataFrame:
            def clean_numeric_id(x):
                val = str(x).strip()
                if val in ["", "NaT", "nan", "None"]:
                    return ""
                try:
                    return str(int(float(val)))
                except:
                    return val

            drift_flags = {}  # collect new flag columns here

            for col in id_cols:
                if col in df.columns:
                    raw = df[col].astype(str)
                    cleaned = raw.apply(clean_numeric_id)
                    df[col] = cleaned
                    drift_flags[f"{col}_Drift_Flag"] = (
                        raw != cleaned
                    )  # collect instead of inserting

            # Add all drift flag columns at once to avoid fragmentation
            if drift_flags:
                df = pd.concat([df, pd.DataFrame(drift_flags, index=df.index)], axis=1)

            # Optional: defragment DataFrame to ensure best performance
            df = df.copy()

            return df

        merged = apply_id_cleanup(merged, id_columns)

        # 12) Reorder columns to match Power Query output
        preferred_order = [
            "Account_Number",
            "Customer_Number",
            "IsMatched",
            "CustomerFullName",
            "CNIC_Number",
            "TitleOfAccount",
            "Purpose_of_Account",
            "ProductCode",
            "CCY",
            "Account_Open_Dt",
            "COB_Date",
            "Branch_Code",
            "Customer_Creation_Date",
            "CustSectorCode",
            "Customer_Occupation",
            "EmploymentStatus",
            "EmployedSince",
            "NameOfEmployer",
            "CustomerPosition",
            "NatureOfBusiness",
            "NameOfBusiness",
            "StatusOfOwnership",
            "SoleProprietorName",
            "RelationWithMinor",
            "PartnerDirectorName",
            "PartnerDirectorIDs",
            "SBPChildIndustry",
            "UNSC_Screening",
            "PoliticalFigure",
            "ApprovalObtainedForPEP",
            "AreaManagerApproval",
            "SourceOfIncome",
            "Salary_Other_Income",
            "InvestmentInBusiness",
            "BusinessTurnover",
            "Account_Turnover",
            "Account_Turnover_in_Numbers",
            "Account_TurnoverGT1050M",
            "KYCAnnualTurnover",
            "TOGreaterThan10M",
            "KYC_Ann_TO_Corporate",
            "TO_Greater_Than_50M",
            "Account_Monthly_Turnover",
            "Account_Monthly_Turnover_in_Numbers",
            "Account_Monthly_TurnoverGT1050M",
            "Expected_Monthly_Turnover_Individual",
            "Expected_Monthly_Turnover_GT",
            "Expected_Month_Tover_CRG",
            "Expected_Month_Tover_CORP",
            "Dominant_Mode_of_Deposit",
            "Dominant_Mode_of_Withdrawal",
            "Expected_Monthly_Credit_Transactions",
            "Posting_Restrictions",
            "LockedAmount",
            "CustomerProfile",
            "KYCRisk",
            # Add remaining Power Query columns here
        ]  # Your full column order list goes here

        def remove_empty_duplicates(df: pd.DataFrame) -> pd.DataFrame:
            """
            Remove duplicate columns that are completely empty.
            Keeps the first occurrence with any value, removes only fully empty duplicates.
            """
            cols_to_keep = []
            seen = {}

            for col in df.columns:
                # If column already seen
                if col in seen:
                    # Remove only if this column is completely empty
                    if df[col].isna().all() or (df[col] == "").all():
                        continue  # skip adding this empty duplicate
                cols_to_keep.append(col)
                seen[col] = True

            df = df[cols_to_keep]
            return df

        # ------------------- Usage -------------------

        merged = reorder_columns_final(merged, preferred_order)

        # Remove empty duplicate columns safely
        merged = remove_empty_duplicates(merged)

        # ✅ Strip suffixes like _Drift_Flag and _RAW to match Power Query
        suffixes_to_remove = ["_Drift_Flag", "_RAW"]

        def strip_suffixes(col_name):
            for suffix in suffixes_to_remove:
                if col_name.endswith(suffix):
                    return col_name.replace(suffix, "")
            return col_name

        merged.columns = [strip_suffixes(col) for col in merged.columns]

        # Remove exact duplicate columns after stripping suffixes
        merged = merged.loc[:, ~merged.columns.duplicated()]
        # 13) Empty-output guard
        if merged.shape[0] == 0:
            print("⚠️ Final merged output is empty. Returning audit placeholder.")
            merged = pd.DataFrame(
                columns=["Account_Number", "Customer_Number", "IsMatched"]
            )
        else:
            print(f"✅ Merge successful. Rows: {merged.shape[0]}")

        # 14) Sample output for validation
        if "CNIC_Number" in merged.columns:
            print(f"✅ Final row count: {merged.shape[0]}")
            print(f"📊 Populated cells: {merged.count().sum()}")

        # 16) Save Excel with ID formatting
        xlsx_out_path = os.path.join(base_path, save_filename.replace(".csv", ".xlsx"))
        export_to_excel(merged, xlsx_out_path, id_columns=id_columns)
        print(f"📁 Merged Excel file saved at: {xlsx_out_path}")

        return merged

    except Exception as e:
        print(f"❌ Critical failure in processing pipeline: {e}")
        return pd.DataFrame(columns=["Account_Number", "Customer_Number", "IsMatched"])
