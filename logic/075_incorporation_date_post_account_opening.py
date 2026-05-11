import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


def parse_account_open(val) -> pd.Timestamp:
    """
    Parse Account_Open_Dt which may be:
    - DD-MM-YYYY               e.g. 24-12-2009
    - YYYYMMDD                 e.g. 20091224
    - YYYY-DD-MM HH:MM:SS      e.g. 2010-06-01 00:00:00  (treated as 2010-Jan-06)
    - YYYY-MM-DD HH:MM:SS      e.g. 2010-06-01 00:00:00  (fallback)
    - YYYY-DD-MM               e.g. 2010-06-01
    - YYYY-MM-DD               e.g. 2010-06-01
    """
    if pd.isna(val):
        return pd.NaT

    s = str(val).strip()

    # Remove Excel float suffix like ".0"
    if s.endswith(".0"):
        s = s[:-2].strip()

    if not s or s.lower() in {"nan", "none", "null"}:
        return pd.NaT

    # 1) DD-MM-YYYY
    try:
        return pd.to_datetime(s, format="%d-%m-%Y", errors="raise")
    except Exception:
        pass

    # 2) YYYYMMDD (digits only)
    s_digits = "".join(ch for ch in s if ch.isdigit())
    if len(s_digits) == 8:
        try:
            return pd.to_datetime(s_digits, format="%Y%m%d", errors="raise")
        except Exception:
            pass

    # 3) Hyphenated with time: try YYYY-DD-MM first, then YYYY-MM-DD
    try:
        return pd.to_datetime(s, format="%Y-%d-%m %H:%M:%S", errors="raise")
    except Exception:
        pass
    try:
        return pd.to_datetime(s, format="%Y-%m-%d %H:%M:%S", errors="raise")
    except Exception:
        pass

    # 4) Hyphenated date only: try YYYY-DD-MM first, then YYYY-MM-DD
    try:
        return pd.to_datetime(s, format="%Y-%d-%m", errors="raise")
    except Exception:
        pass
    try:
        return pd.to_datetime(s, format="%Y-%m-%d", errors="raise")
    except Exception:
        pass

    return pd.NaT


def parse_incorp(val) -> pd.Timestamp:
    """Parse DateOfIncorporationOfBusiness which is always YYYYMMDD."""
    if pd.isna(val):
        return pd.NaT

    s = str(val).strip()

    if s.endswith(".0"):
        s = s[:-2].strip()

    if not s or s.lower() in {"nan", "none", "null"}:
        return pd.NaT

    return pd.to_datetime(s, format="%Y%m%d", errors="coerce")


@register_logic(
    name="075_incorporation_date_post_account_opening",
    description="Flags business entities whose incorporation date is later than their account opening date.",
    category="Compliance & Screening",
)
def logic_075_incorporation_date_post_account_opening(
    dataframes: dict, mode: str = "full"
) -> pd.DataFrame:
    try:
        # 🔍 Locate merged KYC file
        kyc_key = next((k for k in dataframes if "merged_file.xlsx" in k.lower()), None)
        if not kyc_key:
            raise ValueError("Merged KYC file not found.")
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
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "ACCOUNT_OPEN_DT",
            "DATEOFINCORPORATIONOFBUSINESS",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        missing = [col for col in required_cols if col not in df_main.columns]
        if missing:
            raise ValueError(f"Missing required columns in merged file: {missing}")

        df = df_main.copy()

        # 🔍 Parse both dates
        df["ACCOUNT_OPEN_DATE"] = df["ACCOUNT_OPEN_DT"].apply(parse_account_open)
        df["INCORPORATION_DATE"] = df["DATEOFINCORPORATIONOFBUSINESS"].apply(parse_incorp)

        # 🧠 Apply contradiction logic: incorporation AFTER account opening
        mask = (
            df["ACCOUNT_OPEN_DATE"].notna()
            & df["INCORPORATION_DATE"].notna()
            & (df["INCORPORATION_DATE"] > df["ACCOUNT_OPEN_DATE"])
        )
        df_filtered = df.loc[mask].copy()
        df_filtered["CONTRADICTION_REASON"] = (
            "Date of incorporation of business was found post to date of account opening in KYC profiles of business entities."
        )

        # 📤 Prepare output
        final_cols = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "ACCOUNT_OPEN_DT",
            "DATEOFINCORPORATIONOFBUSINESS",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
            "CONTRADICTION_REASON",
        ]
        output = df_filtered[final_cols].reset_index(drop=True)

        # 🧯 Handle empty output
        if output.empty:
            output = pd.DataFrame([{col: pd.NA for col in final_cols}])

        # 📁 Optional export
        if mode == "full" and not output.empty:
            file_path = f"incorporation_date_post_account_opening_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            output.to_excel(file_path, index=False)
        elif mode == "ui" and not output.empty:
            output = output[
                [
                    "CUSTOMER_NUMBER",
                    "ACCOUNT_NUMBER",
                    "TITLEOFACCOUNT",
                    "ACCOUNT_OPEN_DT",
                    "DATEOFINCORPORATIONOFBUSINESS",
                    "PRODUCTDESC",
                    "SECTOR_DESCRIPTION",
                ]
            ]

        return output

    except Exception as e:
        print(f"❌ Error in logic_075_incorporation_date_post_account_opening: {e}")
        return pd.DataFrame(
            [
                {
                    col: pd.NA
                    for col in [
                        "CUSTOMER_NUMBER",
                        "ACCOUNT_NUMBER",
                        "TITLEOFACCOUNT",
                        "ACCOUNT_OPEN_DT",
                        "DATEOFINCORPORATIONOFBUSINESS",
                        "PRODUCTDESC",
                        "SECTOR_DESCRIPTION",
                        "CONTRADICTION_REASON",
                    ]
                }
            ]
        )
