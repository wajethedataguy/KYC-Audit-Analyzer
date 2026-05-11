import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


def convert_to_date(val):
    """Parse values that may be YYYYMMDD or DD-MM-YYYY."""
    s = str(val).strip().replace(".0", "")
    if not s or s.lower() in ["nan", "none", "null"]:
        return pd.NaT
    # Try YYYYMMDD
    try:
        return datetime.strptime(s, "%Y%m%d")
    except ValueError:
        pass
    # Try DD-MM-YYYY
    try:
        return datetime.strptime(s, "%d-%m-%Y")
    except ValueError:
        return pd.NaT


def extract_latest_kyc_date(val):
    """Extract latest YYMMDD-style date from KYC update string."""
    try:
        chunks = str(val).split("?")
        dates = []
        for chunk in chunks:
            raw = chunk.strip()[:6]
            if len(raw) == 6 and raw.isdigit():
                yy, mm, dd = raw[:2], raw[2:4], raw[4:]
                full_yyyymmdd = f"20{yy}{mm}{dd}"
                parsed = pd.to_datetime(full_yyyymmdd, format="%Y%m%d", errors="coerce")
                if pd.notna(parsed):
                    dates.append(parsed)
        return max(dates) if dates else pd.NaT
    except Exception:
        return pd.NaT


@register_logic(
    name="078_kyc_review_frequency_breach",
    description="Flags customers whose KYC profile was not reviewed within prescribed frequency based on risk level, with grace of two months and month-start due date.",
    category="Compliance & Screening",
)
def logic_078_kyc_review_frequency_breach(
    dataframes: dict, mode: str = "full"
) -> pd.DataFrame:
    try:
        kyc_key = next((k for k in dataframes if "merged_file.xlsx" in k.lower()), None)
        if not kyc_key:
            raise ValueError("Merged KYC file not found.")

        df_main = dataframes[kyc_key]
        if isinstance(df_main, tuple):
            df_main = next(iter(df_main[0].values())) if df_main[0] else pd.DataFrame()

        df_main.columns = (
            df_main.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )

        required_cols = [
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "CUSTOMER_NUMBER",
            "CNIC_NUMBER",
            "ACCOUNT_OPEN_DT",
            "COB_DATE",
            "KYCRISK",
            "KYC_UPDATE_DATE",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        missing = [col for col in required_cols if col not in df_main.columns]
        if missing:
            raise ValueError(f"Missing required columns: {missing}")

        df = df_main[required_cols].copy()

        # 🔍 Parse dates
        df["ACCOUNT_OPEN_DATE"] = df["ACCOUNT_OPEN_DT"].apply(convert_to_date)
        df["COB_DATE_PARSED"] = df["COB_DATE"].apply(convert_to_date)
        df["KYC_UPDATE_DATE_PARSED"] = df["KYC_UPDATE_DATE"].apply(
            extract_latest_kyc_date
        )

        today = pd.to_datetime(datetime.today().strftime("%Y-%m-%d"))

        def calculate_overdue(row):
            try:
                risk = str(row["KYCRISK"]).strip().lower()
                last_review = row["KYC_UPDATE_DATE_PARSED"]

                if pd.isna(last_review):
                    return pd.Series([pd.NA] * 3)

                if "high" in risk:
                    freq_years = 1
                elif "medium" in risk:
                    freq_years = 3
                elif "low" in risk:
                    freq_years = 5
                else:
                    return pd.Series([pd.NA] * 3)

                # Grace: skip review month + next month, start from 1st of following month
                effective_start = (last_review + pd.DateOffset(months=2)).replace(day=1)
                overdue_date = effective_start + pd.DateOffset(years=freq_years)

                days_overdue = (today - overdue_date).days

                # ✅ Flag immediately if overdue_date has passed
                if days_overdue >= 0:
                    return (
                        last_review.strftime("%d-%b-%y"),
                        overdue_date.strftime("%d-%b-%y"),
                        days_overdue,
                    )
                else:
                    return (
                        last_review.strftime("%d-%b-%y"),
                        overdue_date.strftime("%d-%b-%y"),
                        pd.NA,
                    )
            except Exception:
                return pd.Series([pd.NA] * 3)

        df[
            [
                "KYC_LAST_REVIEWED_DATE",
                "KYC_REVIEW_OVERDUE_DATE",
                "NUMBER_OF_DAYS_OVERDUE",
            ]
        ] = df.apply(lambda row: pd.Series(calculate_overdue(row)), axis=1)

        output_cols = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "KYCRISK",
            "KYC_LAST_REVIEWED_DATE",
            "KYC_REVIEW_OVERDUE_DATE",
            "NUMBER_OF_DAYS_OVERDUE",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]

        output = df[df["NUMBER_OF_DAYS_OVERDUE"].notna()].copy()
        output = (
            output[output_cols]
            .sort_values(by="NUMBER_OF_DAYS_OVERDUE", ascending=False)
            .reset_index(drop=True)
        )

        if output.empty:
            output = pd.DataFrame([{col: pd.NA for col in output_cols}])

        if mode == "full" and not output.empty:
            file_path = (
                f"kyc_review_frequency_breach_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            )
            output.to_excel(file_path, index=False)
        elif mode == "ui" and not output.empty:
            output = output[output_cols]

        return output

    except Exception as e:
        print(f"❌ Error in logic_078_kyc_review_frequency_breach: {e}")
        return pd.DataFrame(
            [
                {
                    col: pd.NA
                    for col in [
                        "CUSTOMER_NUMBER",
                        "ACCOUNT_NUMBER",
                        "TITLEOFACCOUNT",
                        "KYCRISK",
                        "KYC_LAST_REVIEWED_DATE",
                        "KYC_REVIEW_OVERDUE_DATE",
                        "NUMBER_OF_DAYS_OVERDUE",
                        "PRODUCTDESC",
                        "SECTOR_DESCRIPTION",
                    ]
                }
            ]
        )
