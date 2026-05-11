import re
import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="102_understated_monthly_credit_transactions_flag",
    description=(
        "Flags KYC records where the difference between expected monthly credit "
        "transactions and average actual credit transactions exceeds a threshold "
        "in either direction."
    ),
    category="CDD & EDD Review",
)
def logic_102_understated_monthly_credit_transactions_flag(
    dataframes: dict,
    mode: str = "full",
    top_n=None,  # optional cap if required
    threshold: float = 50.0,  # |RAW_DIFF| threshold
) -> pd.DataFrame:
    try:
        # --- Load KYC file ---
        kyc_key = next((k for k in dataframes if "merged_file.xlsx" in k.lower()), None)
        df_kyc = dataframes.get(kyc_key)
        if isinstance(df_kyc, tuple):
            df_kyc = next(iter(df_kyc[0].values())) if df_kyc[0] else pd.DataFrame()
        if df_kyc is None or df_kyc.empty:
            raise ValueError("KYC file not found or empty.")

        # --- Normalize KYC columns ---
        df_kyc.columns = (
            df_kyc.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
            .str.replace(r"[^\w]", "_", regex=True)
        )

        # --- Find turnover file ---
        df_turnover = None
        for df in dataframes.values():
            if isinstance(df, tuple):
                df = next(iter(df[0].values())) if df[0] else pd.DataFrame()
            if isinstance(df, pd.DataFrame):
                cols = df.columns.str.upper().str.strip()
                if "ACCOUNT_NUM" in cols and "CREDIT_COUNT" in cols:
                    df_turnover = df
                    break
        if df_turnover is None or df_turnover.empty:
            raise ValueError(
                "Turnover file not found or missing ACCOUNT_NUM and CREDIT_COUNT columns."
            )

        # --- Normalize turnover columns ---
        df_turnover.columns = (
            df_turnover.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )

        # --- Normalize merge keys ---
        df_kyc["ACCOUNT_NUMBER_CLEAN"] = (
            df_kyc["ACCOUNT_NUMBER"]
            .astype(str)
            .str.replace(r"\.0$", "", regex=True)
            .str.replace(r"[^\d]", "", regex=True)
        )
        df_turnover["ACCOUNT_NUM_CLEAN"] = (
            df_turnover["ACCOUNT_NUM"].astype(str).str.replace(r"[^\d]", "", regex=True)
        )

        # --- Auto-detect expected monthly credit transaction column ---
        expected_cols = [
            col
            for col in df_kyc.columns
            if "EXPECTED" in col and "CREDIT" in col and "TRANSACTION" in col
        ]
        if not expected_cols:
            raise ValueError(
                "Expected monthly credit transaction column not found in KYC file."
            )
        expected_col = expected_cols[0]

        # --- Convert numeric fields ---
        df_kyc["EXPECTED_MONTHLY_CREDIT_TRANSACTIONS"] = pd.to_numeric(
            df_kyc[expected_col], errors="coerce"
        )
        df_turnover["CREDIT_COUNT"] = pd.to_numeric(
            df_turnover["CREDIT_COUNT"], errors="coerce"
        )

        # --- Explicit mixed-format date parser ---
        def parse_mixed_date_value(val):
            if pd.isna(val):
                return pd.NaT
            s = str(val).strip()
            if not s:
                return pd.NaT
            if re.fullmatch(r"\d{8}", s):
                try:
                    return datetime.strptime(s, "%Y%m%d")
                except ValueError:
                    return pd.NaT
            if re.fullmatch(r"\d{2}/\d{2}/\d{4}", s):
                try:
                    return datetime.strptime(s, "%m/%d/%Y")
                except ValueError:
                    return pd.NaT
            if re.fullmatch(r"\d{2}-\d{2}-\d{4}", s):
                try:
                    return datetime.strptime(s, "%d-%m-%Y")
                except ValueError:
                    return pd.NaT
            if re.match(r"\d{4}-\d{2}-\d{2}", s):
                try:
                    return datetime.strptime(s[:10], "%Y-%m-%d")
                except ValueError:
                    return pd.NaT
            try:
                return pd.to_datetime(s, errors="coerce")
            except Exception:
                return pd.NaT

        def parse_mixed_date(series: pd.Series) -> pd.Series:
            return series.apply(parse_mixed_date_value)

        df_kyc["ACCOUNT_OPEN_DT"] = parse_mixed_date(df_kyc.get("ACCOUNT_OPEN_DT"))
        df_kyc["COB_DATE"] = parse_mixed_date(df_kyc.get("COB_DATE"))

        # --- Merge KYC + turnover ---
        df_joined = pd.merge(
            df_kyc,
            df_turnover[["ACCOUNT_NUM_CLEAN", "CREDIT_COUNT"]],
            left_on="ACCOUNT_NUMBER_CLEAN",
            right_on="ACCOUNT_NUM_CLEAN",
            how="left",
        )

        # --- Account age in months ---
        def calc_age_months(row):
            cob, open_dt = row["COB_DATE"], row["ACCOUNT_OPEN_DT"]
            if pd.notna(cob) and pd.notna(open_dt):
                days = (cob - open_dt).days
                return round((days / 365.0) * 12.0) if days > 0 else 0
            return 0

        df_joined["ACCOUNT_AGE_MONTHS"] = df_joined.apply(calc_age_months, axis=1)

        # --- Average actual monthly credit transactions ---
        def calc_avg_actual(row):
            credit_count, age_months = row["CREDIT_COUNT"], row["ACCOUNT_AGE_MONTHS"]
            if pd.isna(credit_count) or age_months <= 0:
                return None
            return (
                credit_count / 12.0
                if age_months > 12
                else credit_count / float(age_months)
            )

        df_joined["ACTUAL_MONTHLY_AVG_CREDIT_TRANSACTIONS"] = df_joined.apply(
            calc_avg_actual, axis=1
        )

        # --- Differences ---
        df_joined["RAW_DIFF"] = (
            df_joined["EXPECTED_MONTHLY_CREDIT_TRANSACTIONS"]
            - df_joined["ACTUAL_MONTHLY_AVG_CREDIT_TRANSACTIONS"]
        )
        df_joined["DIFFERENCE_EXPECTED_ACTUAL"] = df_joined["RAW_DIFF"].abs()

        # --- Apply filter ---
        df_flagged = df_joined[
            (df_joined["RAW_DIFF"] > threshold) | (df_joined["RAW_DIFF"] < -threshold)
        ].copy()
        df_flagged = df_flagged.sort_values(
            by="DIFFERENCE_EXPECTED_ACTUAL", ascending=False
        )
        if top_n is not None:
            df_flagged = df_flagged.head(int(top_n))

        # --- Final output columns ---
        output_cols = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "CustSectorCode",
            "ProductDesc",
            "Sector_Description",
            "Expected_Monthly_Credit_Transactions",
            "Actual_Monthly_Avg_Credit_Transactions",
            "RAW_DIFF",
            "Difference_Expected_Actual",
            "Account_Age_Months",
            "Credit_Count",
            "Account_Open_Dt",
            "COB_Date",
        ]

        output = (
            df_flagged.rename(
                columns={
                    "CUSTOMER_NUMBER": "Customer_Number",
                    "ACCOUNT_NUMBER": "Account_Number",
                    "TITLEOFACCOUNT": "TitleOfAccount",
                    "CUSTSECTORCODE": "CustSectorCode",
                    "PRODUCTDESC": "ProductDesc",
                    "SECTOR_DESCRIPTION": "Sector_Description",
                    "EXPECTED_MONTHLY_CREDIT_TRANSACTIONS": "Expected_Monthly_Credit_Transactions",
                    "ACTUAL_MONTHLY_AVG_CREDIT_TRANSACTIONS": "Actual_Monthly_Avg_Credit_Transactions",
                    "DIFFERENCE_EXPECTED_ACTUAL": "Difference_Expected_Actual",
                    "ACCOUNT_AGE_MONTHS": "Account_Age_Months",
                    "CREDIT_COUNT": "Credit_Count",
                    "ACCOUNT_OPEN_DT": "Account_Open_Dt",
                    "COB_DATE": "COB_Date",
                }
            )[output_cols]
            .drop_duplicates()
            .reset_index(drop=True)
        )

        if output.empty:
            output = pd.DataFrame([{col: pd.NA for col in output_cols}])

        if mode == "full":
            file_path = f"monthly_credit_txn_gap_over_{int(threshold)}_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            output.to_excel(file_path, index=False)

        return output

    except Exception as e:
        print("❌ Error in logic_102_understated_monthly_credit_transactions_flag:", e)
        return pd.DataFrame(
            [
                {
                    "Customer_Number": pd.NA,
                    "Account_Number": pd.NA,
                    "TitleOfAccount": pd.NA,
                    "CustSectorCode": pd.NA,
                    "ProductDesc": pd.NA,
                    "Sector_Description": pd.NA,
                    "Expected_Monthly_Credit_Transactions": pd.NA,
                    "Actual_Monthly_Avg_Credit_Transactions": pd.NA,
                    "RAW_DIFF": pd.NA,
                    "Difference_Expected_Actual": pd.NA,
                    "Account_Age_Months": pd.NA,
                    "Credit_Count": pd.NA,
                    "Account_Open_Dt": pd.NA,
                    "COB_Date": pd.NA,
                }
            ]
        )
