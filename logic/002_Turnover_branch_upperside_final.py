import pandas as pd
import re
from datetime import datetime
from dateutil.parser import parse
from KYC_Viewer.utils import register_logic


def normalize_cols(df: pd.DataFrame) -> pd.DataFrame:
    """Return a copy with normalized column names."""
    df = df.copy()
    df.columns = (
        df.columns.str.strip()
        .str.upper()
        .str.replace(" ", "_")
        .str.replace("-", "_")
        .str.replace(r"[^\w]", "_", regex=True)
    )
    return df


@register_logic(
    name="002_turnover_breach_upperside_final",
    description="Upper-side breach with strict date normalization (YYYYMMDD), numeric conversion, ranges, multi-values, and float tolerance. Requires Account_Turnover and KYC ranges to match for individuals; uses KYC_Ann_TO_Corporate for corporates.",
    category="Turnover Breach",
)
def logic_002_turnover_breach_upperside_final(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    try:
        required1 = {
            "ACCOUNT_NUMBER",
            "CUSTOMER_NUMBER",
            "TITLEOFACCOUNT",
            "ACCOUNT_OPEN_DT",
            "COB_DATE",
            "ACCOUNT_TURNOVER",
            "ACCOUNT_TURNOVER_IN_NUMBERS",
            "KYC_ANN_TO_CORPORATE",  # corporate KYC range lives in File1
        }
        required2 = {
            "ACCOUNT_NO_",
            "TOTAL_ACTUAL_TO",
            "ACCOUNT_T_O_WITH_TOLERENCE",
            "ACCOUNT_T_O__P_A_KYC_",  # retail/individual KYC range lives in File2
        }

        file1_key, df1 = None, None
        file2_key, df2 = None, None

        for k, df in dataframes.items():
            if not isinstance(df, pd.DataFrame):
                continue
            df_norm = normalize_cols(df)
            if required1.issubset(set(df_norm.columns)):
                file1_key, df1 = k, df_norm
            if required2.issubset(set(df_norm.columns)):
                file2_key, df2 = k, df_norm

        if df1 is None or df2 is None:
            raise ValueError("❌ Required File1 or File2 missing.")

        df1 = df1[
            [
                "ACCOUNT_NUMBER",
                "CUSTOMER_NUMBER",
                "TITLEOFACCOUNT",
                "ACCOUNT_OPEN_DT",
                "COB_DATE",
                "ACCOUNT_TURNOVER",
                "ACCOUNT_TURNOVER_IN_NUMBERS",
                "KYC_ANN_TO_CORPORATE",
                "PRODUCTDESC",
                "SECTOR_DESCRIPTION",
            ]
        ]
        df2 = df2[
            [
                "ACCOUNT_NO_",
                "TOTAL_ACTUAL_TO",
                "ACCOUNT_T_O_WITH_TOLERENCE",
                "ACCOUNT_T_O__P_A_KYC_",
            ]
        ]

        # Align join keys (preserve string form, including leading zeros if any)
        df1["ACCOUNT_NUMBER"] = (
            pd.to_numeric(df1["ACCOUNT_NUMBER"], errors="coerce")
            .astype("Int64")
            .astype(str)
        )
        df2["ACCOUNT_NO_"] = (
            pd.to_numeric(df2["ACCOUNT_NO_"], errors="coerce")
            .astype("Int64")
            .astype(str)
        )

        df = pd.merge(
            df1, df2, left_on="ACCOUNT_NUMBER", right_on="ACCOUNT_NO_", how="left"
        )

        # Numerics
        df["ACCOUNT_TURNOVER_IN_NUMBERS"] = pd.to_numeric(
            df["ACCOUNT_TURNOVER_IN_NUMBERS"], errors="coerce"
        )
        df["TOTAL_ACTUAL_TO"] = pd.to_numeric(df["TOTAL_ACTUAL_TO"], errors="coerce")
        df["ACCOUNT_T_O_WITH_TOLERENCE"] = pd.to_numeric(
            df["ACCOUNT_T_O_WITH_TOLERENCE"], errors="coerce"
        )

        # -----------------------------------------
        # Parity check: Individuals vs Corporates
        # -----------------------------------------
        def parity_check(row):
            # Individual sources (File1)
            indiv_str = str(row.get("ACCOUNT_TURNOVER", "")).strip().lower()
            indiv_num = row.get("ACCOUNT_TURNOVER_IN_NUMBERS", None)

            # Corporate source (File1)
            corp_str = str(row.get("KYC_ANN_TO_CORPORATE", "")).strip().lower()

            # KYC values (File2) — may be multi-line
            kyc_values_raw = str(row.get("ACCOUNT_T_O__P_A_KYC_", "")).split("\n")
            kyc_values = [v.strip().lower() for v in kyc_values_raw if v.strip()]

            # If we have individual fields, try matching them against File2 KYC
            if indiv_str or pd.notna(indiv_num):
                for val in kyc_values:
                    # Direct string equality
                    if indiv_str and indiv_str == val:
                        return True
                    # Special case
                    if indiv_str == "below 10m" or val == "below 10m":
                        return True
                    # Numeric tolerance against KYC numeric-like values
                    try:
                        if pd.notna(indiv_num) and abs(indiv_num - float(val)) < 1:
                            return True
                    except:
                        continue

            # If corporate field exists, also match it against File2 KYC
            if corp_str:
                for val in kyc_values:
                    if corp_str == val:
                        return True
                    if corp_str == "below 10m" or val == "below 10m":
                        return True

            return False

        df["parity_match"] = df.apply(parity_check, axis=1)
        df = df[df["parity_match"]].copy()

        # Dates
        def clean_date(v):
            if pd.isna(v):
                return pd.NaT
            s = str(v).strip()
            s = re.sub(r"\s+\d{2}:\d{2}:\d{2}", "", s)
            try:
                d = parse(s, dayfirst=True)
                return pd.to_datetime(d.strftime("%Y%m%d"))
            except:
                return pd.NaT

        df["ACCOUNT_OPEN_DT"] = df["ACCOUNT_OPEN_DT"].apply(clean_date)
        df["COB_DATE"] = df["COB_DATE"].apply(clean_date)
        df["days_active"] = (df["COB_DATE"] - df["ACCOUNT_OPEN_DT"]).dt.days

        # Breach: actual exceeds tolerance
        df["turnover_breach_upperside_extra"] = df["TOTAL_ACTUAL_TO"].fillna(0) - df[
            "ACCOUNT_T_O_WITH_TOLERENCE"
        ].fillna(0)

        mask = (
            df["TOTAL_ACTUAL_TO"].notna()
            & df["ACCOUNT_T_O_WITH_TOLERENCE"].notna()
            & (df["turnover_breach_upperside_extra"] > 0)
        )
        df_out = df.loc[mask]

        # Output
        output_cols = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "ACCOUNT_TURNOVER",
            "ACCOUNT_T_O__P_A_KYC_",
            "KYC_ANN_TO_CORPORATE",
            "TOTAL_ACTUAL_TO",
            "ACCOUNT_T_O_WITH_TOLERENCE",
            "turnover_breach_upperside_extra",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        df_out = df_out[[c for c in output_cols if c in df_out.columns]]

        df_out.columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "Account_Turnover",
            "Account_T_O__P_A_KYC_",
            "KYC_Ann_TO_Corporate",
            "Total_Actual_TO",
            "Account_T_O_With_Tolerence",
            "turnover_breach_upperside_extra",
            "ProductDesc",
            "Sector_Description",
        ]

        if df_out.empty:
            df_out = pd.DataFrame([{col: pd.NA for col in df_out.columns}])

        if mode == "full":
            file_path = (
                f"turnover_breach_upperside_final_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            )
            df_out.to_excel(file_path, index=False)
            print(f"📁 Saved output to: {file_path}")

        return df_out

    except Exception as e:
        print("❌ ERROR:", e)
        return pd.DataFrame(
            [
                {
                    "Customer_Number": pd.NA,
                    "Account_Number": pd.NA,
                    "TitleOfAccount": pd.NA,
                    "Account_Turnover": pd.NA,
                    "Account_T_O__P_A_KYC_": pd.NA,
                    "KYC_Ann_TO_Corporate": pd.NA,
                    "Total_Actual_TO": pd.NA,
                    "Account_T_O_With_Tolerence": pd.NA,
                    "turnover_breach_upperside_extra": pd.NA,
                    "ProductDesc": pd.NA,
                    "Sector_Description": pd.NA,
                }
            ]
        )
