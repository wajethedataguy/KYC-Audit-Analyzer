import pandas as pd
from datetime import datetime
from dateutil.parser import parse
from KYC_Viewer.utils import register_logic
import re


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
    name="001_turnover_breach_lowerside_final",
    description="Lower-side breach with robust date normalization, numeric conversion, ranges, multi-values, and corporate KYC values.",
    category="Turnover Breach",
)
def logic_001_turnover_breach_lowerside_final(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    try:
        # Required column sets
        required1 = {
            "ACCOUNT_NUMBER",
            "CUSTOMER_NUMBER",
            "TITLEOFACCOUNT",
            "ACCOUNT_OPEN_DT",
            "ACCOUNT_TURNOVER",
            "ACCOUNT_TURNOVER_IN_NUMBERS",
            "KYC_ANN_TO_CORPORATE",
        }
        required2 = {
            "ACCOUNT_NO_",
            "TOTAL_ACTUAL_TO",
            "ACCOUNT_T_O_WITH_TOLERENCE_NEG",
            "ACCOUNT_T_O__P_A_KYC_",
            "COB_DATE",
        }

        df1, df2 = None, None
        for k, df in dataframes.items():
            if not isinstance(df, pd.DataFrame):
                continue
            df_norm = normalize_cols(df)
            if required1.issubset(set(df_norm.columns)):
                df1 = df_norm
            if required2.issubset(set(df_norm.columns)):
                df2 = df_norm

        if df1 is None or df2 is None:
            raise ValueError(
                f"❌ Required File1 or File2 missing.\n"
                f"File1 needs {required1}, File2 needs {required2}\n"
                f"Available columns: {[list(normalize_cols(df).columns) for df in dataframes.values() if isinstance(df, pd.DataFrame)]}"
            )

        # Select required columns
        df1 = df1[
            [
                "ACCOUNT_NUMBER",
                "CUSTOMER_NUMBER",
                "TITLEOFACCOUNT",
                "ACCOUNT_OPEN_DT",
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
                "ACCOUNT_T_O_WITH_TOLERENCE_NEG",
                "ACCOUNT_T_O__P_A_KYC_",
                "COB_DATE",
            ]
        ]

        # Normalize account IDs
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

        # Merge files
        df = pd.merge(
            df1, df2, left_on="ACCOUNT_NUMBER", right_on="ACCOUNT_NO_", how="left"
        )

        # Numeric conversion
        for col in [
            "ACCOUNT_TURNOVER_IN_NUMBERS",
            "TOTAL_ACTUAL_TO",
            "ACCOUNT_T_O_WITH_TOLERENCE_NEG",
        ]:
            df[col] = pd.to_numeric(df[col], errors="coerce")

        # Range mapping
        ranges = {
            "below 1m": (0, 1_000_000),
            "1m to 5m": (1_000_000, 5_000_000),
            "5m to 10m": (5_000_000, 10_000_000),
            "10m to 50m": (10_000_000, 50_000_000),
        }

        # Parity check
        def parity_check(row):
            base_num = row["ACCOUNT_TURNOVER_IN_NUMBERS"]
            base_str = str(row["ACCOUNT_TURNOVER"]).strip().lower()
            corp_str = str(row["KYC_ANN_TO_CORPORATE"]).strip().lower()
            kyc_values = str(row["ACCOUNT_T_O__P_A_KYC_"]).split("\n")
            if corp_str and corp_str != "nan":
                base_str = corp_str
            for v in kyc_values:
                val = v.strip().lower()
                if base_str == val:
                    return True
                if val in ranges and pd.notna(base_num):
                    low, high = ranges[val]
                    if low <= base_num <= high:
                        return True
                try:
                    if pd.notna(base_num) and abs(base_num - float(val)) < 1:
                        return True
                except:
                    continue
            return False

        df["parity_match"] = df.apply(parity_check, axis=1)
        df = df[df["parity_match"]].copy()

        # Robust date parser
        def clean_date(v):
            if pd.isna(v):
                return pd.NaT
            if isinstance(v, (int, float)) and v > 30000:
                try:
                    return pd.to_datetime("1899-12-30") + pd.to_timedelta(
                        int(v), unit="D"
                    )
                except:
                    return pd.NaT
            s = str(v).strip()
            if re.fullmatch(r"\d{8}", s):
                try:
                    return pd.to_datetime(s, format="%Y%m%d")
                except:
                    return pd.NaT
            for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d", "%Y/%m/%d"):
                try:
                    return pd.to_datetime(s, format=fmt)
                except:
                    continue
            try:
                return parse(s, dayfirst=True)
            except:
                return pd.NaT

        df["ACCOUNT_OPEN_DT"] = df["ACCOUNT_OPEN_DT"].apply(clean_date)
        df["COB_DATE"] = df["COB_DATE"].apply(clean_date)

        # Days active
        df["days_active"] = (df["COB_DATE"] - df["ACCOUNT_OPEN_DT"]).dt.days

        # Lower-side breach
        df["turnover_breach_lowerside_extra"] = df["TOTAL_ACTUAL_TO"].fillna(0) - df[
            "ACCOUNT_T_O_WITH_TOLERENCE_NEG"
        ].fillna(0)

        mask = (
            (df["days_active"] > 366)
            & df["TOTAL_ACTUAL_TO"].notna()
            & df["ACCOUNT_T_O_WITH_TOLERENCE_NEG"].notna()
            & (df["turnover_breach_lowerside_extra"] < 0)
        )

        df_out = df.loc[mask].copy()

        # Output columns & rename
        output_cols = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "ACCOUNT_TURNOVER",
            "KYC_ANN_TO_CORPORATE",
            "ACCOUNT_T_O__P_A_KYC_",
            "TOTAL_ACTUAL_TO",
            "ACCOUNT_T_O_WITH_TOLERENCE_NEG",
            "turnover_breach_lowerside_extra",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
            "COB_DATE",
        ]
        df_out = df_out[[c for c in output_cols if c in df_out.columns]]
        df_out.columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "Account_Turnover",
            "KYC_Ann_TO_Corporate",
            "Account_T_O__P_A_KYC_",
            "Total_Actual_TO",
            "Account_T_O_With_Tolerence_Neg",
            "turnover_breach_lowerside_extra",
            "ProductDesc",
            "Sector_Description",
            "COB_Date",
        ]

        if df_out.empty:
            df_out = pd.DataFrame([{col: pd.NA for col in df_out.columns}])

        if mode == "full":
            file_path = f"DEBUG_lowerside_trace_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            df_out.to_excel(file_path, index=False)

        return df_out

    except Exception as e:
        print("❌ ERROR:", e)
        return pd.DataFrame(
            [
                {
                    col: pd.NA
                    for col in [
                        "Customer_Number",
                        "Account_Number",
                        "TitleOfAccount",
                        "Account_Turnover",
                        "KYC_Ann_TO_Corporate",
                        "Account_T_O__P_A_KYC_",
                        "Total_Actual_TO",
                        "Account_T_O_With_Tolerence_Neg",
                        "turnover_breach_lowerside_extra",
                        "ProductDesc",
                        "Sector_Description",
                        "COB_Date",
                    ]
                }
            ]
        )
