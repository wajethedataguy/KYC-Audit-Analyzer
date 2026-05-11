import pandas as pd
from KYC_Viewer.utils import parse_date


def normalize(col):
    return col.strip().replace(" ", "").lower()


def detect_turnover_breach(kyc_df, exceptional_df):
    # ✅ Use normalized column names
    merged = kyc_df.merge(
        exceptional_df,
        left_on="account_num",
        right_on="account_num",
        how="left",
    ).reset_index(
        drop=True
    )  # ✅ Fix: ensure unique index

    numeric_cols = [
        "account_turnover_in_numbers",
        "account_t_o_with_tolerence",
        "account_t_o_with_tolerence_neg",
        "total_actual_to",
        "actual_debit_to",
        "actual_credit_to",
    ]
    for col in numeric_cols:
        if col in merged.columns:
            merged[col] = pd.to_numeric(merged[col], errors="coerce")

    for date_col in ["cob_date", "account_open_dt"]:
        if date_col in merged.columns:
            merged[date_col] = merged[date_col].apply(parse_date)

    def filter_row(row):
        try:
            days = (row["cob_date"] - row["account_open_dt"]).days
        except:
            return False
        return (
            pd.notna(days)
            and days > 366
            and pd.notna(row.get("account_turnover"))
            and pd.notna(row.get("account_t_o__p_a_kyc_"))
            and pd.notna(row.get("total_actual_to"))
            and pd.notna(row.get("account_t_o_with_tolerence_neg"))
            and (row["total_actual_to"] - row["account_t_o_with_tolerence_neg"]) < 0
        )

    filtered = merged[merged.apply(filter_row, axis=1)].copy()

    final = filtered[
        [
            "customer_num",
            "account_num",
            "titleofaccount",
            "total_actual_to",
            "account_t_o_with_tolerence_neg",
        ]
    ].copy()

    final["turnover_breach_lowerside"] = (
        final["total_actual_to"] - final["account_t_o_with_tolerence_neg"]
    )

    return final
