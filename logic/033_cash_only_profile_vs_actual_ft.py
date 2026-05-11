import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


def fallback_output(reason=""):
    return pd.DataFrame(
        [
            {
                "Customer_Number": pd.NA,
                "Account_Number": pd.NA,
                "TitleOfAccount": pd.NA,
                "Dominant_Mode_of_Deposit": pd.NA,
                "Dominant_Mode_of_Withdrawal": pd.NA,
                "Observed_FT_Transaction_Count": pd.NA,
                "ProductDesc": pd.NA,
                "SECTOR_DESCRIPTION": pd.NA,
            }
        ]
    )


def normalize_account_number(val, pad_to=11):
    return (
        str(val)
        .strip()
        .replace("\u200b", "")
        .replace("\xa0", "")
        .replace(" ", "")
        .replace("-", "")
        .split(".")[0]
        .zfill(pad_to)
    )


def generate_ft_transaction_summary(dataframes):
    insight_key = next(
        (k for k in dataframes if "ftt" in k.lower() and "kyc" in k.lower()), None
    )
    if not insight_key:
        return fallback_output()

    df_raw = dataframes[insight_key].copy()
    df_raw.columns = (
        df_raw.columns.str.strip()
        .str.upper()
        .str.replace(" ", "_")
        .str.replace("-", "_")
    )

    required = {"DEBIT_ACCOUNT", "CR_ACCOUNT", "TRANSACTION_AMOUNT"}
    if not required.issubset(df_raw.columns):
        return fallback_output()

    debit_df = df_raw[["DEBIT_ACCOUNT", "TRANSACTION_AMOUNT"]].rename(
        columns={"DEBIT_ACCOUNT": "ACCOUNT", "TRANSACTION_AMOUNT": "TR_AMOUNT"}
    )
    credit_df = df_raw[["CR_ACCOUNT", "TRANSACTION_AMOUNT"]].rename(
        columns={"CR_ACCOUNT": "ACCOUNT", "TRANSACTION_AMOUNT": "TR_AMOUNT"}
    )
    combined_df = pd.concat([debit_df, credit_df], ignore_index=True)

    combined_df["TR_AMOUNT"] = (
        combined_df["TR_AMOUNT"]
        .astype(str)
        .str.extract(r"([\d\.]+)")
        .fillna("0")
        .astype(float)
    )

    return (
        combined_df.groupby("ACCOUNT", dropna=False)
        .agg(TOTAL_COUNT=("TR_AMOUNT", "count"), TOTAL_AMOUNT=("TR_AMOUNT", "sum"))
        .reset_index()
    )


@register_logic(
    name="033_cash_only_profile_vs_actual_ft",
    description="Flags accounts with cash-only KYC profile but >30 fund transfer transactions observed.",
    category="Compliance & Screening",
)
def logic_033_cash_only_profile_vs_actual_ft(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    try:
        kyc_key = next((k for k in dataframes if "merged" in k.lower()), None)
        if not kyc_key:
            raise ValueError()

        df_kyc = dataframes[kyc_key].copy()
        df_kyc.columns = (
            df_kyc.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )

        required_cols = {
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "DOMINANT_MODE_OF_DEPOSIT",
            "DOMINANT_MODE_OF_WITHDRAWAL",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        }
        if not required_cols.issubset(df_kyc.columns):
            raise ValueError()

        # 🔁 Search fallback files for ProductDesc and Sector_Description using Account_Num
        fallback_df = None
        for key, df_other in dataframes.items():
            if key == kyc_key:
                continue
            df_other.columns = (
                df_other.columns.str.strip()
                .str.upper()
                .str.replace(" ", "_")
                .str.replace("-", "_")
            )
            if "ACCOUNT_NUM" in df_other.columns and (
                "PRODUCTDESC" in df_other.columns
                or "SECTOR_DESCRIPTION" in df_other.columns
            ):
                fallback_df = df_other.copy()
                break

        # 🔁 Merge fallback if found
        if fallback_df is not None:
            df_kyc["ACCOUNT_NUMBER"] = df_kyc["ACCOUNT_NUMBER"].astype(str)
            fallback_df["ACCOUNT_NUM"] = fallback_df["ACCOUNT_NUM"].astype(str)

            df_kyc = df_kyc.merge(
                fallback_df[
                    ["ACCOUNT_NUM"]
                    + [
                        col
                        for col in ["PRODUCTDESC", "SECTOR_DESCRIPTION"]
                        if col in fallback_df.columns
                    ]
                ],
                left_on="ACCOUNT_NUMBER",
                right_on="ACCOUNT_NUM",
                how="left",
                suffixes=("", "_FALLBACK"),
            )

            for col in ["PRODUCTDESC", "SECTOR_DESCRIPTION"]:
                fallback_col = f"{col}_FALLBACK"
                if fallback_col in df_kyc.columns:
                    df_kyc[col] = df_kyc[col].combine_first(df_kyc[fallback_col])

        df_kyc["COMBINED_MODE"] = (
            df_kyc["DOMINANT_MODE_OF_DEPOSIT"].astype(str).str.lower().str.strip()
            + "|"
            + df_kyc["DOMINANT_MODE_OF_WITHDRAWAL"].astype(str).str.lower().str.strip()
        )

        df_filtered = df_kyc[
            df_kyc["COMBINED_MODE"] == "cash|cash withdrawls through cheque"
        ].copy()
        if df_filtered.empty:
            return fallback_output()

        ft_summary = generate_ft_transaction_summary(dataframes)
        if ft_summary.empty or "ACCOUNT" not in ft_summary.columns:
            return fallback_output()

        df_filtered["ACCOUNT_NUMBER"] = df_filtered["ACCOUNT_NUMBER"].apply(
            normalize_account_number
        )
        ft_summary["ACCOUNT"] = ft_summary["ACCOUNT"].apply(normalize_account_number)

        df_merged = df_filtered.merge(
            ft_summary, left_on="ACCOUNT_NUMBER", right_on="ACCOUNT", how="left"
        )
        df_merged["TOTAL_COUNT"] = pd.to_numeric(
            df_merged["TOTAL_COUNT"], errors="coerce"
        )

        df_final = df_merged[df_merged["TOTAL_COUNT"] > 30].copy()
        if df_final.empty:
            return fallback_output()

        output = df_final[
            [
                "CUSTOMER_NUMBER",
                "ACCOUNT_NUMBER",
                "TITLEOFACCOUNT",
                "DOMINANT_MODE_OF_DEPOSIT",
                "DOMINANT_MODE_OF_WITHDRAWAL",
                "TOTAL_COUNT",
                "PRODUCTDESC",
                "SECTOR_DESCRIPTION",
            ]
        ].copy()

        output.columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "Dominant_Mode_of_Deposit",
            "Dominant_Mode_of_Withdrawal",
            "Observed_FT_Transaction_Count",
            "ProductDesc",
            "SECTOR_DESCRIPTION",
        ]
        output = output.reset_index(drop=True)

        if mode == "full":
            file_path = (
                f"cash_only_profile_vs_actual_ft_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            )
            output.to_excel(file_path, index=False)

        return output

    except Exception:
        return fallback_output()
