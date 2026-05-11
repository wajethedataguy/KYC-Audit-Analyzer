import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="083_high_turnover_mismatch_with_declared_profile",
    description="Flags high-turnover accounts where declared occupation suggests non-business profile.",
    category="CDD & EDD Review",
)
def logic_083_high_turnover_mismatch_with_declared_profile(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    try:
        # 🔍 Load merged KYC file
        kyc_key = next((k for k in dataframes if "merged_file.xlsx" in k.lower()), None)
        df_kyc = dataframes.get(kyc_key)
        if isinstance(df_kyc, tuple):
            df_kyc = next(iter(df_kyc[0].values())) if df_kyc[0] else pd.DataFrame()
        if df_kyc is None or df_kyc.empty:
            raise ValueError("merged_file.xlsx not found or empty.")

        # 🔧 Normalize column names
        df_kyc.columns = (
            df_kyc.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )

        # 🔧 Normalize account numbers consistently
        df_kyc["ACCOUNT_NUMBER_RAW"] = df_kyc["ACCOUNT_NUMBER"]
        df_kyc["ACCOUNT_NUMBER"] = (
            df_kyc["ACCOUNT_NUMBER_RAW"].astype(str).str.strip().str.lstrip("0")
        )

        # 📊 Collect turnover fragments
        turnover_frames = []
        for k, df in dataframes.items():
            if isinstance(df, tuple):
                df = next(iter(df[0].values())) if df[0] else pd.DataFrame()
            if df.empty:
                continue
            df.columns = (
                df.columns.str.strip()
                .str.upper()
                .str.replace(" ", "_")
                .str.replace("-", "_")
            )
            if "ACCOUNT_NUM" in df.columns and "TOTAL_TURNOVER" in df.columns:
                temp = df[["ACCOUNT_NUM", "TOTAL_TURNOVER"]].copy()
                temp["ACCOUNT_NUM"] = (
                    temp["ACCOUNT_NUM"].astype(str).str.strip().str.lstrip("0")
                )
                temp["TOTAL_TURNOVER"] = pd.to_numeric(
                    temp["TOTAL_TURNOVER"], errors="coerce"
                )
                turnover_frames.append(temp)

        df_turnover = (
            pd.concat(turnover_frames, ignore_index=True)
            if turnover_frames
            else pd.DataFrame(columns=["ACCOUNT_NUM", "TOTAL_TURNOVER"])
        )
        df_turnover_grouped = df_turnover.groupby("ACCOUNT_NUM", as_index=False).agg(
            {"TOTAL_TURNOVER": "sum"}
        )

        # 📊 Collect transactional fragments
        insight_frames = []
        for k, df in dataframes.items():
            if isinstance(df, tuple):
                df = next(iter(df[0].values())) if df[0] else pd.DataFrame()
            if df.empty:
                continue
            df.columns = (
                df.columns.str.strip()
                .str.upper()
                .str.replace(" ", "_")
                .str.replace("-", "_")
            )

            # ✅ Flexible detection of date column
            if (
                "DEBIT_ACCOUNT" in df.columns
                and "CR_ACCOUNT" in df.columns
                and "TRANSACTION_AMOUNT" in df.columns
                and (
                    "TRANSACTION_DATE" in df.columns
                    or "DT_OF_TRANSACTION" in df.columns
                )
            ):
                date_col = (
                    "TRANSACTION_DATE"
                    if "TRANSACTION_DATE" in df.columns
                    else "DT_OF_TRANSACTION"
                )

                df_debit = df[["DEBIT_ACCOUNT", "TRANSACTION_AMOUNT", date_col]].rename(
                    columns={
                        "DEBIT_ACCOUNT": "ACCOUNT",
                        "TRANSACTION_AMOUNT": "TR_AMOUNT",
                        date_col: "TRANSACTION_DATE",
                    }
                )
                df_credit = df[["CR_ACCOUNT", "TRANSACTION_AMOUNT", date_col]].rename(
                    columns={
                        "CR_ACCOUNT": "ACCOUNT",
                        "TRANSACTION_AMOUNT": "TR_AMOUNT",
                        date_col: "TRANSACTION_DATE",
                    }
                )
                combined = pd.concat([df_debit, df_credit], ignore_index=True)

                combined["ACCOUNT"] = (
                    combined["ACCOUNT"].astype(str).str.strip().str.lstrip("0")
                )
                combined["TR_AMOUNT"] = pd.to_numeric(
                    combined["TR_AMOUNT"]
                    .astype(str)
                    .str.replace(r"[^\d.]", "", regex=True),
                    errors="coerce",
                )
                combined["TXN_ID"] = (
                    combined["TRANSACTION_DATE"].astype(str).str.strip()
                    + "_"
                    + combined["TR_AMOUNT"].astype(str).str.strip()
                )

                insight_frames.append(combined)

        df_ft_summary = (
            pd.concat(insight_frames, ignore_index=True)
            if insight_frames
            else pd.DataFrame(
                columns=["ACCOUNT", "TR_AMOUNT", "TRANSACTION_DATE", "TXN_ID"]
            )
        )

        if not df_ft_summary.empty:
            df_ft_summary = df_ft_summary.drop_duplicates(subset=["ACCOUNT", "TXN_ID"])
            df_ft_summary = df_ft_summary.groupby("ACCOUNT", as_index=False).agg(
                TOTAL_COUNT=("TXN_ID", "count"), TOTAL_AMOUNT=("TR_AMOUNT", "sum")
            )
        else:
            df_ft_summary = pd.DataFrame(
                columns=["ACCOUNT", "TOTAL_COUNT", "TOTAL_AMOUNT"]
            )

        # 🔗 Join all sources
        df_merged = pd.merge(
            df_kyc,
            df_turnover_grouped,
            left_on="ACCOUNT_NUMBER",
            right_on="ACCOUNT_NUM",
            how="left",
        )
        df_merged = pd.merge(
            df_merged,
            df_ft_summary,
            left_on="ACCOUNT_NUMBER",
            right_on="ACCOUNT",
            how="left",
        )

        # ✅ Ensure numeric defaults for transaction summaries
        df_merged["TOTAL_COUNT"] = df_merged["TOTAL_COUNT"].fillna(0).astype(int)
        df_merged["TOTAL_AMOUNT"] = df_merged["TOTAL_AMOUNT"].fillna(0).astype(float)

        # 🔍 Normalize fields
        df_merged["CUSTSECTORCODE_CLEAN"] = pd.to_numeric(
            df_merged["CUSTSECTORCODE"], errors="coerce"
        )
        df_merged["CUSTOMER_OCCUPATION_CLEAN"] = (
            df_merged["CUSTOMER_OCCUPATION"].astype(str).str.strip().str.lower()
        )
        df_merged["TOTAL_TURNOVER_CLEAN"] = pd.to_numeric(
            df_merged["TOTAL_TURNOVER"], errors="coerce"
        )

        # 🧠 Apply contradiction logic
        excluded_keywords = ["business", "landlord"]
        df_merged["IS_BUSINESS_LIKE"] = df_merged["CUSTOMER_OCCUPATION_CLEAN"].apply(
            lambda x: any(keyword in x for keyword in excluded_keywords)
        )

        contradiction_mask = (
            (~df_merged["IS_BUSINESS_LIKE"])
            & (df_merged["CUSTSECTORCODE_CLEAN"] < 1100)
            & (
                (df_merged["TOTAL_TURNOVER_CLEAN"] > 50000000)
                | (df_merged["TOTAL_COUNT"] > 30)
            )
            & (df_merged["CUSTOMER_OCCUPATION_CLEAN"] != "")
            & (~df_merged["CUSTOMER_OCCUPATION"].isna())
        )

        # 📤 Final output
        output_cols = [
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "CUSTOMER_NUMBER",
            "CUSTSECTORCODE",
            "CUSTOMER_OCCUPATION",
            "TOTAL_TURNOVER",
            "TOTAL_COUNT",
            "TOTAL_AMOUNT",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        output = (
            df_merged[contradiction_mask][output_cols]
            .drop_duplicates()
            .reset_index(drop=True)
        )

        if output.empty:
            output = pd.DataFrame([{col: pd.NA for col in output_cols}])

        return output

    except Exception as e:
        print(
            f"❌ Error in logic_083_high_turnover_mismatch_with_declared_profile: {e}"
        )
        return pd.DataFrame(
            [
                {
                    col: pd.NA
                    for col in [
                        "ACCOUNT_NUMBER",
                        "TITLEOFACCOUNT",
                        "CUSTOMER_NUMBER",
                        "CUSTSECTORCODE",
                        "CUSTOMER_OCCUPATION",
                        "TOTAL_TURNOVER",
                        "TOTAL_COUNT",
                        "TOTAL_AMOUNT",
                        "PRODUCTDESC",
                        "SECTOR_DESCRIPTION",
                    ]
                }
            ]
        )
