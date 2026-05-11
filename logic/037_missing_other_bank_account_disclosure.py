import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="037_missing_other_bank_account_disclosure",
    description="Flags high-risk customers with RTGS outflows to their own accounts in other banks, based only on account number and title match.",
    category="Compliance & Screening",
)
def logic_037_missing_other_bank_account_disclosure(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    try:
        # 🔍 Identify RTGS file by signature columns
        required_rtgs_cols = [
            "ORDERING_ACCOUNT",
            "ORDERING_TITLE",
            "BENEFICIARY_TITLE",
            "BENEFICIARY_ACCOUNT",
        ]
        df_rtgs = None
        for key, df in dataframes.items():
            if not isinstance(df, pd.DataFrame):
                continue
            df_temp = df.copy()
            df_temp.columns = (
                df_temp.columns.str.strip()
                .str.upper()
                .str.replace(" ", "_")
                .str.replace("-", "_")
            )
            if all(col in df_temp.columns for col in required_rtgs_cols):
                df_rtgs = df_temp
                break

        if df_rtgs is None:
            raise ValueError("No file contains all required RTGS columns.")

        # ✅ Normalize RTGS fields
        df_rtgs["ORDERING_ACCOUNT_NORM"] = (
            df_rtgs["ORDERING_ACCOUNT"]
            .astype(str)
            .str.replace("'", "")  # remove apostrophe
            .str.strip()
            .str[-10:]  # take last 10 digits
        )
        df_rtgs["ORDERING_TITLE_CLEAN"] = (
            df_rtgs["ORDERING_TITLE"].astype(str).str.strip()
        )
        df_rtgs["BENEFICIARY_TITLE_CLEAN"] = (
            df_rtgs["BENEFICIARY_TITLE"].astype(str).str.strip()
        )

        # ✅ Ensure BENEFICIARY_ACCOUNT is distinct
        df_rtgs = df_rtgs.drop_duplicates(subset=["BENEFICIARY_ACCOUNT"])

        # 🔍 Load merged KYC file (fixed name)
        kyc_key = next((k for k in dataframes if "merged_file.xlsx" in k.lower()), None)
        if not kyc_key:
            raise ValueError("Merged KYC file not found.")
        df_kyc = dataframes[kyc_key].copy()
        df_kyc.columns = (
            df_kyc.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )

        required_kyc_cols = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "CCY",
            "KYCRISK",
            "OTHER_BANK_ACCOUNT",
        ]
        missing_kyc = [c for c in required_kyc_cols if c not in df_kyc.columns]
        if missing_kyc:
            raise ValueError(f"Merged KYC file missing required columns: {missing_kyc}")

        # ✅ Normalize KYC fields
        df_kyc["ACCOUNT_NUMBER_NORM"] = (
            df_kyc["ACCOUNT_NUMBER"].astype(str).str.strip().str[-10:]
        )
        df_kyc["TITLEOFACCOUNT_CLEAN"] = (
            df_kyc["TITLEOFACCOUNT"].astype(str).str.strip()
        )
        df_kyc["KYCRISK_CLEAN"] = df_kyc["KYCRISK"].astype(str).str.strip().str.upper()

        # ✅ Filter only High risk customers
#        df_kyc_high = df_kyc[df_kyc["KYCRISK_CLEAN"] == "HIGH"].copy()
        df_kyc_high = df_kyc[
            df_kyc["KYCRISK_CLEAN"]
            .astype(str)
            .str.upper()
            .str.strip()
            .str.contains(r"\bHIGH\b", regex=True)
        ].copy()
        
        # 🔗 Join strictly on account number + title
        df_joined = pd.merge(
            df_kyc_high,
            df_rtgs,
            left_on=["ACCOUNT_NUMBER_NORM", "TITLEOFACCOUNT_CLEAN"],
            right_on=["ORDERING_ACCOUNT_NORM", "ORDERING_TITLE_CLEAN"],
            how="inner",
        )

        # ✅ Apply extra condition: titles must all match
        df_matched = df_joined[
            (
                df_joined["TITLEOFACCOUNT"].astype(str).str.strip().str.upper()
                == df_joined["ORDERING_TITLE"].astype(str).str.strip().str.upper()
            )
            & (
                df_joined["TITLEOFACCOUNT"].astype(str).str.strip().str.upper()
                == df_joined["BENEFICIARY_TITLE"].astype(str).str.strip().str.upper()
            )
        ].copy()

        if df_matched.empty:
            output = pd.DataFrame(
                [
                    {
                        "CUSTOMER_NUMBER": pd.NA,
                        "ACCOUNT_NUMBER": pd.NA,
                        "TITLEOFACCOUNT": pd.NA,
                        "OTHER_BANK_ACCOUNT": pd.NA,
                        "ORDERING_TITLE": pd.NA,
                        "BENEFICIARY_TITLE": pd.NA,
                        "BENEFICIARY_ACCOUNT": pd.NA,
                        "BANK": pd.NA,
                    }
                ]
            )
        else:
            # ✅ Select desired output columns (keep original names)
            output = df_matched[
                [
                    "CUSTOMER_NUMBER",
                    "ACCOUNT_NUMBER",  # BAFL account number
                    "TITLEOFACCOUNT",  # BAFL account title
                    "OTHER_BANK_ACCOUNT",  # from merged KYC file
                    "ORDERING_TITLE",  # RTGS ordering title
                    "BENEFICIARY_TITLE",  # RTGS beneficiary title
                    "BENEFICIARY_ACCOUNT",  # RTGS beneficiary account
                    "BANK",  # RTGS bank name
                ]
            ].copy()

        # 📁 Optional export
        if mode == "full":
            file_path = f"missing_other_bank_account_disclosure_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            output.to_excel(file_path, index=False)
            print(f"📁 Saved output to: {file_path}")

        return output

    except Exception as e:
        print(f"❌ Exception occurred: {e}")
        return pd.DataFrame(
            [
                {
                    "CUSTOMER_NUMBER": pd.NA,
                    "ACCOUNT_NUMBER": pd.NA,
                    "TITLEOFACCOUNT": pd.NA,
                    "OTHER_BANK_ACCOUNT": pd.NA,
                    "ORDERING_TITLE": pd.NA,
                    "BENEFICIARY_TITLE": pd.NA,
                    "BENEFICIARY_ACCOUNT": pd.NA,
                    "BANK": pd.NA,
                }
            ]
        )
