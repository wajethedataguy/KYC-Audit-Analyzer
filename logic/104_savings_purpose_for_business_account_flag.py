import pandas as pd
import re
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="104_savings_purpose_for_business_account_flag",
    description="Flags business accounts opened with 'Savings' purpose but used for high-volume business transactions.",
    category="CDD & EDD Review",
)
def logic_104_savings_purpose_for_business_account_flag(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    try:
        # 🔍 Load KYC file (always merged_file.xlsx)
        kyc_key = next((k for k in dataframes if "merged_file.xlsx" in k.lower()), None)
        df_kyc = dataframes.get(kyc_key)
        if isinstance(df_kyc, tuple):
            df_kyc = next(iter(df_kyc[0].values())) if df_kyc[0] else pd.DataFrame()
        if df_kyc is None or df_kyc.empty:
            raise ValueError("KYC file not found or empty")

        # 🔧 Normalize column names
        df_kyc.columns = (
            df_kyc.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )

        # 🔧 Column mapping
        col_map = {
            "ACCOUNT_NUMBER": "Account_Number",
            "TITLEOFACCOUNT": "TitleOfAccount",
            "CUSTOMER_NUMBER": "Customer_Number",
            "CUSTOMERFULLNAME": "CustomerFullName",
            "CUSTSECTORCODE": "CustSectorCode",
            "CUSTOMER_OCCUPATION": "Customer_Occupation",
            "PURPOSE_OF_ACCOUNT": "Purpose_Of_Account",
            "PRODUCTDESC": "ProductDesc",
            "SECTOR_DESCRIPTION": "Sector_Description",
        }
        df_kyc = df_kyc.rename(
            columns={k: v for k, v in col_map.items() if k in df_kyc.columns}
        )

        required = list(col_map.values())
        missing = [c for c in required if c not in df_kyc.columns]
        if missing:
            raise ValueError(f"Missing required KYC columns: {missing}")

        # 🔍 Load NILL_COMBINATIONS reference from uploaded files
        raw_nill_values = set()
        for df_nill in dataframes.values():
            if isinstance(df_nill, dict):
                for sheet in df_nill.values():
                    sheet.columns = (
                        sheet.columns.str.strip()
                        .str.upper()
                        .str.replace(" ", "_")
                        .str.replace("-", "_")
                    )
                    if "NILL_COMBINATIONS" in sheet.columns:
                        raw_nill_values.update(
                            sheet["NILL_COMBINATIONS"]
                            .dropna()
                            .astype(str)
                            .str.strip()
                            .str.lower()
                        )
            elif isinstance(df_nill, pd.DataFrame):
                df_nill.columns = (
                    df_nill.columns.str.strip()
                    .str.upper()
                    .str.replace(" ", "_")
                    .str.replace("-", "_")
                )
                if "NILL_COMBINATIONS" in df_nill.columns:
                    raw_nill_values.update(
                        df_nill["NILL_COMBINATIONS"]
                        .dropna()
                        .astype(str)
                        .str.strip()
                        .str.lower()
                    )

        def normalize(val):
            return re.sub(r"[^\w]+", "", str(val).strip().lower())

        nill_values = set(map(normalize, raw_nill_values))

        # 🔧 Normalize fields
        df_kyc["Account_Number"] = (
            pd.to_numeric(df_kyc["Account_Number"], errors="coerce")
            .astype("Int64")
            .astype(str)
        )
        df_kyc["Purpose_Of_Account_Clean"] = (
            df_kyc["Purpose_Of_Account"].astype(str).map(normalize)
        )
        df_kyc["Customer_Occupation_Clean"] = (
            df_kyc["Customer_Occupation"].astype(str).map(normalize)
        )
        df_kyc["CustSectorCode_Str"] = df_kyc["CustSectorCode"].apply(
            lambda x: str(int(x)) if pd.notna(x) else ""
        )

        # 🔍 Dynamically find FT insight file by column match
        df_ft_raw = None
        for df in dataframes.values():
            if isinstance(df, tuple):
                df = next(iter(df[0].values())) if df[0] else pd.DataFrame()
            if isinstance(df, pd.DataFrame):
                cols = df.columns.str.lower().str.strip()
                if {"debit_account", "cr_account", "transaction_amount"}.issubset(
                    set(cols)
                ):
                    df_ft_raw = df
                    break
        if df_ft_raw is None or df_ft_raw.empty:
            raise ValueError("FT insight file not found or missing required columns")

        df_ft_raw.columns = df_ft_raw.columns.str.strip().str.lower()
        df_ft_raw["debit_account"] = (
            pd.to_numeric(df_ft_raw["debit_account"], errors="coerce")
            .astype("Int64")
            .astype(str)
        )
        df_ft_raw["cr_account"] = (
            pd.to_numeric(df_ft_raw["cr_account"], errors="coerce")
            .astype("Int64")
            .astype(str)
        )

        # 📊 Build FT summary
        debit = df_ft_raw[["debit_account", "transaction_amount"]].rename(
            columns={"debit_account": "ACCOUNT", "transaction_amount": "TR_AMOUNT"}
        )
        credit = df_ft_raw[["cr_account", "transaction_amount"]].rename(
            columns={"cr_account": "ACCOUNT", "transaction_amount": "TR_AMOUNT"}
        )
        combined = pd.concat([debit, credit], ignore_index=True)
        combined["TR_AMOUNT"] = pd.to_numeric(
            combined["TR_AMOUNT"].astype(str).str.replace(r"[^\d.]", "", regex=True),
            errors="coerce",
        ).fillna(0)

        df_ft_summary = combined.groupby("ACCOUNT", as_index=False).agg(
            TOTAL_COUNT=("TR_AMOUNT", "count"), TOTAL_AMOUNT=("TR_AMOUNT", "sum")
        )

        # 🧠 Savings keywords (normalized)
        savings_keywords = {
            "saving",
            "savings",
            "personalsaving",
            "personalsavings",
            "personelsavings",
        }

        # 🧠 Sector codes considered business
        sector_codes = {
            "1000",
            "1100",
            "1110",
            "1111",
            "1120",
            "1121",
            "1122",
            "1140",
            "1000.0",
            "1100.0",
            "1110.0",
            "1111.0",
            "1120.0",
            "1121.0",
            "1122.0",
            "1140.0",
        }

        # Exclude when purpose also hints at business
        business_exclusions = {"busin", "buis", "buss"}

        contradiction_mask = (
            (df_kyc["Customer_Occupation_Clean"] == "business")
            & (df_kyc["CustSectorCode_Str"].isin(sector_codes))
            & df_kyc["Purpose_Of_Account_Clean"].notna()
            & (~df_kyc["Purpose_Of_Account_Clean"].isin(nill_values))
            & df_kyc["Purpose_Of_Account_Clean"].apply(
                lambda val: any(kw in val for kw in savings_keywords)
                and not any(ex in val for ex in business_exclusions)
            )
        )

        df_filtered = df_kyc[contradiction_mask].copy()

        # 🔗 Join with FT summary
        df_joined = pd.merge(
            df_filtered,
            df_ft_summary,
            left_on="Account_Number",
            right_on="ACCOUNT",
            how="left",
        )

        # 📊 Final filter: Total_Count ≥ 30
        df_final = df_joined[df_joined["TOTAL_COUNT"].fillna(0) > 30].copy()

        # 📤 Final output
        output_cols = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "CustomerFullName",
            "CustSectorCode",
            "Customer_Occupation",
            "Purpose_Of_Account",
            "ProductDesc",
            "Sector_Description",
            "TOTAL_COUNT",
            "TOTAL_AMOUNT",
        ]
        output = df_final[output_cols].drop_duplicates().reset_index(drop=True)

        if output.empty:
            output = pd.DataFrame([{col: pd.NA for col in output_cols}])

        if mode == "full":
            file_path = (
                f"saving_purpose_for_business_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            )
            output.to_excel(file_path, index=False)

        return output

    except Exception as e:
        print(f"❌ Error in logic_104_savings_purpose_for_business_account_flag: {e}")
        return pd.DataFrame(
            [
                {
                    "Customer_Number": pd.NA,
                    "Account_Number": pd.NA,
                    "TitleOfAccount": pd.NA,
                    "CustomerFullName": pd.NA,
                    "CustSectorCode": pd.NA,
                    "Customer_Occupation": pd.NA,
                    "Purpose_Of_Account": pd.NA,
                    "ProductDesc": pd.NA,
                    "Sector_Description": pd.NA,
                }
            ]
        )
