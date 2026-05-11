import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


def normalize_text(value: str) -> str:
    return str(value).strip().upper().replace(".", "").replace("/", "").replace(" ", "")


@register_logic(
    name="061_investment_fed_for_non_business_profiles",
    description="Flags non-business customers where InvestmentInBusiness is fed with a positive value.",
    category="Compliance & Screening",
)
def logic_061_investment_fed_for_non_business_profiles(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    try:
        # 🔍 Locate merged KYC file
        kyc_key = next((k for k in dataframes if "merged_file.xlsx" in k.lower()), None)
        if not kyc_key:
            raise ValueError("KYC profile file not found.")
        df_main = dataframes[kyc_key].copy()
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
            "CUSTOMERFULLNAME",
            "CUSTSECTORCODE",
            "CUSTOMER_OCCUPATION",
            "INVESTMENTINBUSINESS",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        missing = [col for col in required_cols if col not in df_main.columns]
        if missing:
            raise ValueError(f"Merged file missing required columns: {missing}")

        df = df_main[required_cols].copy()

        # 🔁 Search fallback files for ProductDesc and SECTOR_DESCRIPTION using Account_Num
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

        if fallback_df is not None:
            df["ACCOUNT_NUMBER"] = df["ACCOUNT_NUMBER"].astype(str)
            fallback_df["ACCOUNT_NUM"] = fallback_df["ACCOUNT_NUM"].astype(str)

            df = df.merge(
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
                if fallback_col in df.columns:
                    df[col] = df[col].combine_first(df[fallback_col])

        # 🔍 Normalize fields
        df["CUSTSECTORCODE_STR"] = (
            pd.to_numeric(df["CUSTSECTORCODE"], errors="coerce")
            .astype("Int64")
            .astype(str)
        )
        df["CUSTOMER_OCCUPATION_CLEAN"] = (
            df["CUSTOMER_OCCUPATION"].astype(str).str.strip()
        )
        df["CUSTOMER_OCCUPATION_NORMALIZED"] = df["CUSTOMER_OCCUPATION_CLEAN"].apply(
            normalize_text
        )
        df["INVESTMENTINBUSINESS_STR"] = (
            df["INVESTMENTINBUSINESS"].astype(str).str.strip()
        )
        df["INVESTMENTINBUSINESS_NORMALIZED"] = df["INVESTMENTINBUSINESS_STR"].apply(
            normalize_text
        )

        def parse_investment(val):
            try:
                return float(val)
            except:
                return None

        df["INVESTMENTINBUSINESS_NUMERIC"] = df["INVESTMENTINBUSINESS_STR"].apply(
            parse_investment
        )

        # 🔍 Load NILL_COMBINATIONS reference
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
                            sheet["NILL_COMBINATIONS"].dropna().astype(str).str.strip()
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
                        df_nill["NILL_COMBINATIONS"].dropna().astype(str).str.strip()
                    )

        nill_set_normalized = {normalize_text(val) for val in raw_nill_values}

        # 🧠 Define valid (sector, occupation) pairs
        sec_occu_pairs = {
            ("1000", "SALARIED"),
            ("1000", "RETIRED"),
            ("1000", "HOUSEWIFE"),
            ("1000", "STUDENT"),
            ("1005", "SALARIED"),
            ("1006", "MINOR"),
        }

        # 🧠 Apply contradiction logic
        df_filtered = df[
            (df["INVESTMENTINBUSINESS_NUMERIC"].notna())
            & (df["INVESTMENTINBUSINESS_NUMERIC"] > 0)
            & (~df["INVESTMENTINBUSINESS_NORMALIZED"].isin(nill_set_normalized))
            & df.apply(
                lambda row: (
                    row["CUSTSECTORCODE_STR"],
                    normalize_text(row["CUSTOMER_OCCUPATION_CLEAN"]),
                )
                in sec_occu_pairs,
                axis=1,
            )
        ].copy()

        # 📤 Prepare output
        output = df_filtered[
            [
                "CUSTOMER_NUMBER",
                "ACCOUNT_NUMBER",
                "TITLEOFACCOUNT",
                "CUSTOMER_OCCUPATION",
                "INVESTMENTINBUSINESS",
                "PRODUCTDESC",
                "SECTOR_DESCRIPTION",
            ]
        ].copy()

        output.columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "Customer_Occupation",
            "InvestmentInBusiness",
            "ProductDesc",
            "SECTOR_DESCRIPTION",
        ]
        output = output.reset_index(drop=True)

        # 🧯 Handle empty output
        if output.empty:
            output = pd.DataFrame(
                [
                    {
                        "Customer_Number": pd.NA,
                        "Account_Number": pd.NA,
                        "TitleOfAccount": pd.NA,
                        "Customer_Occupation": pd.NA,
                        "InvestmentInBusiness": pd.NA,
                        "ProductDesc": pd.NA,
                        "SECTOR_DESCRIPTION": pd.NA,
                    }
                ]
            )

        # 📁 Optional export
        if mode == "full":
            file_path = (
                f"investment_fed_for_non_business_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            )
            output.to_excel(file_path, index=False)

        return output

    except Exception as e:
        print(f"❌ Exception occurred: {e}")
        return pd.DataFrame(
            [
                {
                    "Customer_Number": pd.NA,
                    "Account_Number": pd.NA,
                    "TitleOfAccount": pd.NA,
                    "Customer_Occupation": pd.NA,
                    "InvestmentInBusiness": pd.NA,
                    "ProductDesc": pd.NA,
                    "SECTOR_DESCRIPTION": pd.NA,
                }
            ]
        )
