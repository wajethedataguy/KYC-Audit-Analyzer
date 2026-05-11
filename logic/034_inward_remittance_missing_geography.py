import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="034_inward_remittance_missing_geography",
    description="Flags customers with 'Inward Foreign Remittance' as deposit mode but expected geography is set to Pakistan or PK.",
    category="Compliance & Screening",
)
def logic_034_inward_remittance_missing_geography(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    try:
        merged_key = next(
            (k for k in dataframes if "merged_file.xlsx" in k.lower()), None
        )
        if not merged_key:
            raise ValueError("Merged file not found.")

        df_main = dataframes[merged_key].copy()
        df_main.columns = (
            df_main.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )

        required_cols = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "DOMINANT_MODE_OF_DEPOSIT",
            "EXPINTERNATIONALGEOGRAPHY",
        ]
        missing = [col for col in required_cols if col not in df_main.columns]
        if missing:
            raise ValueError(f"Missing required columns: {missing}")

        # 🔁 Search fallback files for ProductDesc and Sector_Description using Account_Num
        fallback_df = None
        for key, df_other in dataframes.items():
            if key == merged_key:
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
            df_main["ACCOUNT_NUMBER"] = df_main["ACCOUNT_NUMBER"].astype(str)
            fallback_df["ACCOUNT_NUM"] = fallback_df["ACCOUNT_NUM"].astype(str)

            df_main = df_main.merge(
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
                if fallback_col in df_main.columns:
                    df_main[col] = df_main[col].combine_first(df_main[fallback_col])

        # ✅ Normalize and filter
        df_main["DEPOSIT_MODE_CLEAN"] = (
            df_main["DOMINANT_MODE_OF_DEPOSIT"].astype(str).str.lower().str.strip()
        )
        df_main["GEOGRAPHY_CLEAN"] = (
            df_main["EXPINTERNATIONALGEOGRAPHY"].astype(str).str.upper().str.strip()
        )

        df_filtered = df_main[
            df_main["DEPOSIT_MODE_CLEAN"].str.contains(
                "inward foreign remittance", na=False
            )
            & df_main["GEOGRAPHY_CLEAN"].isin(["PAKISTAN", "PK"])
        ].copy()

        output = df_filtered[
            [
                "CUSTOMER_NUMBER",
                "ACCOUNT_NUMBER",
                "TITLEOFACCOUNT",
                "EXPINTERNATIONALGEOGRAPHY",
                "PRODUCTDESC",
                "SECTOR_DESCRIPTION",
            ]
        ].copy()

        output.columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "ExpInternationalGeography",
            "ProductDesc",
            "SECTOR_DESCRIPTION",
        ]
        output = output.reset_index(drop=True)

        if output.empty:
            output = pd.DataFrame(
                [
                    {
                        "Customer_Number": pd.NA,
                        "Account_Number": pd.NA,
                        "TitleOfAccount": pd.NA,
                        "ExpInternationalGeography": pd.NA,
                        "ProductDesc": pd.NA,
                        "SECTOR_DESCRIPTION": pd.NA,
                    }
                ]
            )

        if mode == "full":
            file_path = f"inward_remittance_missing_geography_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            output.to_excel(file_path, index=False)

        return output

    except Exception as e:
        print(f"❌ Exception occurred: {e}")
        fallback_columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "ExpInternationalGeography",
            "ProductDesc",
            "SECTOR_DESCRIPTION",
        ]
        return pd.DataFrame(columns=fallback_columns)
