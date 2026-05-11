import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="026_missing_source_of_income",
    description="Flags customers with missing or invalid Source of Income where CustSectorCode ≤ 1100.",
    category="Compliance & Screening",
)
def logic_026_missing_source_of_income(dataframes: dict, mode="full") -> pd.DataFrame:
    try:
        # Locate merged file
        merged_key = next(
            (k for k in dataframes if "merged_file.xlsx" in k.lower()), None
        )
        if not merged_key:
            raise ValueError("Merged file not found in input dataframes.")

        df_main = dataframes[merged_key].copy()
        df_main.columns = (
            df_main.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )

        # Validate required columns
        required_columns = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "CUSTSECTORCODE",
            "SOURCEOFINCOME",
        ]
        missing = [col for col in required_columns if col not in df_main.columns]
        if missing:
            raise ValueError(f"Missing required columns in merged file: {missing}")

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

        # Load Nill_Combinations
        nill_set = set()
        for key, df_nill in dataframes.items():
            normalized_cols = df_nill.columns.str.upper().str.strip()
            if "NILL_COMBINATIONS" in normalized_cols.values:
                actual_col = df_nill.columns[normalized_cols == "NILL_COMBINATIONS"][0]
                nill_values = (
                    df_nill[actual_col]
                    .dropna()
                    .astype(str)
                    .str.lower()
                    .str.strip()
                    .str.replace(r"\s+", " ", regex=True)
                    .str.replace(r"[^\w\s]", "", regex=True)
                    .unique()
                )
                nill_set = set(nill_values)
                break

        if not nill_set:
            raise ValueError(
                "Column 'Nill_Combinations' not found in any uploaded file."
            )

        # Normalize and filter
        df_main["SOURCEOFINCOME_CLEAN"] = (
            df_main["SOURCEOFINCOME"]
            .astype(str)
            .str.lower()
            .str.strip()
            .str.replace(r"\s+", " ", regex=True)
            .str.replace(r"[^\w\s]", "", regex=True)
        )
        df_main["CUSTSECTORCODE_NUM"] = pd.to_numeric(
            df_main["CUSTSECTORCODE"], errors="coerce"
        )

        df_filtered = df_main[
            (
                df_main["SOURCEOFINCOME_CLEAN"].isnull()
                | (df_main["SOURCEOFINCOME_CLEAN"] == "")
                | (df_main["SOURCEOFINCOME_CLEAN"].isin(nill_set))
            )
            & (df_main["CUSTSECTORCODE_NUM"] <= 1100)
        ].copy()

        # Prepare output
        output = df_filtered[
            [
                "CUSTOMER_NUMBER",
                "ACCOUNT_NUMBER",
                "TITLEOFACCOUNT",
                "PRODUCTDESC",
                "SECTOR_DESCRIPTION",
            ]
        ].copy()

        output.columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "ProductDesc",
            "SECTOR_DESCRIPTION",
        ]
        output = output.reset_index(drop=True)

        # ✅ Return only headers if no contradictions found
        if output.empty:
            return pd.DataFrame(columns=output.columns)

        # 📁 Optional export
        if mode == "full":
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_path = f"missing_source_of_income_{timestamp}.xlsx"
            output.to_excel(file_path, index=False)

        return output

    except Exception as e:
        print(f"❌ Exception occurred: {e}")
        fallback_columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "ProductDesc",
            "SECTOR_DESCRIPTION",
        ]
        return pd.DataFrame(columns=fallback_columns)
