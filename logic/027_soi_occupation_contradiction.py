import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="027_soi_occupation_contradiction",
    description="Flags customers whose Source of Income contradicts their Occupation based on known valid combinations.",
    category="Compliance & Screening",
)
def logic_027_soi_occupation_contradiction(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    try:
        # -----------------------------------------
        # Load KYC profile
        # -----------------------------------------
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

        required = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "CUSTOMER_OCCUPATION",
            "SOURCEOFINCOME",
        ]
        missing = [col for col in required if col not in df_main.columns]
        if missing:
            raise ValueError(f"Missing required columns: {missing}")

        # -----------------------------------------
        # Search fallback files for ProductDesc and Sector_Description
        # -----------------------------------------
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

        # -----------------------------------------
        # Find reference file by columns
        # -----------------------------------------
        ref_key, ref = None, None
        for k, df_other in dataframes.items():
            if not isinstance(df_other, pd.DataFrame):
                continue
            df_other.columns = (
                df_other.columns.str.strip()
                .str.upper()
                .str.replace(" ", "_")
                .str.replace("-", "_")
            )
            if {"PURPOSE", "OCC_PUR_COMBINATION"}.issubset(df_other.columns):
                ref_key, ref = k, df_other.copy()
                break

        if ref is None:
            raise ValueError(
                "Occupation-Purpose combination file not found (no file with required columns)."
            )

        # -----------------------------------------
        # Normalize reference values
        # -----------------------------------------
        purpose_list = (
            ref["PURPOSE"]
            .dropna()
            .astype(str)
            .str.lower()
            .str.replace(r"[^\w\s]", "", regex=True)
            .str.replace(r"\s+", "", regex=True)
            .unique()
            .tolist()
        )
        valid_combos = (
            ref["OCC_PUR_COMBINATION"]
            .dropna()
            .astype(str)
            .str.lower()
            .str.replace(r"[^\w\s]", "", regex=True)
            .str.replace(r"\s+", "", regex=True)
            .unique()
            .tolist()
        )

        # -----------------------------------------
        # Normalize KYC fields
        # -----------------------------------------
        df_main["SOURCEOFINCOME_CLEAN"] = (
            df_main["SOURCEOFINCOME"]
            .astype(str)
            .str.lower()
            .str.replace(r"[^\w\s]", "", regex=True)
            .str.replace(r"\s+", "", regex=True)
        )
        df_main["CUSTOMER_OCCUPATION_CLEAN"] = (
            df_main["CUSTOMER_OCCUPATION"]
            .astype(str)
            .str.lower()
            .str.replace(r"[^\w\s]", "", regex=True)
            .str.replace(r"\s+", "", regex=True)
        )

        # -----------------------------------------
        # Build Matching_SOI
        # -----------------------------------------
        def match_soi(row):
            soi, occ = row["SOURCEOFINCOME_CLEAN"], row["CUSTOMER_OCCUPATION_CLEAN"]
            matches = [p for p in purpose_list if p in soi]
            return occ + matches[0] if matches else soi

        df_main["MATCHING_SOI"] = (
            df_main.apply(match_soi, axis=1)
            .astype(str)
            .str.lower()
            .str.replace(r"\s+", "", regex=True)
        )
        df_main["MATCHING_STATUS"] = df_main["MATCHING_SOI"].apply(
            lambda x: "Matched" if x in valid_combos else "Unmatched"
        )

        # -----------------------------------------
        # Filter matched rows
        # -----------------------------------------
        matched = df_main[df_main["MATCHING_STATUS"] == "Matched"].copy()

        # -----------------------------------------
        # Prepare output
        # -----------------------------------------
        output = matched[
            [
                "CUSTOMER_NUMBER",
                "ACCOUNT_NUMBER",
                "TITLEOFACCOUNT",
                "CUSTOMER_OCCUPATION",
                "SOURCEOFINCOME",
                "PRODUCTDESC",
                "SECTOR_DESCRIPTION",
            ]
        ].copy()

        output.columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "Customer_Occupation",
            "SourceOfIncome",
            "ProductDesc",
            "SECTOR_DESCRIPTION",
        ]
        output = output.reset_index(drop=True)

        # ✅ Return only headers if no contradictions found
        if output.empty:
            return pd.DataFrame(columns=output.columns)

        # 📁 Optional export
        if mode == "full":
            file_path = (
                f"soi_occupation_contradiction_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            )
            output.to_excel(file_path, index=False)

        return output

    except Exception as e:
        print(f"❌ Exception occurred: {e}")
        fallback_columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "Customer_Occupation",
            "SourceOfIncome",
            "ProductDesc",
            "SECTOR_DESCRIPTION",
        ]
        return pd.DataFrame(columns=fallback_columns)
