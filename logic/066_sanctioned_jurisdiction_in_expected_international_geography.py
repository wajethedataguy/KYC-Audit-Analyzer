import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="066_sanctioned_jurisdiction_in_expected_international_geography",
    description="Flags customers whose expected international geography includes a sanctioned or banned jurisdiction.",
    category="Compliance & Screening",
)
def logic_066_sanctioned_jurisdiction_in_expected_international_geography(
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
            "EXPINTERNATIONALGEOGRAPHY",
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

        # 🔍 Normalize geography field
        df["EXPINTERNATIONALGEOGRAPHY_STR"] = (
            df["EXPINTERNATIONALGEOGRAPHY"].astype(str).str.upper()
        )

        # 🧠 Define sanctioned jurisdictions
        sanctioned_list = {
            "RUSSIAN FEDERATION",
            "IRAN",
            "SYRIA",
            "CUBA",
            "NORTH KOREA",
            "BELARUS",
            "MYANMAR",
            "DONETSK",
            "LUHANSK",
            "ZAPORIZHZHIA",
        }

        # 🔍 Match logic: substring containment
        df["MATCH_FOUND"] = df["EXPINTERNATIONALGEOGRAPHY_STR"].apply(
            lambda geo: any(sanc in geo for sanc in sanctioned_list)
        )

        # 🧠 Apply contradiction logic
        df_filtered = df[
            df["EXPINTERNATIONALGEOGRAPHY"].notna()
            & df["MATCH_FOUND"]
            & df["ACCOUNT_NUMBER"].notna()
        ].copy()

        # 📤 Prepare output
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

        # 🧯 Handle empty output
        if output.empty:
            output = pd.DataFrame(
                [
                    {
                        "Customer_Number": pd.NA,
                        "Account_Number": pd.NA,
                        "TitleOfAccount": pd.NA,
                        "ProductDesc": pd.NA,
                        "SECTOR_DESCRIPTION": pd.NA,
                    }
                ]
            )

        # 📁 Optional export
        if mode == "full":
            file_path = f"sanctioned_jurisdiction_in_expected_international_geography_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
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
                    "ProductDesc": pd.NA,
                    "SECTOR_DESCRIPTION": pd.NA,
                }
            ]
        )
