import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


def resolve_column_if_missing(
    df_main: pd.DataFrame, dataframes: dict, column_name: str
) -> pd.Series:
    normalized_name = column_name.strip().upper().replace(" ", "_").replace("-", "_")
    df_main.columns = (
        df_main.columns.str.strip()
        .str.upper()
        .str.replace(" ", "_")
        .str.replace("-", "_")
    )
    if normalized_name in df_main.columns:
        return df_main[normalized_name].copy()

    for df in dataframes.values():
        if isinstance(df, dict):
            for sheet_df in df.values():
                sheet_df.columns = (
                    sheet_df.columns.str.strip()
                    .str.upper()
                    .str.replace(" ", "_")
                    .str.replace("-", "_")
                )
                if normalized_name in sheet_df.columns:
                    return sheet_df[normalized_name].copy()
        else:
            df.columns = (
                df.columns.str.strip()
                .str.upper()
                .str.replace(" ", "_")
                .str.replace("-", "_")
            )
            if normalized_name in df.columns:
                return df[normalized_name].copy()

    raise ValueError(f"Column '{column_name}' not found in any file.")


@register_logic(
    name="047_missing_ownership_status_for_business_individual",
    description="Flags business individual customers (sector code 1000) whose StatusOfOwnership is missing or vague.",
    category="Compliance & Screening",
)
def logic_047_missing_ownership_status_for_business_individual(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    try:
        # 🔍 Locate merged KYC file
        kyc_key = next((k for k in dataframes if "merged_file.xlsx" in k.lower()), None)
        if not kyc_key:
            raise ValueError("KYC profile file not found.")
        df_main = dataframes[kyc_key].copy()

        # 🔧 Normalize column names
        df_main.columns = (
            df_main.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )

        # ✅ Required columns (added PRODUCTDESC and SECTOR_DESCRIPTION)
        required_cols = [
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "CUSTOMER_NUMBER",
            "CUSTOMERFULLNAME",
            "CUSTSECTORCODE",
            "CUSTOMER_OCCUPATION",
            "STATUSOFOWNERSHIP",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        resolved = {
            col: resolve_column_if_missing(df_main, dataframes, col)
            for col in required_cols
        }
        df = pd.DataFrame(resolved)

        # 🔍 Load Nill_Combinations reference
        nill_list = set()
        for df_nill in dataframes.values():
            if isinstance(df_nill, dict):
                for sheet_df in df_nill.values():
                    sheet_df.columns = (
                        sheet_df.columns.str.strip()
                        .str.upper()
                        .str.replace(" ", "_")
                        .str.replace("-", "_")
                    )
                    if "NILL_COMBINATIONS" in sheet_df.columns:
                        nill_list.update(
                            sheet_df["NILL_COMBINATIONS"]
                            .dropna()
                            .astype(str)
                            .str.strip()
                            .str.upper()
                            .unique()
                        )
            else:
                df_nill.columns = (
                    df_nill.columns.str.strip()
                    .str.upper()
                    .str.replace(" ", "_")
                    .str.replace("-", "_")
                )
                if "NILL_COMBINATIONS" in df_nill.columns:
                    nill_list.update(
                        df_nill["NILL_COMBINATIONS"]
                        .dropna()
                        .astype(str)
                        .str.strip()
                        .str.upper()
                        .unique()
                    )

        # 🔍 Normalize fields
        df["STATUSOFOWNERSHIP"] = (
            df["STATUSOFOWNERSHIP"]
            .astype(str)
            .str.strip()
            .str.upper()
            .replace("NAN", "")
        )
        df["CUSTOMER_OCCUPATION"] = (
            df["CUSTOMER_OCCUPATION"].astype(str).str.strip().str.upper()
        )
        df["CUSTSECTORCODE"] = pd.to_numeric(df["CUSTSECTORCODE"], errors="coerce")

        # 🧠 Apply contradiction logic
        sector_mask = df["CUSTSECTORCODE"] == 1000
        occupation_mask = df["CUSTOMER_OCCUPATION"] == "BUSINESS"
        ownership_mask = (
            df["STATUSOFOWNERSHIP"].isin(nill_list)
            | df["STATUSOFOWNERSHIP"].isna()
            | (df["STATUSOFOWNERSHIP"] == "")
        )

        df_flagged = df[sector_mask & occupation_mask & ownership_mask].copy()

        # 📤 Prepare output (added PRODUCTDESC and SECTOR_DESCRIPTION)
        output = df_flagged[
            [
                "ACCOUNT_NUMBER",
                "TITLEOFACCOUNT",
                "CUSTOMER_NUMBER",
                "CUSTOMERFULLNAME",
                "CUSTSECTORCODE",
                "CUSTOMER_OCCUPATION",
                "STATUSOFOWNERSHIP",
                "PRODUCTDESC",
                "SECTOR_DESCRIPTION",
            ]
        ].reset_index(drop=True)

        # 🧯 Handle empty output
        if output.empty:
            output = pd.DataFrame(
                [
                    {
                        "ACCOUNT_NUMBER": pd.NA,
                        "TITLEOFACCOUNT": pd.NA,
                        "CUSTOMER_NUMBER": pd.NA,
                        "CUSTOMERFULLNAME": pd.NA,
                        "CUSTSECTORCODE": pd.NA,
                        "CUSTOMER_OCCUPATION": pd.NA,
                        "STATUSOFOWNERSHIP": pd.NA,
                        "PRODUCTDESC": pd.NA,
                        "SECTOR_DESCRIPTION": pd.NA,
                    }
                ]
            )

        # 📁 Optional export
        if mode == "full":
            file_path = f"missing_ownership_status_for_business_individual_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            output.to_excel(file_path, index=False)

        return output

    except Exception as e:
        print(f"❌ Exception occurred: {e}")
        return pd.DataFrame(
            [
                {
                    "ACCOUNT_NUMBER": pd.NA,
                    "TITLEOFACCOUNT": pd.NA,
                    "CUSTOMER_NUMBER": pd.NA,
                    "CUSTOMERFULLNAME": pd.NA,
                    "CUSTSECTORCODE": pd.NA,
                    "CUSTOMER_OCCUPATION": pd.NA,
                    "STATUSOFOWNERSHIP": pd.NA,
                    "PRODUCTDESC": pd.NA,
                    "SECTOR_DESCRIPTION": pd.NA,
                }
            ]
        )
