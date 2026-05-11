import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


def get_empty_output(columns: list) -> pd.DataFrame:
    return pd.DataFrame([{col: pd.NA for col in columns}])


@register_logic(
    name="006_missing_purpose_of_account",
    description="Flags customers whose KYC profile is missing the purpose for opening the account.",
    category="Customer Name Filter",
)
def logic_006_missing_purpose_of_account(
    dataframes: dict, mode: str = "full"
) -> pd.DataFrame:
    try:
        output_columns = [
            "Customer_Number",
            "Account_Number",
            "Purpose_of_Account",
            "TitleOfAccount",
            "ProductDesc",
            "Sector_Description",
        ]

        # --- Locate merged KYC file ---
        merged_key = next(
            (k for k in dataframes if "merged_file.xlsx" in k.lower()), None
        )
        if not merged_key:
            raise ValueError("Merged file not found in input dataframes.")
        df_main = dataframes[merged_key].copy()

        # Normalize column names
        df_main.columns = (
            df_main.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )

        # Ensure required columns exist
        for col in output_columns:
            if col.upper() not in df_main.columns:
                df_main[col.upper()] = ""

        # --- Load Nill_Combinations reference robustly ---
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
                            .str.upper()
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
                        .str.upper()
                    )

        nill_set = {val for val in raw_nill_values}

        if not nill_set:
            print("✅ No nill combinations found.")
            return get_empty_output(output_columns)

        # --- Normalize Purpose_of_Account ---
        df_main["PURPOSE_OF_ACCOUNT"] = (
            df_main["PURPOSE_OF_ACCOUNT"].fillna("").astype(str).str.strip().str.upper()
        )

        # --- Apply contradiction logic ---
        df_flagged = df_main[df_main["PURPOSE_OF_ACCOUNT"].isin(nill_set)].copy()

        # --- Prepare output ---
        output = df_flagged[[col.upper() for col in output_columns]].copy()
        output.columns = output_columns
        output = output.drop_duplicates().reset_index(drop=True)

        if output.empty:
            output = get_empty_output(output_columns)

        if mode == "full":
            file_path = (
                f"missing_purpose_of_account_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            )
            output.to_excel(file_path, index=False)

        return output

    except Exception as e:
        print(f"❌ Error in logic_006_missing_purpose_of_account: {e}")
        return get_empty_output(
            [
                "Customer_Number",
                "Account_Number",
                "Purpose_of_Account",
                "TitleOfAccount",
                "ProductDesc",
                "Sector_Description",
            ]
        )
