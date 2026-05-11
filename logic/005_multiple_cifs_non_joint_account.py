import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


def get_empty_output(columns: list) -> pd.DataFrame:
    return pd.DataFrame([{col: pd.NA for col in columns}])


@register_logic(
    name="005_multiple_cifs_non_joint_account",
    description="Flags personal accounts with multiple CIFs but no joint indicators in title, matching Power Query logic.",
    category="Customer Integrity",
)
def logic_005_multiple_cifs_non_joint_account(
    dataframes: dict, mode: str = "full"
) -> pd.DataFrame:
    try:
        # 🔍 Locate merged file
        kyc_key = next((k for k in dataframes if "merged_file.xlsx" in k.lower()), None)
        df = dataframes.get(kyc_key)
        if isinstance(df, tuple):
            df = next(iter(df[0].values())) if df[0] else pd.DataFrame()
        if df is None or df.empty:
            raise ValueError("Merged file not found or empty.")

        # 🔧 Normalize column names
        df.columns = (
            df.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )

        # ✅ Required columns
        required_columns = [
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "CUSTOMER_NUMBER",
            "CUSTOMERFULLNAME",
            "JOINT_HOLDER",
            "CUSTSECTORCODE",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        missing = [col for col in required_columns if col not in df.columns]
        if missing:
            raise ValueError(f"Missing required columns: {missing}")

        # 🔧 Normalize fields
        df["CUSTSECTORCODE"] = pd.to_numeric(df["CUSTSECTORCODE"], errors="coerce")
        df["TITLEOFACCOUNT"] = df["TITLEOFACCOUNT"].fillna("").astype(str).str.strip()
        df["JOINT_HOLDER"] = df["JOINT_HOLDER"].fillna("").astype(str).str.strip()
        invalid_joint_holders = {"", "0", "0.0", "0.00"}

        # ❌ Exclusion patterns (must NOT be present in title)
        exclusion_patterns = ["/", "\\", " \\", " AND ", " OR ", " & ", "& "]

        def is_clean_title(title: str) -> bool:
            return all(p not in title for p in exclusion_patterns)

        # 🧠 Apply contradiction logic
        df_flagged = df[
            (df["CUSTSECTORCODE"] <= 1005)
            & (~df["JOINT_HOLDER"].isin(invalid_joint_holders))
            & df["TITLEOFACCOUNT"].apply(is_clean_title)
        ].copy()

        # 🧼 Deduplicate to match Power Query behavior
        df_flagged = df_flagged.drop_duplicates(
            subset=["ACCOUNT_NUMBER", "CUSTOMER_NUMBER"]
        )

        # 📤 Prepare output
        output_columns = [
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "CUSTOMER_NUMBER",
            "CUSTOMERFULLNAME",
            "JOINT_HOLDER",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        output = df_flagged[output_columns].drop_duplicates().reset_index(drop=True)

        # 🧯 Fallback for empty output
        if output.empty:
            output = get_empty_output(output_columns)

        # 📁 Controlled export
        if mode == "full":
            file_path = (
                f"multiple_cifs_non_joint_account_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            )
            output.to_excel(file_path, index=False)
        elif mode == "ui":
            output = output[
                [
                    "ACCOUNT_NUMBER",
                    "TITLEOFACCOUNT",
                    "CUSTOMER_NUMBER",
                    "CUSTOMERFULLNAME",
                    "JOINT_HOLDER",
                    "PRODUCTDESC",
                    "SECTOR_DESCRIPTION",
                ]
            ]

        return output

    except Exception as e:
        print(f"❌ Error in logic_005_multiple_cifs_non_joint_account: {e}")
        return get_empty_output(
            [
                "ACCOUNT_NUMBER",
                "TITLEOFACCOUNT",
                "CUSTOMER_NUMBER",
                "CUSTOMERFULLNAME",
                "JOINT_HOLDER",
                "PRODUCTDESC",
                "SECTOR_DESCRIPTION",
            ]
        )
