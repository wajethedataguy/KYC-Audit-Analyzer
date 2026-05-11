import pandas as pd
import re
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="007_occupation_purpose_mismatch",
    description="Flags accounts where occupation and purpose of account contradict each other.",
    category="Purpose & Occupation Filter",
)
def logic_007_occupation_purpose_mismatch(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    try:
        required_columns = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "PURPOSE_OF_ACCOUNT",
            "TITLEOFACCOUNT",
            "CUSTOMER_OCCUPATION",
            "NAMEOFBUSINESS",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]

        # 🔍 Load KYC data
        kyc_key = next((k for k in dataframes if "merged_file.xlsx" in k.lower()), None)
        if kyc_key is None:
            raise ValueError("Merged file not found.")
        df = dataframes[kyc_key].copy()

        # 🔧 Normalize column names
        df.columns = (
            df.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )

        # ✅ Validate required columns
        missing = [col for col in required_columns if col not in df.columns]
        if missing:
            raise ValueError(f"Missing required columns: {missing}")

        df = df[required_columns].copy()

        # 🔧 Normalize fields
        df["PURPOSE_OF_ACCOUNT"] = (
            df["PURPOSE_OF_ACCOUNT"].astype(str).str.lower().str.strip()
        )
        df["CUSTOMER_OCCUPATION"] = (
            df["CUSTOMER_OCCUPATION"].astype(str).str.lower().str.strip()
        )
        df["TITLEOFACCOUNT"] = df["TITLEOFACCOUNT"].astype(str).str.upper().str.strip()
        df["NAMEOFBUSINESS"] = df["NAMEOFBUSINESS"].astype(str).str.lower().str.strip()

        # 🧠 Define mismatch rules
        mismatch_rules = {
            "business": ["sal", "salar", "salary"],
            "salaried": ["busin", "buis", "buss"],
            "retired": ["busin", "buis", "buss", "sal", "salar", "salary"],
            "landlord": ["sal", "salar", "salary", "retir", "stud"],
            "student": [
                "busin",
                "buis",
                "buss",
                "sal",
                "salar",
                "salary",
                "retir",
                "house",
            ],
            "house wife": [
                "busin",
                "buis",
                "buss",
                "sal",
                "salar",
                "salary",
                "retir",
                "land",
                "stud",
            ],
        }

        # 🧠 Apply mismatch detection
        mask = pd.Series(False, index=df.index)
        for occ, forbidden_list in mismatch_rules.items():
            occ_mask = df["CUSTOMER_OCCUPATION"].str.contains(occ, na=False)
            forb_mask = df["PURPOSE_OF_ACCOUNT"].str.contains(
                "|".join(forbidden_list), na=False
            )
            mask |= occ_mask & forb_mask

        df_filtered = df.loc[mask].copy()

        # 🔧 Clean numeric fields
        def clean_number(x):
            try:
                return str(int(float(x)))
            except:
                return str(x).strip()

        df_filtered["ACCOUNT_NUMBER"] = df_filtered["ACCOUNT_NUMBER"].apply(
            clean_number
        )
        df_filtered["CUSTOMER_NUMBER"] = df_filtered["CUSTOMER_NUMBER"].apply(
            clean_number
        )

        df_filtered = df_filtered[
            df_filtered["ACCOUNT_NUMBER"].notna()
            & (df_filtered["ACCOUNT_NUMBER"] != "")
        ]

        # 📤 Prepare output with final columns
        output = (
            df_filtered[
                [
                    "CUSTOMER_NUMBER",
                    "ACCOUNT_NUMBER",
                    "TITLEOFACCOUNT",
                    "PRODUCTDESC",
                    "SECTOR_DESCRIPTION",
                    "CUSTOMER_OCCUPATION",
                    "PURPOSE_OF_ACCOUNT",
                ]
            ]
            .drop_duplicates()
            .reset_index(drop=True)
        )

        # 🧯 Fallback for empty output
        if output.empty:
            print("✅ No occupation-purpose mismatches found. Returning NA row.")
            output = pd.DataFrame([{col: pd.NA for col in output.columns}])

        # 📁 Controlled export
        if mode == "full" and not output.isna().all(axis=1).iloc[0]:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_path = f"occupation_purpose_mismatch_{timestamp}.xlsx"
            output.to_excel(file_path, index=False)
            print(f"📁 Saved output to: {file_path}")
        elif mode == "full":
            print("✅ No matches found. Excel file not created.")

        return output

    except Exception as e:
        print(f"❌ Exception occurred: {e}")
        fallback_columns = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
            "CUSTOMER_OCCUPATION",
            "PURPOSE_OF_ACCOUNT",
        ]
        return pd.DataFrame([{col: pd.NA for col in fallback_columns}])
