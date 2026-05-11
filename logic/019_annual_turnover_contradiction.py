import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


def get_empty_output(columns: list) -> pd.DataFrame:
    return pd.DataFrame([{col: pd.NA for col in columns}])


@register_logic(
    name="019_annual_turnover_contradiction",
    description="Flags customers where declared annual turnover is 'Above 10M' or 'Above 50M' but numeric value is below threshold.",
    category="Compliance & Screening",
)
def logic_019_annual_turnover_contradiction(
    dataframes: dict, mode: str = "preview"
) -> pd.DataFrame:
    try:
        # 🔍 Locate merged file
        merged_key = next(
            (k for k in dataframes if "merged_file.xlsx" in k.lower()), None
        )
        if merged_key is None:
            raise ValueError("Merged file not found in input dataframes.")
        df = dataframes[merged_key].copy()

        # 🔧 Normalize column names
        df.columns = (
            df.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )

        # ✅ Required columns
        required_columns = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "PURPOSE_OF_ACCOUNT",
            "ACCOUNT_TURNOVER",
            "ACCOUNT_TURNOVERGT1050M",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        missing = [col for col in required_columns if col not in df.columns]
        if missing:
            raise ValueError(f"Missing required columns: {missing}")

        # 🔧 Clean identifiers
        def clean_id(val):
            try:
                num = float(val)
                return str(int(num)) if num.is_integer() else str(val).strip()
            except:
                return str(val).strip()

        df["CUSTOMER_NUMBER"] = df["CUSTOMER_NUMBER"].fillna("").apply(clean_id)
        df["ACCOUNT_NUMBER"] = df["ACCOUNT_NUMBER"].fillna("").apply(clean_id)

        # 🔍 Normalize narrative field
        df["NARRATIVE_ANNUAL"] = (
            df["ACCOUNT_TURNOVER"].astype(str).str.lower().str.strip()
        )

        # 🔍 Normalize numeric field (drop commas / PKR, keep decimal)
        cleaned_numeric = (
            df["ACCOUNT_TURNOVERGT1050M"]
            .astype(str)
            .str.replace(",", "", regex=False)
            .str.replace("PKR", "", case=False, regex=False)
            .str.strip()
        )
        df["NUMERIC_ANNUAL"] = pd.to_numeric(cleaned_numeric, errors="coerce")

        # 🧠 Contradiction logic
        def detect_mismatch(row):
            turnover = row["NARRATIVE_ANNUAL"]
            amount = row["NUMERIC_ANNUAL"]

            if pd.isna(amount) or amount <= 0:
                return None

            if turnover == "above 10m" and amount < 10_000_000:
                return "Declared Above 10M but numeric < 10M"
            elif turnover == "above 50m" and amount < 50_000_000:
                return "Declared Above 50M but numeric < 50M"
            return None

        df["MISMATCH_REASON"] = df.apply(detect_mismatch, axis=1)

        # 🔒 Focus only on mismatches
        df_filtered = df[
            df["MISMATCH_REASON"].notnull()
            & df["NARRATIVE_ANNUAL"].isin(["above 10m", "above 50m"])
        ].copy()

        # 📤 Prepare output
        output_columns = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "PURPOSE_OF_ACCOUNT",
            "ACCOUNT_TURNOVER",
            "ACCOUNT_TURNOVERGT1050M",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        output = df_filtered[output_columns].reset_index(drop=True)

        # 🧯 Handle empty output
        if output.empty:
            print("✅ No annual turnover contradictions found.")
            output = get_empty_output(output_columns)

        # 📁 Controlled export
        if mode == "full" and not output.isna().all(axis=1).iloc[0]:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_path = f"annual_turnover_contradiction_{timestamp}.xlsx"
            output.to_excel(file_path, index=False)
            print(f"📁 Saved output to: {file_path}")
        elif mode == "full":
            print("✅ No annual turnover contradictions found. Excel file not created.")

        return output

    except Exception as e:
        print(f"❌ Exception occurred: {e}")
        fallback_columns = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "PURPOSE_OF_ACCOUNT",
            "ACCOUNT_TURNOVER",
            "ACCOUNT_TURNOVERGT1050M",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        return get_empty_output(fallback_columns)
