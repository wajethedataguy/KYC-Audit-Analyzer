import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


def get_empty_output(columns: list) -> pd.DataFrame:
    return pd.DataFrame([{col: pd.NA for col in columns}])


@register_logic(
    name="015_account_turnover_mismatch",
    description=(
        "Flags individual customers (CUSTSECTORCODE ≤ 1100) with narrative 'Above 10M' "
        "and corporate customers (CUSTSECTORCODE > 1100) with narrative 'Above 50M' "
        "where numeric turnover (ACCOUNT_TURNOVERGT1050M) is missing, zero, or below "
        "the respective threshold."
    ),
    category="Compliance & Screening",
)
def logic_015_account_turnover_mismatch_corrected(
    dataframes: dict, mode: str = "full"
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
            "KYC_ANN_TO_CORPORATE",
            "ACCOUNT_TURNOVERGT1050M",
            "CUSTSECTORCODE",
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
            except Exception:
                return str(val).strip()

        df["CUSTOMER_NUMBER"] = df["CUSTOMER_NUMBER"].fillna("").apply(clean_id)
        df["ACCOUNT_NUMBER"] = df["ACCOUNT_NUMBER"].fillna("").apply(clean_id)

        # 🔧 Normalize sector code
        df["CUSTSECTORCODE"] = pd.to_numeric(df["CUSTSECTORCODE"], errors="coerce")

        # 🔍 Normalize narrative labels
        df["TURNOVER_10M_NARRATIVE"] = (
            df["ACCOUNT_TURNOVER"].astype(str).str.strip().str.lower()
        )
        df["TURNOVER_50M_NARRATIVE"] = (
            df["KYC_ANN_TO_CORPORATE"].astype(str).str.strip().str.lower()
        )

        # 🔍 Numeric fields
        df["NUMERIC_TURNOVER_10M"] = pd.to_numeric(
            df["ACCOUNT_TURNOVERGT1050M"], errors="coerce"
        )
        df["NUMERIC_TURNOVER_50M"] = pd.to_numeric(
            df["ACCOUNT_TURNOVERGT1050M"], errors="coerce"
        )

        # 🧠 Mismatch 10M logic (Individuals)
        cond_10m = (
            (df["CUSTSECTORCODE"] <= 1100)
            & (df["TURNOVER_10M_NARRATIVE"] == "above 10m")
            & (df["NUMERIC_TURNOVER_10M"].isna() | (df["NUMERIC_TURNOVER_10M"] == 0))
        )
        df_10m = df[cond_10m].copy()
        if not df_10m.empty:
            df_10m["Turnover_Value"] = df_10m["NUMERIC_TURNOVER_10M"]

        # 🧠 Mismatch 50M logic (Corporate)
        cond_50m = (
            (df["CUSTSECTORCODE"] > 1100)
            & (df["TURNOVER_50M_NARRATIVE"] == "above 50m")
            & (df["NUMERIC_TURNOVER_50M"].isna() | (df["NUMERIC_TURNOVER_50M"] == 0))
        )
        df_50m = df[cond_50m].copy()
        if not df_50m.empty:
            df_50m["Turnover_Value"] = df_50m["NUMERIC_TURNOVER_50M"]

        # 📌 Combine results
        frames = []
        if not df_10m.empty:
            frames.append(df_10m)
        if not df_50m.empty:
            frames.append(df_50m)

        if not frames:
            print("✅ No mismatches found.")
            output_columns = [
                "Customer_Number",
                "Account_Number",
                "TitleOfAccount",
                "Purpose_of_Account",
                "Account_Turnover",
                "KYC_Ann_TO_Corporate",
                "Turnover_Value",
                "ProductDesc",
                "Sector_Description",
            ]
            return get_empty_output(output_columns)

        df_filtered = pd.concat(frames, ignore_index=True)

        df_filtered["SourceLogicName"] = "015_account_turnover_mismatch_corrected"
        df_filtered["Logic_Version"] = "015_v3.7"

        # 📤 Prepare output
        output = df_filtered[
            [
                "CUSTOMER_NUMBER",
                "ACCOUNT_NUMBER",
                "TITLEOFACCOUNT",
                "PURPOSE_OF_ACCOUNT",
                "ACCOUNT_TURNOVER",
                "KYC_ANN_TO_CORPORATE",
                "Turnover_Value",
                "PRODUCTDESC",
                "SECTOR_DESCRIPTION",
            ]
        ].copy()

        output.columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "Purpose_of_Account",
            "Account_Turnover",
            "KYC_Ann_TO_Corporate",
            "Turnover_Value",
            "ProductDesc",
            "Sector_Description",
        ]
        output = output.reset_index(drop=True)

        # 📁 Export if full mode
        if mode == "full" and not output.empty:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_path = f"account_turnover_mismatch_corrected_{timestamp}.xlsx"
            output.to_excel(file_path, index=False)
            print(f"📁 Saved output to: {file_path}")
        elif mode == "full":
            print("✅ No mismatches found. Excel file not created.")

        return output

    except Exception as e:
        print(f"❌ Exception occurred: {e}")
        fallback_columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "Purpose_of_Account",
            "Account_Turnover",
            "KYC_Ann_TO_Corporate",
            "Turnover_Value",
            "ProductDesc",
            "Sector_Description",
        ]
        return get_empty_output(fallback_columns)
