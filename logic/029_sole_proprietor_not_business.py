import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


def get_empty_output(columns: list) -> pd.DataFrame:
    """
    Returns a single-row DataFrame filled with pd.NA for the given columns.
    Used when no contradictions are found or an error occurs.
    """
    return pd.DataFrame([{col: pd.NA for col in columns}])


@register_logic(
    name="029_sole_proprietor_not_business",
    description="Flags sole proprietor accounts (CustSectorCode = 1100) where occupation is not 'Business'.",
    category="Compliance & Screening",
)
def logic_029_sole_proprietor_not_business(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    try:
        # 🔍 Locate merged file
        merged_key = next(
            (k for k in dataframes if "merged_file.xlsx" in k.lower()), None
        )
        if not merged_key:
            raise ValueError("Merged file not found.")

        df = dataframes[merged_key].copy()
        df.columns = (
            df.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )

        # ✅ Required columns
        required_cols = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "CUSTSECTORCODE",
            "CUSTOMER_OCCUPATION",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        missing = [col for col in required_cols if col not in df.columns]
        if missing:
            raise ValueError(f"Missing required columns: {missing}")

        # 🔧 Normalize fields
        df["CUSTSECTORCODE_CLEAN"] = pd.to_numeric(
            df["CUSTSECTORCODE"], errors="coerce"
        )
        df["CUSTOMER_OCCUPATION_CLEAN"] = (
            df["CUSTOMER_OCCUPATION"].astype(str).str.lower().str.strip()
        )

        # 🧠 Apply contradiction logic
        df_filtered = df[
            (df["CUSTSECTORCODE_CLEAN"] == 1100)  # numeric comparison
            & (df["CUSTOMER_OCCUPATION_CLEAN"] != "business")
        ].copy()

        # 📤 Prepare output
        output_columns_raw = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "CUSTSECTORCODE",
            "CUSTOMER_OCCUPATION",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        output_columns_final = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "CustSectorCode",
            "Customer_Occupation",
            "ProductDesc",
            "Sector_Description",
        ]

        output = df_filtered[output_columns_raw].copy()
        output.columns = output_columns_final
        output = output.reset_index(drop=True)

        # 🧯 Handle empty output
        if output.empty:
            print("✅ No sole proprietor mismatches found.")
            return get_empty_output(output_columns_final)

        output = output.astype(object).where(pd.notna(output), None)

        # 📁 Optional export
        if mode == "full" and not output.isna().all(axis=1).iloc[0]:
            file_path = (
                f"sole_proprietor_not_business_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            )
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
            "CustSectorCode",
            "Customer_Occupation",
            "ProductDesc",
            "Sector_Description",
        ]
        return get_empty_output(fallback_columns)
