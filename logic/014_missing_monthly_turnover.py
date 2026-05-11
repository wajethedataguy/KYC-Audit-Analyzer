import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="014_missing_monthly_turnover",
    description="Flags customers whose Account Expected Monthly Credit Turnover is missing or recorded as zero in both ACCOUNT_MONTHLY_TURNOVER and MON_TOVER_CRG.",
    category="Compliance & Screening",
)
def logic_014_missing_monthly_turnover(dataframes: dict, mode="full") -> pd.DataFrame:
    try:
        merged_key = next(
            (k for k in dataframes if "merged_file.xlsx" in k.lower()), None
        )
        if merged_key is None:
            raise ValueError("Merged file not found in input dataframes.")
        df = dataframes[merged_key].copy()

        df.columns = (
            df.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )

        required_columns = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "PURPOSE_OF_ACCOUNT",
            "STATUS_OF_ACCOUNT",
            "ACCOUNT_MONTHLY_TURNOVER",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        missing = [col for col in required_columns if col not in df.columns]
        if missing:
            raise ValueError(f"Missing required columns: {missing}")

        def clean_id(val):
            try:
                num = float(val)
                return str(int(num)) if num.is_integer() else str(val).strip()
            except:
                return str(val).strip()

        def normalize_turnover(val):
            val = str(val).strip().upper()
            if val in {"", "0", "0.0", "0.00", "0.000", "NULL", "NONE"}:
                return "0.00"
            try:
                val = val.replace("PKR", "").replace(",", "").strip()
                num = float(val)
                return f"{num:.2f}"
            except:
                return val

        df["CUSTOMER_NUMBER"] = df["CUSTOMER_NUMBER"].fillna("").apply(clean_id)
        df["ACCOUNT_NUMBER"] = df["ACCOUNT_NUMBER"].fillna("").apply(clean_id)
        df["STATUS_OF_ACCOUNT"] = (
            df["STATUS_OF_ACCOUNT"].fillna("").astype(str).str.strip()
        )
        df["TITLEOFACCOUNT"] = df["TITLEOFACCOUNT"].fillna("").astype(str).str.strip()
        df["PURPOSE_OF_ACCOUNT"] = (
            df["PURPOSE_OF_ACCOUNT"].fillna("").astype(str).str.strip()
        )

        df["TURNOVER_1"] = (
            df["ACCOUNT_MONTHLY_TURNOVER"].fillna("").apply(normalize_turnover)
        )
       
        df_filtered = df[
            (df["TITLEOFACCOUNT"] != "")
            & (df["PURPOSE_OF_ACCOUNT"] != "")
            & (df["CUSTOMER_NUMBER"].str.len() >= 4)
        ]

        filtered_df = df_filtered[
            (df_filtered["TURNOVER_1"] == "0.00")
        ].copy()

        filtered_df["Monthly_Turnover"] = "0.00"

        output = filtered_df[
            [
                "CUSTOMER_NUMBER",
                "ACCOUNT_NUMBER",
                "TITLEOFACCOUNT",
                "Monthly_Turnover",
                "PRODUCTDESC",
                "SECTOR_DESCRIPTION",
            ]
        ].copy()

        output.columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "Monthly_Turnover",
            "ProductDesc",
            "Sector_Description",
        ]
        output = output.reset_index(drop=True)

        # --- Handle empty output consistently ---
        if output.empty:
            print(
                "✅ No missing monthly turnover records found. Returning preview-friendly empty row."
            )
            output = pd.DataFrame(
                [
                    {
                        "Customer_Number": "",
                        "Account_Number": "",
                        "TitleOfAccount": "",
                        "Monthly_Turnover": "",
                        "ProductDesc": "",
                        "Sector_Description": "",
                    }
                ]
            )

        # --- Export if mode is full ---
        if mode == "full":
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_path = f"missing_monthly_turnover_{timestamp}.xlsx"
            output.to_excel(file_path, index=False)
            print(f"📁 Saved output to: {file_path}")

        return output

    except Exception as e:
        print(f"❌ Exception occurred: {e}")
        return pd.DataFrame(
            [
                {
                    "Customer_Number": "",
                    "Account_Number": "",
                    "TitleOfAccount": "",
                    "Monthly_Turnover": "",
                    "ProductDesc": "",
                    "Sector_Description": "",
                }
            ]
        )
