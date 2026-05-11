import pandas as pd
import re
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="035_cash_mode_in_fcy_without_sbp_approval",
    description="Flags FCY accounts (CustSectorCode >= 1100, CCY != PKR) where deposit or withdrawal modes mention 'cash'.",
    category="Compliance & Screening",
)
def logic_035_cash_mode_in_fcy_without_sbp_approval(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    try:
        # Step 1: Locate merged file
        merged_key = next(
            (k for k in dataframes if "merged_file.xlsx" in k.lower()), None
        )
        if not merged_key:
            raise ValueError("Merged file not found.")

        df = dataframes[merged_key].copy()

        # Step 2: Normalize columns
        df.columns = (
            df.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )

        # Step 3: Ensure required columns exist
        required_cols = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "DOMINANT_MODE_OF_DEPOSIT",
            "DOMINANT_MODE_OF_WITHDRAWAL",
            "CCY",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
            "CUSTSECTORCODE",
        ]
        if not all(col in df.columns for col in required_cols):
            raise ValueError("Missing required columns.")

        # Step 4: Clean fields
        df["DEPOSIT_CLEAN"] = (
            df["DOMINANT_MODE_OF_DEPOSIT"].astype(str).str.lower().str.strip()
        )
        df["WITHDRAWAL_CLEAN"] = (
            df["DOMINANT_MODE_OF_WITHDRAWAL"].astype(str).str.lower().str.strip()
        )
        df["CCY_CLEAN"] = df["CCY"].astype(str).str.upper().str.strip()
        df["SECTORCODE_NUM"] = pd.to_numeric(df["CUSTSECTORCODE"], errors="coerce")

        # Step 5: Filter conditions
        df_filtered = df[
            (df["SECTORCODE_NUM"] >= 1100)  # handles both 1100 and 1100.0
            & (df["CCY_CLEAN"] != "PKR")
            & (
                df["DEPOSIT_CLEAN"].str.contains("cash", na=False)
                | df["WITHDRAWAL_CLEAN"].str.contains("cash", na=False)
            )
        ].copy()

        # Step 6: Prepare output with requested columns only
        output = df_filtered[
            [
                "CUSTOMER_NUMBER",
                "ACCOUNT_NUMBER",
                "TITLEOFACCOUNT",
                "CCY",
                "PRODUCTDESC",
                "SECTOR_DESCRIPTION",
            ]
        ].copy()

        output.columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "CCY",
            "ProductDesc",
            "Sector_Description",
        ]
        output = output.reset_index(drop=True)

        # Step 7: Empty row if no records found
        if output.empty:
            output = pd.DataFrame(
                [
                    {
                        "Customer_Number": pd.NA,
                        "Account_Number": pd.NA,
                        "TitleOfAccount": pd.NA,
                        "CCY": pd.NA,
                        "ProductDesc": pd.NA,
                        "Sector_Description": pd.NA,
                    }
                ]
            )

        # Step 8: Export if mode=full
        if mode == "full":
            file_path = f"cash_mode_in_fcy_without_sbp_approval_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            output.to_excel(file_path, index=False)
            print(f"📁 Saved output to: {file_path}")

        return output

    except Exception as e:
        print(f"❌ Exception occurred: {e}")
        return pd.DataFrame(
            [
                {
                    "Customer_Number": pd.NA,
                    "Account_Number": pd.NA,
                    "TitleOfAccount": pd.NA,
                    "CCY": pd.NA,
                    "ProductDesc": pd.NA,
                    "Sector_Description": pd.NA,
                }
            ]
        )
