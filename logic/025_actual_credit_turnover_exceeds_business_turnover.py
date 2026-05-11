import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="025_actual_credit_turnover_exceeds_business_turnover",
    description="Flags top 15 customers where actual credit turnover exceeds declared business turnover.",
    category="Compliance & Screening",
)
def logic_025_actual_credit_turnover_exceeds_business_turnover(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    try:
        # Step 1: Locate merged file
        merged_key = next(
            (k for k in dataframes if "merged_file.xlsx" in k.lower()), None
        )
        if merged_key is None:
            raise ValueError("Merged file not found in input dataframes.")
        df_main = dataframes[merged_key].copy()

        # Step 2: Normalize columns
        df_main.columns = (
            df_main.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )

        # Step 3: Ensure required columns exist
        required_main = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "BUSINESSTURNOVER",
            "TOTAL_CREDIT_TOTAL",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
            "CUSTSECTORCODE",
        ]
        missing_main = [col for col in required_main if col not in df_main.columns]
        if missing_main:
            raise ValueError(f"Missing required columns in merged file: {missing_main}")

        # Step 4: Clean numeric fields
        df_main["CREDIT_TURNOVER_NUM"] = pd.to_numeric(
            df_main["TOTAL_CREDIT_TOTAL"]
            .astype(str)
            .str.replace(r"[^\d.]", "", regex=True),
            errors="coerce",
        )
        df_main["BUSINESS_TURNOVER_NUM"] = pd.to_numeric(
            df_main["BUSINESSTURNOVER"]
            .astype(str)
            .str.replace(r"[^\d.]", "", regex=True),
            errors="coerce",
        )

        # Step 5: Filter contradictions (no occupation condition)
        df_filtered = df_main[
            df_main["BUSINESS_TURNOVER_NUM"].notna()
            & (df_main["BUSINESS_TURNOVER_NUM"] > 0)
            & (df_main["CREDIT_TURNOVER_NUM"] > df_main["BUSINESS_TURNOVER_NUM"])
        ].copy()

        # Step 6: Calculate difference
        df_filtered["DIFFERENCE"] = (
            df_filtered["CREDIT_TURNOVER_NUM"] - df_filtered["BUSINESS_TURNOVER_NUM"]
        )

        # Step 7: Sort and select top 15
        df_top15 = df_filtered.sort_values(by="DIFFERENCE", ascending=False).head(15)

        # Step 8: Prepare output
        output = df_top15[
            [
                "CUSTOMER_NUMBER",
                "ACCOUNT_NUMBER",
                "TITLEOFACCOUNT",
                "PRODUCTDESC",
                "SECTOR_DESCRIPTION",
                "TOTAL_CREDIT_TOTAL",
                "BUSINESSTURNOVER",
                "DIFFERENCE",
            ]
        ].copy()

        output.columns = [
            "Customer ID",
            "Account Number",
            "Account Title",
            "ProductDesc",
            "SECTOR_DESCRIPTION",
            "Account Annual Actual Credit Turnover",
            "Business Turnover",
            "Difference",
        ]

        output = output.reset_index(drop=True)

        # Step 9: Return preview-friendly empty row if no records found
        if output.empty:
            print(
                "✅ No credit turnover vs business turnover contradictions found. Returning preview-friendly empty row."
            )
            output = pd.DataFrame(
                [
                    {
                        "Customer ID": "",
                        "Account Number": "",
                        "Account Title": "",
                        "ProductDesc": "",
                        "SECTOR_DESCRIPTION": "",
                        "Account Annual Actual Credit Turnover": "",
                        "Business Turnover": "",
                        "Difference": "",
                    }
                ]
            )

        # Step 10: Export if needed
        if mode == "full":
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_path = (
                f"actual_credit_turnover_exceeds_business_turnover_{timestamp}.xlsx"
            )
            output.to_excel(file_path, index=False)
            print(f"📁 Saved output to: {file_path}")

        return output

    except Exception as e:
        print(f"❌ Exception occurred: {e}")
        return pd.DataFrame(
            [
                {
                    "Customer ID": "",
                    "Account Number": "",
                    "Account Title": "",
                    "ProductDesc": "",
                    "SECTOR_DESCRIPTION": "",
                    "Account Annual Actual Credit Turnover": "",
                    "Business Turnover": "",
                    "Difference": "",
                }
            ]
        )
