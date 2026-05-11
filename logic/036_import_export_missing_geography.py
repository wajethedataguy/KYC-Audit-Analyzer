import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="036_import_export_missing_geography",
    description="Flags customers with 'Import' or 'Export' nature of business but expected geography is set to Pakistan or PK.",
    category="Compliance & Screening",
)
def logic_036_import_export_missing_geography(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    try:
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

        required_cols = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "CCY",
            "NATUREOFBUSINESS",
            "EXPINTERNATIONALGEOGRAPHY",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        if not all(col in df.columns for col in required_cols):
            raise ValueError("Missing required columns.")

        df["BUSINESS_CLEAN"] = (
            df["NATUREOFBUSINESS"].astype(str).str.lower().str.strip()
        )
        df["GEOGRAPHY_CLEAN"] = (
            df["EXPINTERNATIONALGEOGRAPHY"].astype(str).str.upper().str.strip()
        )

        keywords = ["import", "export"]
        df_filtered = df[
            df["BUSINESS_CLEAN"].apply(lambda x: any(k in x for k in keywords))
            & df["GEOGRAPHY_CLEAN"].isin(["PAKISTAN", "PK"])
        ].copy()

        output = df_filtered[
            [
                "CUSTOMER_NUMBER",
                "ACCOUNT_NUMBER",
                "TITLEOFACCOUNT",
                "CCY",
                "NATUREOFBUSINESS",
                "EXPINTERNATIONALGEOGRAPHY",
                "PRODUCTDESC",
                "SECTOR_DESCRIPTION",
            ]
        ].copy()

        output.columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "CCY",
            "NatureOfBusiness",
            "ExpInternationalGeography",
            "ProductDesc",
            "Sector_Description",
        ]
        output = output.reset_index(drop=True)

        if output.empty:
            output = pd.DataFrame(
                [
                    {
                        "Customer_Number": pd.NA,
                        "Account_Number": pd.NA,
                        "TitleOfAccount": pd.NA,
                        "CCY": pd.NA,
                        "NatureOfBusiness": pd.NA,
                        "ExpInternationalGeography": pd.NA,
                        "ProductDesc": pd.NA,
                        "Sector_Description": pd.NA,
                    }
                ]
            )

        if mode == "full":
            file_path = (
                f"import_export_missing_geography_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            )
            output.to_excel(file_path, index=False)

        return output

    except Exception:
        return pd.DataFrame(
            [
                {
                    "Customer_Number": pd.NA,
                    "Account_Number": pd.NA,
                    "TitleOfAccount": pd.NA,
                    "CCY": pd.NA,
                    "NatureOfBusiness": pd.NA,
                    "ExpInternationalGeography": pd.NA,
                    "ProductDesc": pd.NA,
                    "Sector_Description": pd.NA,
                }
            ]
        )
