import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


def normalize_text(value: str) -> str:
    """
    Normalize text values for consistent comparison.
    Removes spaces, dots, slashes, and converts to lowercase.
    """
    return str(value).strip().lower().replace(".", "").replace("/", "").replace(" ", "")


@register_logic(
    name="071_pep_customer_not_marked_high_risk",
    description="Flags PEP customers whose KYC risk rating is not marked as 'High'.",
    category="Compliance & Screening",
)
def logic_071_pep_customer_not_marked_high_risk(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    try:
        # 🔍 Locate merged KYC file
        kyc_key = next((k for k in dataframes if "merged_file.xlsx" in k.lower()), None)
        if not kyc_key:
            raise ValueError("KYC profile file not found.")
        df_main = dataframes[kyc_key].copy()

        # 🔧 Normalize column names
        df_main.columns = (
            df_main.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )

        # ✅ Required columns
        required_cols = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "CUSTOMERFULLNAME",
            "CNIC_NUMBER",
            "APPROVALOBTAINEDFORPEP",
            "KYCRISK",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        missing = [col for col in required_cols if col not in df_main.columns]
        if missing:
            raise ValueError(f"Merged file missing required columns: {missing}")

        df = df_main[required_cols].copy()

        # 🔍 Normalize fields
        df["PEP_APPROVAL_STR"] = (
            df["APPROVALOBTAINEDFORPEP"].fillna("").apply(normalize_text)
        )
        df["KYCRISK_STR"] = df["KYCRISK"].fillna("").apply(normalize_text)

        # 🧠 Define incorrect risk types
        incorrect_risk = {"low", "medium"}

        # 🧠 Apply contradiction logic
        df_filtered = df[
            (df["PEP_APPROVAL_STR"] == "yes") & (df["KYCRISK_STR"].isin(incorrect_risk))
        ].copy()

        # 📤 Prepare output (select raw cols, then rename)
        output = df_filtered[
            [
                "CUSTOMER_NUMBER",
                "ACCOUNT_NUMBER",
                "TITLEOFACCOUNT",
                "CUSTOMERFULLNAME",
                "CNIC_NUMBER",
                "APPROVALOBTAINEDFORPEP",
                "KYCRISK",
                "PRODUCTDESC",
                "SECTOR_DESCRIPTION",
            ]
        ].copy()

        output.columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "CustomerFullName",
            "CNIC_Number",
            "ApprovalObtainedForPEP",
            "KYCRisk",
            "ProductDesc",
            "Sector_Description",
        ]
        output = output.reset_index(drop=True)

        # 🧯 Handle empty output
        if output.empty:
            output = pd.DataFrame(
                [
                    {
                        "Customer_Number": pd.NA,
                        "Account_Number": pd.NA,
                        "TitleOfAccount": pd.NA,
                        "CustomerFullName": pd.NA,
                        "CNIC_Number": pd.NA,
                        "ApprovalObtainedForPEP": pd.NA,
                        "KYCRisk": pd.NA,
                        "ProductDesc": pd.NA,
                        "Sector_Description": pd.NA,
                    }
                ]
            )

        # 📁 Optional export
        if mode == "full":
            file_path = (
                f"pep_customer_not_marked_high_risk_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            )
            output.to_excel(file_path, index=False)
            print(f"✅ Logic 071 output saved at {file_path}")

        return output

    except Exception as e:
        print(f"❌ Exception occurred: {e}")
        return pd.DataFrame(
            [
                {
                    "Customer_Number": pd.NA,
                    "Account_Number": pd.NA,
                    "TitleOfAccount": pd.NA,
                    "CustomerFullName": pd.NA,
                    "CNIC_Number": pd.NA,
                    "ApprovalObtainedForPEP": pd.NA,
                    "KYCRisk": pd.NA,
                    "ProductDesc": pd.NA,
                    "Sector_Description": pd.NA,
                }
            ]
        )
