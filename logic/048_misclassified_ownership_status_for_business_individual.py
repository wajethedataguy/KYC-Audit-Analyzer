import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="048_misclassified_ownership_status_for_business_individual",
    description="Flags business individuals with 'Sole Proprietor / Owner' status whose business name matches another account holder, indicating partnership or directorship.",
    category="Compliance & Screening",
)
def logic_048_misclassified_ownership_status_for_business_individual(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    try:
        # 🔍 Locate merged KYC file
        kyc_key = next((k for k in dataframes if "merged_file.xlsx" in k.lower()), None)
        if not kyc_key:
            raise ValueError("KYC profile file not found.")
        df = dataframes[kyc_key].copy()

        # 🔧 Normalize column names
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
            "CUSTOMERFULLNAME",
            "CUSTSECTORCODE",
            "CUSTOMER_OCCUPATION",
            "STATUSOFOWNERSHIP",
            "NAMEOFBUSINESS",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        missing = [col for col in required_cols if col not in df.columns]
        if missing:
            raise ValueError(f"Merged file missing required columns: {missing}")

        df = df[required_cols].copy()

        # 🔍 Normalize fields
        df["STATUSOFOWNERSHIP_CLEAN"] = (
            df["STATUSOFOWNERSHIP"].astype(str).str.strip().str.upper()
        )
        df["CUSTSECTORCODE"] = pd.to_numeric(df["CUSTSECTORCODE"], errors="coerce")

        # 🧠 Ownership keywords
        ownership_keywords = ["SOLE", "PROP", "OWNE"]

        # 🔍 Filter original accounts
        mask_sector_1000 = df["CUSTSECTORCODE"] == 1000
        mask_status_nonempty = df["STATUSOFOWNERSHIP_CLEAN"] != ""
        mask_status_matches = df["STATUSOFOWNERSHIP_CLEAN"].apply(
            lambda val: any(keyword in val for keyword in ownership_keywords)
        )
        original_accounts = df[
            mask_sector_1000 & mask_status_nonempty & mask_status_matches
        ].copy()

        # 🔍 Search pool: sector code > 1100
        search_pool = df[df["CUSTSECTORCODE"] > 1100].copy()

        # 🔧 Clean join keys
        search_pool["SEARCH_TITLE_CLEAN"] = (
            search_pool["TITLEOFACCOUNT"].astype(str).str.strip().str.upper()
        )
        original_accounts["BUSINESS_NAME_CLEAN"] = (
            original_accounts["NAMEOFBUSINESS"].astype(str).str.strip().str.upper()
        )

        # 🔗 Inner join on cleaned names
        joined = original_accounts.merge(
            search_pool[["SEARCH_TITLE_CLEAN", "ACCOUNT_NUMBER", "TITLEOFACCOUNT"]],
            left_on="BUSINESS_NAME_CLEAN",
            right_on="SEARCH_TITLE_CLEAN",
            how="inner",
            suffixes=("", "_SEARCH"),
        )

        # 📤 Prepare output (Customer_Number always first)
        output = joined[
            [
                "CUSTOMER_NUMBER",
                "ACCOUNT_NUMBER",
                "TITLEOFACCOUNT",
                "STATUSOFOWNERSHIP",
                "CUSTSECTORCODE",
                "NAMEOFBUSINESS",
                "ACCOUNT_NUMBER_SEARCH",
                "TITLEOFACCOUNT_SEARCH",
                "PRODUCTDESC",
                "SECTOR_DESCRIPTION",
            ]
        ].copy()

        output.columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "StatusOfOwnership",
            "CustSectorCode",
            "NameOfBusiness",
            "Search_Account_Number",
            "Search_TitleOfAccount",
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
                        "StatusOfOwnership": pd.NA,
                        "CustSectorCode": pd.NA,
                        "NameOfBusiness": pd.NA,
                        "Search_Account_Number": pd.NA,
                        "Search_TitleOfAccount": pd.NA,
                        "ProductDesc": pd.NA,
                        "Sector_Description": pd.NA,
                    }
                ]
            )

        # 📁 Optional export
        if mode == "full":
            file_path = f"misclassified_ownership_status_for_business_individual_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            output.to_excel(file_path, index=False)

        return output

    except Exception as e:
        print(f"❌ Exception occurred: {e}")
        return pd.DataFrame(
            [
                {
                    "Customer_Number": pd.NA,
                    "Account_Number": pd.NA,
                    "TitleOfAccount": pd.NA,
                    "StatusOfOwnership": pd.NA,
                    "CustSectorCode": pd.NA,
                    "NameOfBusiness": pd.NA,
                    "Search_Account_Number": pd.NA,
                    "Search_TitleOfAccount": pd.NA,
                    "ProductDesc": pd.NA,
                    "Sector_Description": pd.NA,
                }
            ]
        )
