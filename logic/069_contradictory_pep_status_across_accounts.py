import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


def normalize_text(value: str) -> str:
    return str(value).strip().upper().replace(".", "").replace("/", "").replace(" ", "")


@register_logic(
    name="069_contradictory_pep_status_across_accounts",
    description="Flags CNICs where ApprovalObtainedForPEP status is contradictory across multiple accounts tagged with different customer IDs.",
    category="Compliance & Screening",
)
def logic_069_contradictory_pep_status_across_accounts(
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
            "CUSTSECTORCODE",
            "CUSTOMER_OCCUPATION",
            "APPROVALOBTAINEDFORPEP",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        missing = [col for col in required_cols if col not in df_main.columns]
        if missing:
            raise ValueError(f"Merged file missing required columns: {missing}")

        df = df_main[required_cols].copy()

        # 🔍 Normalize CNIC and ApprovalObtainedForPEP fields
        df["CNIC_NUMBER_STR"] = df["CNIC_NUMBER"].fillna("").apply(normalize_text)
        df["APPROVALOBTAINEDFORPEP_STR"] = (
            df["APPROVALOBTAINEDFORPEP"].fillna("").astype(str).str.lower().str.strip()
        )

        # 🧠 Group by CNIC and count unique ApprovalObtainedForPEP statuses
        grouped = (
            df.groupby("CNIC_NUMBER_STR")
            .agg(
                ApprovalObtainedForPEP_Status=(
                    "APPROVALOBTAINEDFORPEP_STR",
                    lambda x: "|".join(sorted(set(x.dropna()))),
                ),
                UNIQUE_PEP_STATUSES=(
                    "APPROVALOBTAINEDFORPEP_STR",
                    lambda x: x.dropna().nunique(),
                ),
            )
            .reset_index()
        )

        # 🔍 Flag CNICs with contradictory ApprovalObtainedForPEP status
        flagged_cnic = grouped[grouped["UNIQUE_PEP_STATUSES"] > 1][
            "CNIC_NUMBER_STR"
        ].tolist()

        # 🔍 Filter original rows for flagged CNICs
        df_flagged = df[df["CNIC_NUMBER_STR"].isin(flagged_cnic)].copy()

        # 🔗 Merge contradiction info back
        df_flagged = df_flagged.merge(
            grouped[["CNIC_NUMBER_STR", "ApprovalObtainedForPEP_Status"]],
            on="CNIC_NUMBER_STR",
            how="left",
        )

        # 📤 Prepare output (select raw cols, then rename)
        output = df_flagged[
            [
                "CUSTOMER_NUMBER",
                "ACCOUNT_NUMBER",
                "TITLEOFACCOUNT",
                "CUSTOMERFULLNAME",
                "CNIC_NUMBER",
                "CUSTSECTORCODE",
                "CUSTOMER_OCCUPATION",
                "APPROVALOBTAINEDFORPEP",
                "PRODUCTDESC",
                "SECTOR_DESCRIPTION",
                "ApprovalObtainedForPEP_Status",  # merged contradiction info
            ]
        ].copy()

        output.columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "CustomerFullName",
            "CNIC_Number",
            "CustSectorCode",
            "Customer_Occupation",
            "ApprovalObtainedForPEP_Original",
            "ProductDesc",
            "Sector_Description",
            "ApprovalObtainedForPEP_Contradictions",
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
                        "CustSectorCode": pd.NA,
                        "Customer_Occupation": pd.NA,
                        "ApprovalObtainedForPEP_Original": pd.NA,
                        "ProductDesc": pd.NA,
                        "Sector_Description": pd.NA,
                        "ApprovalObtainedForPEP_Contradictions": pd.NA,
                    }
                ]
            )

        # 📁 Optional export
        if mode == "full":
            file_path = f"contradictory_pep_status_across_accounts_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
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
                    "CustomerFullName": pd.NA,
                    "CNIC_Number": pd.NA,
                    "CustSectorCode": pd.NA,
                    "Customer_Occupation": pd.NA,
                    "ApprovalObtainedForPEP_Original": pd.NA,
                    "ProductDesc": pd.NA,
                    "Sector_Description": pd.NA,
                    "ApprovalObtainedForPEP_Contradictions": pd.NA,
                }
            ]
        )
