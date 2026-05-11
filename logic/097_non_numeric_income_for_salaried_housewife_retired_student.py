import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="097_non_numeric_income_for_salaried_housewife_retired_student",
    description="Flags KYC records where non-numeric salary/income was fed for salaried, housewife, retired, or student customers.",
    category="CDD & EDD Review",
)
def logic_097_non_numeric_income_for_salaried_housewife_retired_student(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    try:
        # 🔍 Load merged KYC file
        kyc_key = next((k for k in dataframes if "merged_file.xlsx" in k.lower()), None)
        df_kyc = dataframes.get(kyc_key)
        if isinstance(df_kyc, tuple):
            df_kyc = next(iter(df_kyc[0].values())) if df_kyc[0] else pd.DataFrame()
        if df_kyc is None or df_kyc.empty:
            raise ValueError("merged_file.xlsx not found or empty.")

        # 🔧 Normalize column names
        df_kyc.columns = (
            df_kyc.columns.str.strip()
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
            "CUSTOMER_OCCUPATION",
            "SALARY_OTHER_INCOME",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        missing = [col for col in required_cols if col not in df_kyc.columns]
        if missing:
            raise ValueError(f"Missing required columns: {missing}")

        # 🔧 Normalize fields
        df_kyc["CUSTOMER_OCCUPATION_CLEAN"] = (
            df_kyc["CUSTOMER_OCCUPATION"].astype(str).str.strip().str.lower()
        )

        # Try to convert salary/income to numeric
        df_kyc["SALARY_OTHER_INCOME_CLEAN"] = pd.to_numeric(
            df_kyc["SALARY_OTHER_INCOME"], errors="coerce"
        )

        # 🧠 Define target occupations
        target_occupations = {"salaried", "housewife", "retired", "student"}

        orig = (
            df_kyc["SALARY_OTHER_INCOME"]
            .astype(str)
            .str.strip()
            .str.lower()
        )
        # 🔍 Contradiction: occupation in target set AND salary/income is non-numeric
        contradiction_mask = (
            df_kyc["CUSTOMER_OCCUPATION_CLEAN"].isin(target_occupations)
            & df_kyc["SALARY_OTHER_INCOME_CLEAN"].isna()
            & df_kyc["SALARY_OTHER_INCOME"].notna()
            & ~orig.isin({"", "nan", "none", "na", "n/a", "null"})
        )

        # 📤 Final output
        output = (
            df_kyc[contradiction_mask][required_cols]
            .drop_duplicates()
            .reset_index(drop=True)
        )

        output.columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "CustomerFullName",
            "CNIC_Number",
            "Customer_Occupation",
            "Salary_Other_Income",
            "ProductDesc",
            "Sector_Description",
        ]

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
                        "Customer_Occupation": pd.NA,
                        "Salary_Other_Income": pd.NA,
                        "ProductDesc": pd.NA,
                        "Sector_Description": pd.NA,
                    }
                ]
            )

        # 📁 Optional export
        if mode == "full":
            file_path = f"non_numeric_income_flag_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            output.to_excel(file_path, index=False)
        elif mode == "ui":
            output = output[
                [
                    "Customer_Number",
                    "Account_Number",
                    "TitleOfAccount",
                    "CustomerFullName",
                    "CNIC_Number",
                    "Customer_Occupation",
                    "Salary_Other_Income",
                    "ProductDesc",
                    "Sector_Description",
                ]
            ]

        return output

    except Exception as e:
        print(
            f"❌ Error in logic_097_non_numeric_income_for_salaried_housewife_retired_student: {e}"
        )
        return pd.DataFrame(
            [
                {
                    "Customer_Number": pd.NA,
                    "Account_Number": pd.NA,
                    "TitleOfAccount": pd.NA,
                    "CustomerFullName": pd.NA,
                    "CNIC_Number": pd.NA,
                    "Customer_Occupation": pd.NA,
                    "Salary_Other_Income": pd.NA,
                    "ProductDesc": pd.NA,
                    "Sector_Description": pd.NA,
                }
            ]
        )
