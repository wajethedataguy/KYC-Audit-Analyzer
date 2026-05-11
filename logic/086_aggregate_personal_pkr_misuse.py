import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="086_aggregate_personal_pkr_misuse",
    description="Flags multiple personal PKR accounts used for business purposes where aggregate annual credit turnover exceeds Rs. 450M.",
    category="CDD & EDD Review",
)
def logic_086_aggregate_personal_pkr_misuse(
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
            "CUSTSECTORCODE",
            "CUSTOMER_OCCUPATION",
            "CCY",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        missing = [col for col in required_cols if col not in df_kyc.columns]
        if missing:
            raise ValueError(f"Missing required columns: {missing}")

        # 🔧 Normalize account numbers and CNICs
        df_kyc["ACCOUNT_NUMBER_RAW"] = df_kyc["ACCOUNT_NUMBER"]
        df_kyc["ACCOUNT_NUMBER"] = (
            pd.to_numeric(df_kyc["ACCOUNT_NUMBER_RAW"], errors="coerce")
            .astype("Int64")
            .astype(str)
        )
        df_kyc["CNIC_NUMBER_CLEAN"] = (
            df_kyc["CNIC_NUMBER"].astype(str).str.replace(r"[^\d]", "", regex=True)
        )

        # 🔍 Load CRP file (normalize before detection)
        crp_key = next(
            (
                k
                for k, df in dataframes.items()
                if isinstance(df, (pd.DataFrame, tuple))
                and "ACTUAL_CREDIT_T_O_PER_ANNUM"
                in [
                    c.strip().upper().replace(" ", "_").replace("-", "_")
                    for c in (df[0] if isinstance(df, tuple) else df).columns
                ]
            ),
            None,
        )
        if not crp_key:
            raise ValueError("CRP file with credit turnover not found.")
        df_crp = dataframes[crp_key]
        if isinstance(df_crp, tuple):
            df_crp = next(iter(df_crp[0].values())) if df_crp[0] else pd.DataFrame()

        # 🔧 Normalize CRP columns
        df_crp.columns = (
            df_crp.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )
        df_crp["ACCOUNT_NO"] = (
            pd.to_numeric(df_crp["ACCOUNT_NO"], errors="coerce")
            .astype("Int64")
            .astype(str)
        )
        df_crp["ID_NO_OF_CUSTOMER"] = (
            df_crp["ID_NO_OF_CUSTOMER"]
            .astype(str)
            .str.replace(r"[^\d]", "", regex=True)
        )
        df_crp["ACTUAL_CREDIT_T_O_PER_ANNUM"] = pd.to_numeric(
            df_crp["ACTUAL_CREDIT_T_O_PER_ANNUM"]
            .astype(str)
            .str.replace(",", "")
            .str.replace("٫", "."),
            errors="coerce",
        )

        # 🔗 VLOOKUP: Credit_Lookup per account
        df_kyc = pd.merge(
            df_kyc,
            df_crp[["ACCOUNT_NO", "ACTUAL_CREDIT_T_O_PER_ANNUM"]],
            left_on="ACCOUNT_NUMBER",
            right_on="ACCOUNT_NO",
            how="left",
        ).rename(columns={"ACTUAL_CREDIT_T_O_PER_ANNUM": "CREDIT_LOOKUP"})

        # 🔗 SUMIF: Sum_Credit per CNIC
        sum_credit = (
            df_crp.groupby("ID_NO_OF_CUSTOMER", as_index=False)
            .agg({"ACTUAL_CREDIT_T_O_PER_ANNUM": "sum"})
            .rename(columns={"ACTUAL_CREDIT_T_O_PER_ANNUM": "SUM_CREDIT"})
        )
        df_kyc = pd.merge(
            df_kyc,
            sum_credit,
            left_on="CNIC_NUMBER_CLEAN",
            right_on="ID_NO_OF_CUSTOMER",
            how="left",
        )

        # 🔍 Normalize fields
        df_kyc["CUSTSECTORCODE_CLEAN"] = pd.to_numeric(
            df_kyc["CUSTSECTORCODE"], errors="coerce"
        )
        df_kyc["CCY_CLEAN"] = df_kyc["CCY"].astype(str).str.strip().str.upper()
        df_kyc["CUSTOMER_OCCUPATION_CLEAN"] = (
            df_kyc["CUSTOMER_OCCUPATION"].astype(str).str.strip().str.lower()
        )
        df_kyc["CREDIT_LOOKUP"] = pd.to_numeric(
            df_kyc["CREDIT_LOOKUP"], errors="coerce"
        )
        df_kyc["SUM_CREDIT"] = pd.to_numeric(df_kyc["SUM_CREDIT"], errors="coerce")

        # 🧠 Apply contradiction logic
        contradiction_mask = (
            (df_kyc["CUSTSECTORCODE_CLEAN"] == 1000)
            & (df_kyc["CCY_CLEAN"] == "PKR")
            & (df_kyc["CUSTOMER_OCCUPATION_CLEAN"] == "business")
            & (df_kyc["CREDIT_LOOKUP"] <= 450_000_000)
            & (df_kyc["SUM_CREDIT"] > 450_000_000)
        )

        # 📤 Final output (select raw cols, then rename)
        output_cols = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "CUSTOMERFULLNAME",
            "CNIC_NUMBER",
            "CUSTSECTORCODE",
            "CUSTOMER_OCCUPATION",
            "CCY",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
            "CREDIT_LOOKUP",
            "SUM_CREDIT",
        ]
        output = (
            df_kyc[contradiction_mask][output_cols]
            .drop_duplicates()
            .reset_index(drop=True)
        )

        output.columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "CustomerFullName",
            "CNIC_Number",
            "CustSectorCode",
            "Customer_Occupation",
            "CCY",
            "ProductDesc",
            "Sector_Description",
            "Credit_Lookup",
            "Sum_Credit",
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
                        "CustSectorCode": pd.NA,
                        "Customer_Occupation": pd.NA,
                        "CCY": pd.NA,
                        "ProductDesc": pd.NA,
                        "Sector_Description": pd.NA,
                        "Credit_Lookup": pd.NA,
                        "Sum_Credit": pd.NA,
                    }
                ]
            )

        # 📁 Optional export
        if mode == "full":
            file_path = (
                f"aggregate_personal_pkr_misuse_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            )
            output.to_excel(file_path, index=False)
        elif mode == "ui":
            output = output[
                [
                    "Customer_Number",
                    "Account_Number",
                    "TitleOfAccount",
                    "CustomerFullName",
                    "CustSectorCode",
                    "Customer_Occupation",
                    "CCY",
                    "ProductDesc",
                    "Sector_Description",
                ]
            ]

        return output

    except Exception as e:
        print(f"❌ Error in logic_086_aggregate_personal_pkr_misuse: {e}")
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
                    "CCY": pd.NA,
                    "ProductDesc": pd.NA,
                    "Sector_Description": pd.NA,
                    "Credit_Lookup": pd.NA,
                    "Sum_Credit": pd.NA,
                }
            ]
        )
