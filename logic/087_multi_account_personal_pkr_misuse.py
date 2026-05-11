import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="087_multi_account_business_use_violation",
    description="Flags multiple personal PKR accounts used for business purposes instead of a single designated account, violating Compliance Operational Manual.",
    category="CDD & EDD Review",
)
def logic_087_multi_account_business_use_violation(
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
            "NAMEOFBUSINESS",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        missing = [col for col in required_cols if col not in df_kyc.columns]
        if missing:
            raise ValueError(f"Missing required columns: {missing}")

        # 🔧 Normalize KYC fields
        df_kyc["ACCOUNT_NUMBER"] = (
            pd.to_numeric(df_kyc["ACCOUNT_NUMBER"], errors="coerce")
            .astype("Int64")
            .astype(str)
        )
        df_kyc["CNIC_NUMBER_CLEAN"] = (
            df_kyc["CNIC_NUMBER"].astype(str).str.replace(r"[^\d]", "", regex=True)
        )
        df_kyc["CCY_CLEAN"] = df_kyc["CCY"].astype(str).str.strip().str.upper()
        df_kyc["CUSTOMER_OCCUPATION_CLEAN"] = (
            df_kyc["CUSTOMER_OCCUPATION"].astype(str).str.strip().str.lower()
        )
        df_kyc["NAMEOFBUSINESS_CLEAN"] = (
            df_kyc["NAMEOFBUSINESS"].astype(str).str.strip().str.lower()
        )
        df_kyc["CUSTSECTORCODE_CLEAN"] = pd.to_numeric(
            df_kyc["CUSTSECTORCODE"], errors="coerce"
        )

        # 📊 Load FT Transaction Summary
        insight_frames = []
        for k, df in dataframes.items():
            if isinstance(df, tuple):
                df = next(iter(df[0].values())) if df[0] else pd.DataFrame()
            df.columns = (
                df.columns.str.strip()
                .str.upper()
                .str.replace(" ", "_")
                .str.replace("-", "_")
            )
            if all(
                col in df.columns
                for col in ["DEBIT_ACCOUNT", "CR_ACCOUNT", "TRANSACTION_AMOUNT"]
            ):
                df_debit = df[["DEBIT_ACCOUNT", "TRANSACTION_AMOUNT"]].rename(
                    columns={
                        "DEBIT_ACCOUNT": "ACCOUNT",
                        "TRANSACTION_AMOUNT": "TR_AMOUNT",
                    }
                )
                df_credit = df[["CR_ACCOUNT", "TRANSACTION_AMOUNT"]].rename(
                    columns={"CR_ACCOUNT": "ACCOUNT", "TRANSACTION_AMOUNT": "TR_AMOUNT"}
                )
                insight_frames.append(
                    pd.concat([df_debit, df_credit], ignore_index=True)
                )

        df_ft_summary = (
            pd.concat(insight_frames, ignore_index=True)
            if insight_frames
            else pd.DataFrame(columns=["ACCOUNT", "TR_AMOUNT"])
        )
        df_ft_summary["ACCOUNT"] = (
            pd.to_numeric(df_ft_summary["ACCOUNT"], errors="coerce")
            .astype("Int64")
            .astype(str)
        )
        df_ft_summary["TR_AMOUNT"] = pd.to_numeric(
            df_ft_summary["TR_AMOUNT"]
            .astype(str)
            .str.replace(r"[^\d.]", "", regex=True),
            errors="coerce",
        )
        df_ft_summary = df_ft_summary.groupby("ACCOUNT", as_index=False).agg(
            TOTAL_COUNT=("TR_AMOUNT", "count"), TOTAL_AMOUNT=("TR_AMOUNT", "sum")
        )

        # 🔗 Join KYC with transaction summary
        df_merged = pd.merge(
            df_kyc,
            df_ft_summary,
            left_on="ACCOUNT_NUMBER",
            right_on="ACCOUNT",
            how="left",
        )

        # 🧠 Apply filter: personal PKR accounts with business-like activity
        filtered = df_merged[
            (df_merged["CUSTSECTORCODE_CLEAN"] == 1000)  # personal accounts
            & (df_merged["CCY_CLEAN"] == "PKR")
            & (df_merged["TOTAL_COUNT"] > 30)  # threshold for business activity
        ].copy()

        # 🔍 Group by CNIC only
        grouped = filtered.groupby("CNIC_NUMBER_CLEAN")
        flagged_groups = grouped.filter(lambda x: len(x) > 1)

        # 📤 Final output
        output_cols = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "CUSTOMERFULLNAME",
            "CNIC_NUMBER",
            "CUSTSECTORCODE",
            "CUSTOMER_OCCUPATION",
            "CCY",
            "NAMEOFBUSINESS",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
            "TOTAL_COUNT",
            "TOTAL_AMOUNT",
        ]
        output = (
            flagged_groups[
                [col for col in output_cols if col in flagged_groups.columns]
            ]
            .drop_duplicates()
            .reset_index(drop=True)
        )

        # Rename to friendly names
        output = output.rename(
            columns={
                "CUSTOMER_NUMBER": "Customer_Number",
                "ACCOUNT_NUMBER": "Account_Number",
                "TITLEOFACCOUNT": "TitleOfAccount",
                "CUSTOMERFULLNAME": "CustomerFullName",
                "CNIC_NUMBER": "CNIC_Number",
                "CUSTSECTORCODE": "CustSectorCode",
                "CUSTOMER_OCCUPATION": "Customer_Occupation",
                "CCY": "CCY",
                "NAMEOFBUSINESS": "NameOfBusiness",
                "PRODUCTDESC": "ProductDesc",
                "SECTOR_DESCRIPTION": "Sector_Description",
                "TOTAL_COUNT": "Total_Count",
                "TOTAL_AMOUNT": "Total_Amount",
            }
        )

        if output.empty:
            output = pd.DataFrame([{col: pd.NA for col in output.columns}])

        # 📁 Optional export / UI mode
        if mode == "full":
            file_path = (
                f"multi_account_business_violation_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            )
            output.to_excel(file_path, index=False)
        elif mode == "ui":
            output = output[
                [
                    "Customer_Number",
                    "Account_Number",
                    "TitleOfAccount",
                    "CustomerFullName",
                    "CNIC_Number",
                    "CustSectorCode",
                    "Customer_Occupation",
                    "CCY",
                    "NameOfBusiness",
                    "ProductDesc",
                    "Sector_Description",
                    "Total_Count",
                    "Total_Amount",
                ]
            ]

        return output

    except Exception as e:
        print(f"❌ Error in logic_087_multi_account_business_use_violation: {e}")
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
                    "NameOfBusiness": pd.NA,
                    "ProductDesc": pd.NA,
                    "Sector_Description": pd.NA,
                    "Total_Count": pd.NA,
                    "Total_Amount": pd.NA,
                }
            ]
        )
