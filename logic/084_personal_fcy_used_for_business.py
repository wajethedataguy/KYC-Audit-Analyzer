import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="084_personal_fcy_used_for_business",
    description="Flags personal foreign currency accounts used for business/commercial purposes in violation of Compliance Operational Manual.",
    category="CDD & EDD Review",
)
def logic_084_personal_fcy_used_for_business(
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
            "CUSTSECTORCODE",
            "CUSTOMER_OCCUPATION",
            "CCY",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        missing = [col for col in required_cols if col not in df_kyc.columns]
        if missing:
            raise ValueError(f"Missing required columns: {missing}")

        # 🔧 Normalize account numbers
        df_kyc["ACCOUNT_NUMBER_RAW"] = df_kyc["ACCOUNT_NUMBER"]
        df_kyc["ACCOUNT_NUMBER"] = (
            pd.to_numeric(df_kyc["ACCOUNT_NUMBER_RAW"], errors="coerce")
            .astype("Int64")
            .astype(str)
        )

        # 📊 Collect transactional fragments
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

        # 🔍 Normalize fields
        df_merged["CUSTSECTORCODE_CLEAN"] = pd.to_numeric(
            df_merged["CUSTSECTORCODE"], errors="coerce"
        )
        df_merged["CCY_CLEAN"] = df_merged["CCY"].astype(str).str.strip().str.upper()
        df_merged["TOTAL_COUNT"] = pd.to_numeric(
            df_merged["TOTAL_COUNT"], errors="coerce"
        )

        # 🧠 Apply contradiction logic
        contradiction_mask = (
            (df_merged["CUSTSECTORCODE_CLEAN"] == 1000)  # personal accounts
            & (df_merged["CCY_CLEAN"] != "PKR")  # foreign currency
            & (
                df_merged["CUSTOMER_OCCUPATION"].astype(str).str.strip().str.lower()
                == "business"
            )  # declared business use
            & (df_merged["TOTAL_COUNT"] > 30)  # threshold for activity
        )

        # 📤 Final output (select raw cols, then rename)
        output_cols = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "CUSTOMERFULLNAME",
            "CUSTSECTORCODE",
            "CUSTOMER_OCCUPATION",
            "CCY",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
            "TOTAL_COUNT",
            "TOTAL_AMOUNT",
        ]
        output = (
            df_merged[contradiction_mask][output_cols]
            .drop_duplicates()
            .reset_index(drop=True)
        )

        output.columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "CustomerFullName",
            "CustSectorCode",
            "Customer_Occupation",
            "CCY",
            "ProductDesc",
            "Sector_Description",
            "Total_Transactions",
            "Total_Amount",
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
                        "CustSectorCode": pd.NA,
                        "Customer_Occupation": pd.NA,
                        "CCY": pd.NA,
                        "ProductDesc": pd.NA,
                        "Sector_Description": pd.NA,
                        "Total_Transactions": pd.NA,
                        "Total_Amount": pd.NA,
                    }
                ]
            )

        # 📁 Optional export
        if mode == "full":
            file_path = f"personal_fcy_misuse_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
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
        print(f"❌ Error in logic_084_personal_fcy_used_for_business: {e}")
        return pd.DataFrame(
            [
                {
                    "Customer_Number": pd.NA,
                    "Account_Number": pd.NA,
                    "TitleOfAccount": pd.NA,
                    "CustomerFullName": pd.NA,
                    "CustSectorCode": pd.NA,
                    "Customer_Occupation": pd.NA,
                    "CCY": pd.NA,
                    "ProductDesc": pd.NA,
                    "Sector_Description": pd.NA,
                    "Total_Transactions": pd.NA,
                    "Total_Amount": pd.NA,
                }
            ]
        )
