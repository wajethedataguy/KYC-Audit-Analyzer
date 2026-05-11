import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="088_personal_used_despite_company_account",
    description="Flags company PKR accounts when personal accounts are used for business purposes despite availability of company accounts, violating Compliance Operational Manual.",
    category="CDD & EDD Review",
)
def logic_088_personal_used_despite_company_account(
    dataframes: dict, mode: str = "full"
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

        # 🔧 Normalize key KYC fields
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
        # Normalize sector codes: convert to integer if possible
        df_kyc["CUSTSECTORCODE_CLEAN"] = (
            pd.to_numeric(df_kyc["CUSTSECTORCODE"], errors="coerce")
            .fillna(-1)
            .astype(int)
        )

        # 📊 Load FT Transaction Summary from all transaction files
        insight_frames = []
        for k, df in dataframes.items():
            if k == kyc_key:
                continue  # skip KYC itself

            if isinstance(df, tuple):
                df = next(iter(df[0].values())) if df[0] else pd.DataFrame()
            if df is None or df.empty:
                continue

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
                    columns={
                        "CR_ACCOUNT": "ACCOUNT",
                        "TRANSACTION_AMOUNT": "TR_AMOUNT",
                    }
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
            .str.strip()
            .str.replace(r"[^\d.]", "", regex=True),
            errors="coerce",
        )
        df_ft_summary = df_ft_summary.groupby("ACCOUNT", as_index=False).agg(
            TOTAL_COUNT=("TR_AMOUNT", "count"),
            TOTAL_AMOUNT=("TR_AMOUNT", "sum"),
        )

        # 🔗 Attach FT summary to KYC
        df_merged = pd.merge(
            df_kyc,
            df_ft_summary,
            left_on="ACCOUNT_NUMBER",
            right_on="ACCOUNT",
            how="left",
        )

        # Only PKR accounts
        df_pkr = df_merged[df_merged["CCY_CLEAN"] == "PKR"].copy()

        # Sort for stable CNIC-wise output
        df_pkr = df_pkr.sort_values(
            by=["CNIC_NUMBER_CLEAN", "CUSTSECTORCODE_CLEAN", "ACCOUNT_NUMBER"]
        )

        # 🧠 Group by CNIC and build layout
        grouped = df_pkr.groupby("CNIC_NUMBER_CLEAN", dropna=True)

        layout_rows = []

        for _, group in grouped:
            personal = group[
                (group["CUSTSECTORCODE_CLEAN"] == 1000)
                & (group["CUSTOMER_OCCUPATION_CLEAN"] == "business")
                & (group["TOTAL_COUNT"] > 30)
            ]
            company = group[
                (group["CUSTSECTORCODE_CLEAN"] == 1100)
                & (group["CUSTOMER_OCCUPATION_CLEAN"] == "business")
            ]

            if personal.empty or company.empty:
                continue

            # include both personal and company rows
            for _, p_row in personal.iterrows():
                layout_rows.append(
                    {
                        "Account_Number": p_row["ACCOUNT_NUMBER"],
                        "TitleOfAccount": p_row["TITLEOFACCOUNT"],
                        "Total_Count": p_row["TOTAL_COUNT"],
                        "CustomerFullName": p_row["CUSTOMERFULLNAME"],
                        "CNIC_Number": p_row["CNIC_NUMBER"],
                        "CustSectorCode": p_row["CUSTSECTORCODE"],
                        "Customer_Occupation": p_row["CUSTOMER_OCCUPATION"],
                        "ProductDesc": p_row["PRODUCTDESC"],
                        "Sector_Description": p_row["SECTOR_DESCRIPTION"],
                        "Account_Type": "Personal",
                    }
                )

            for _, c_row in company.iterrows():
                layout_rows.append(
                    {
                        "Account_Number": c_row["ACCOUNT_NUMBER"],
                        "TitleOfAccount": c_row["TITLEOFACCOUNT"],
                        "Total_Count": c_row["TOTAL_COUNT"],
                        "CustomerFullName": c_row["CUSTOMERFULLNAME"],
                        "CNIC_Number": c_row["CNIC_NUMBER"],
                        "CustSectorCode": c_row["CUSTSECTORCODE"],
                        "Customer_Occupation": c_row["CUSTOMER_OCCUPATION"],
                        "ProductDesc": c_row["PRODUCTDESC"],
                        "Sector_Description": c_row["SECTOR_DESCRIPTION"],
                        "Account_Type": "Company",
                    }
                )

        # 🧯 Build final output frame
        if not layout_rows:
            output = pd.DataFrame(
                [
                    {
                        "Account_Number": pd.NA,
                        "TitleOfAccount": pd.NA,
                        "Total_Count": pd.NA,
                        "CustomerFullName": pd.NA,
                        "CNIC_Number": pd.NA,
                        "CustSectorCode": pd.NA,
                        "Customer_Occupation": pd.NA,
                        "ProductDesc": pd.NA,
                        "Sector_Description": pd.NA,
                        "Account_Type": pd.NA,
                    }
                ]
            )
        else:
            output = pd.DataFrame(layout_rows)

        # 📁 Export / UI mode handling
        if mode == "full":
            file_path = f"personal_used_despite_company_account_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            output.to_excel(file_path, index=False)
        elif mode == "ui":
            output = output[
                [
                    "Account_Number",
                    "TitleOfAccount",
                    "Total_Count",
                    "CustomerFullName",
                    "CNIC_Number",
                    "CustSectorCode",
                    "Customer_Occupation",
                    "ProductDesc",
                    "Sector_Description",
                    "Account_Type",
                ]
            ]
        # 🧯 Build final output frame
        if not layout_rows:
            output = pd.DataFrame(
                [
                    {
                        "Account_Number": pd.NA,
                        "TitleOfAccount": pd.NA,
                        "Total_Count": pd.NA,
                        "CustomerFullName": pd.NA,
                        "CNIC_Number": pd.NA,
                        "CustSectorCode": pd.NA,
                        "Customer_Occupation": pd.NA,
                        "ProductDesc": pd.NA,
                        "Sector_Description": pd.NA,
                        "Account_Type": pd.NA,
                    }
                ]
            )
        else:
            output = pd.DataFrame(layout_rows)

        # 📁 Export / UI mode handling
        if mode == "full":
            file_path = f"personal_used_despite_company_account_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            output.to_excel(file_path, index=False)
        elif mode == "ui":
            output = output[
                [
                    "Account_Number",
                    "TitleOfAccount",
                    "Total_Count",
                    "CustomerFullName",
                    "CNIC_Number",
                    "CustSectorCode",
                    "Customer_Occupation",
                    "ProductDesc",
                    "Sector_Description",
                    "Account_Type",
                ]
            ]

        return output

    except Exception as e:
        print(f"❌ Error in logic_088_personal_used_despite_company_account: {e}")
        return pd.DataFrame(
            [
                {
                    "Account_Number": pd.NA,
                    "TitleOfAccount": pd.NA,
                    "Total_Count": pd.NA,
                    "CustomerFullName": pd.NA,
                    "CNIC_Number": pd.NA,
                    "CustSectorCode": pd.NA,
                    "Customer_Occupation": pd.NA,
                    "ProductDesc": pd.NA,
                    "Sector_Description": pd.NA,
                    "Account_Type": pd.NA,
                }
            ]
        )
