import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


def normalize(val: str) -> str:
    return str(val).strip().lower()


def normalize_account_number(val) -> str:
    try:
        return str(int(float(val))).strip().lower()
    except:
        return str(val).strip().lower()


def find_column(df, target_name: str) -> str:
    for col in df.columns:
        if normalize(col) == normalize(target_name):
            return col
    return None


def find_edd_status_column(df) -> str:
    for col in df.columns:
        if "edd" in normalize(col):
            return col
    return None


@register_logic(
    name="074_edd_not_made_for_high_risk_customers",
    description="Flags high-risk customers whose EDD was not made or available in the core banking system.",
    category="Compliance & Screening",
)
def logic_074_edd_not_made_for_high_risk_customers(
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
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "CUSTOMER_NUMBER",
            "CUSTOMERFULLNAME",
            "CNIC_NUMBER",
            "POLITICALFIGURE",
            "KYCRISK",
            "ECRP_RISK",
            "ACCOUNT_OPEN_DT",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        missing = [col for col in required_cols if col not in df_main.columns]
        if missing:
            raise ValueError(f"Missing required columns: {missing}")

        df = df_main[required_cols].copy()

        # 🔍 Normalize fields
        df["ACCOUNT_NUMBER_NORM"] = df["ACCOUNT_NUMBER"].apply(normalize_account_number)
        df["KYCRISK_NORM"] = (
            df["KYCRISK"].fillna("").astype(str).str.strip().str.lower()
        )
        df["ECRP_RISK_NORM"] = (
            df["ECRP_RISK"].fillna("").astype(str).str.strip().str.lower()
        )

        # 🔍 Collect EDD accounts from reference sheets
        edd_accounts = set()
        for key, df_edd in dataframes.items():
            if "edd" in key.lower():
                sheets = (
                    df_edd.items() if isinstance(df_edd, dict) else [("Sheet1", df_edd)]
                )
                for _, sheet in sheets:
                    sheet.columns = sheet.columns.str.strip()
                    acc_col = find_column(sheet, "Account_Number")
                    status_col = find_edd_status_column(sheet)
                    if acc_col:
                        if status_col:
                            filtered = sheet[
                                sheet[status_col]
                                .fillna("")
                                .astype(str)
                                .str.strip()
                                .str.lower()
                                .str.contains("not|required|missing")
                            ]
                            edd_accounts.update(
                                filtered[acc_col]
                                .dropna()
                                .apply(normalize_account_number)
                            )
                        else:
                            if find_column(sheet, "KYC_Risk") or find_column(
                                sheet, "eCRP_Risk"
                            ):
                                edd_accounts.update(
                                    sheet[acc_col]
                                    .dropna()
                                    .apply(normalize_account_number)
                                )

        # 🧠 Apply contradiction logic
        df_flagged = df[
            (
                df["KYCRISK_NORM"].str.contains("high")
                | df["ECRP_RISK_NORM"].str.contains("high")
            )
            & df["ACCOUNT_NUMBER_NORM"].isin(edd_accounts)
        ].copy()

        # 📤 Prepare output
        output = df_flagged[
            [
                "CUSTOMER_NUMBER",
                "ACCOUNT_NUMBER",
                "TITLEOFACCOUNT",
                "ACCOUNT_OPEN_DT",
                "KYCRISK",
                "PRODUCTDESC",
                "SECTOR_DESCRIPTION",
            ]
        ].copy()

        output.columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "Account_Open_Dt",
            "KYCRISK",
            "ProductDesc",
            "SECTOR_DESCRIPTION",
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
                        "Account_Open_Dt": pd.NA,
                        "KYCRISK": pd.NA,
                        "ProductDesc": pd.NA,
                        "SECTOR_DESCRIPTION": pd.NA,
                    }
                ]
            )

        # 📁 Optional export
        if mode == "full":
            file_path = f"edd_not_made_for_high_risk_customers_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
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
                    "Account_Open_Dt": pd.NA,
                    "KYCRISK": pd.NA,
                    "ProductDesc": pd.NA,
                    "SECTOR_DESCRIPTION": pd.NA,
                }
            ]
        )
