import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


def normalize_text(value: str) -> str:
    return str(value).strip().upper().replace(".", "").replace("/", "").replace(" ", "")


@register_logic(
    name="079_customer_profile_snapshot_missing",
    description="Flags customers whose Customer Profile (Snapshot) is blank, null, matches known nil values, or equals Title/Customer Name only.",
    category="Compliance & Screening",
)
def logic_079_customer_profile_snapshot_missing(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    try:
        # 🔍 Locate merged KYC file
        kyc_key = next((k for k in dataframes if "merged_file.xlsx" in k.lower()), None)
        if not kyc_key:
            raise ValueError("Merged KYC file not found.")
        df_main = dataframes[kyc_key]
        if isinstance(df_main, tuple):
            df_main = next(iter(df_main[0].values())) if df_main[0] else pd.DataFrame()

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
            "ACCOUNT_OPEN_DT",
            "KYCRISK",
            "CUSTOMERPROFILE",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        missing = [col for col in required_cols if col not in df_main.columns]
        if missing:
            raise ValueError(f"Missing required columns: {missing}")

        df = df_main[required_cols].copy()

        # 🔍 Normalize key fields
        df["CUSTOMERPROFILE_CLEAN"] = df["CUSTOMERPROFILE"].fillna("").apply(normalize_text)
        df["TITLEOFACCOUNT_CLEAN"] = df["TITLEOFACCOUNT"].fillna("").apply(normalize_text)
        df["CUSTOMERFULLNAME_CLEAN"] = df["CUSTOMERFULLNAME"].fillna("").apply(normalize_text)

        # 🔍 Load NILL_COMBINATIONS reference
        raw_nill_values = set()
        for df_nill in dataframes.values():
            if isinstance(df_nill, dict):
                for sheet in df_nill.values():
                    if isinstance(sheet, pd.DataFrame):
                        sheet.columns = (
                            sheet.columns.str.strip()
                            .str.upper()
                            .str.replace(" ", "_")
                            .str.replace("-", "_")
                        )
                        if "NILL_COMBINATIONS" in sheet.columns:
                            raw_nill_values.update(
                                sheet["NILL_COMBINATIONS"]
                                .dropna()
                                .astype(str)
                                .str.strip()
                            )
            elif isinstance(df_nill, pd.DataFrame):
                df_nill.columns = (
                    df_nill.columns.str.strip()
                    .str.upper()
                    .str.replace(" ", "_")
                    .str.replace("-", "_")
                )
                if "NILL_COMBINATIONS" in df_nill.columns:
                    raw_nill_values.update(
                        df_nill["NILL_COMBINATIONS"].dropna().astype(str).str.strip()
                    )

        if not raw_nill_values:
            raise ValueError("NILL_COMBINATIONS column not found in any uploaded file.")

        nill_set_normalized = {normalize_text(val) for val in raw_nill_values}

        # 🧠 Additional invalid-profile conditions:
        # CustomerProfile is just TitleOfAccount OR just CustomerFullName
        profile_equals_title = (
            df["CUSTOMERPROFILE_CLEAN"].ne("") &
            df["CUSTOMERPROFILE_CLEAN"].eq(df["TITLEOFACCOUNT_CLEAN"])
        )
        profile_equals_name = (
            df["CUSTOMERPROFILE_CLEAN"].ne("") &
            df["CUSTOMERPROFILE_CLEAN"].eq(df["CUSTOMERFULLNAME_CLEAN"])
        )

        # 🧠 Apply contradiction logic
        df["CONTRADICTION"] = df["ACCOUNT_NUMBER"].notna() & (
            df["CUSTOMERPROFILE"].isna()
            | df["CUSTOMERPROFILE"].astype(str).str.strip().eq("")
            | df["CUSTOMERPROFILE_CLEAN"].isin(nill_set_normalized)
            | profile_equals_title
            | profile_equals_name
        )

        # 📤 Prepare output (select raw cols, then rename)
        output = df[df["CONTRADICTION"]][
            [
                "CUSTOMER_NUMBER",
                "ACCOUNT_NUMBER",
                "TITLEOFACCOUNT",
                "CUSTOMERFULLNAME",
                "CNIC_NUMBER",
                "ACCOUNT_OPEN_DT",
                "KYCRISK",
                "CUSTOMERPROFILE",
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
            "Account_Open_Date",
            "KYCRisk",
            "CustomerProfile",
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
                        "Account_Open_Date": pd.NA,
                        "KYCRisk": pd.NA,
                        "CustomerProfile": pd.NA,
                        "ProductDesc": pd.NA,
                        "Sector_Description": pd.NA,
                    }
                ]
            )

        # 📁 Optional export
        if mode == "full":
            file_path = f"customer_profile_snapshot_missing_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            output.to_excel(file_path, index=False)
        elif mode == "ui":
            output = output[
                [
                    "Customer_Number",
                    "Account_Number",
                    "TitleOfAccount",
                    "CustomerFullName",
                    "CNIC_Number",
                    "Account_Open_Date",
                    "KYCRisk",
                    "CustomerProfile",
                    "ProductDesc",
                    "Sector_Description",
                ]
            ]

        return output

    except Exception as e:
        print(f"❌ Error in logic_079_customer_profile_snapshot_missing: {e}")
        return pd.DataFrame(
            [
                {
                    "Customer_Number": pd.NA,
                    "Account_Number": pd.NA,
                    "TitleOfAccount": pd.NA,
                    "CustomerFullName": pd.NA,
                    "CNIC_Number": pd.NA,
                    "Account_Open_Date": pd.NA,
                    "KYCRisk": pd.NA,
                    "CustomerProfile": pd.NA,
                    "ProductDesc": pd.NA,
                    "Sector_Description": pd.NA,
                }
            ]
        )
