import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


def normalize_position(value: str) -> str:
    return str(value).strip().upper().replace(".", "").replace(" ", "").replace("/", "")


@register_logic(
    name="051_missing_title_or_position_for_salaried_individuals",
    description="Flags salaried individuals with missing or vague CustomerPosition using Nill_Comb reference list.",
    category="Compliance & Screening",
)
def logic_051_missing_title_or_position_for_salaried_individuals(
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
            "CUSTSECTORCODE",
            "CUSTOMER_OCCUPATION",
            "CUSTOMERPOSITION",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        missing = [col for col in required_cols if col not in df_main.columns]
        if missing:
            raise ValueError(f"Merged file missing required columns: {missing}")

        df = df_main[required_cols].copy()

        # 🔍 Load Nill_Combinations reference
        raw_nill_values = set()
        for df_nill in dataframes.values():
            if isinstance(df_nill, dict):
                for sheet in df_nill.values():
                    sheet.columns = (
                        sheet.columns.str.strip()
                        .str.upper()
                        .str.replace(" ", "_")
                        .str.replace("-", "_")
                    )
                    if "NILL_COMBINATIONS" in sheet.columns:
                        raw_nill_values.update(
                            sheet["NILL_COMBINATIONS"].dropna().astype(str).str.strip()
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

        nill_set_normalized = {normalize_position(val) for val in raw_nill_values}

        # 🔍 Normalize fields
        df["CUSTSECTORCODE_NUMERIC"] = pd.to_numeric(
            df["CUSTSECTORCODE"], errors="coerce"
        )
        df["CUSTOMER_OCCUPATION_CLEAN"] = (
            df["CUSTOMER_OCCUPATION"].astype(str).str.strip().str.upper()
        )
        df["CUSTOMERPOSITION_CLEAN"] = df["CUSTOMERPOSITION"].astype(str).str.strip()
        df["CUSTOMERPOSITION_NORMALIZED"] = df["CUSTOMERPOSITION_CLEAN"].apply(
            normalize_position
        )

        # 🧠 Apply contradiction logic
        df_filtered = df[
            (
                (df["CUSTOMERPOSITION_CLEAN"] == "")
                | (df["CUSTOMERPOSITION_CLEAN"].isna())
                | (df["CUSTOMERPOSITION_NORMALIZED"].isin(nill_set_normalized))
            )
            & (df["CUSTSECTORCODE_NUMERIC"] <= 1005)
            & (df["CUSTOMER_OCCUPATION_CLEAN"] == "SALARIED")
        ].copy()

        # 📤 Prepare output (select raw cols, then rename)
        output = df_filtered[
            [
                "CUSTOMER_NUMBER",
                "ACCOUNT_NUMBER",
                "TITLEOFACCOUNT",
                "CUSTOMERFULLNAME",
                "CUSTSECTORCODE",
                "CUSTOMER_OCCUPATION",
                "CUSTOMERPOSITION",
                "PRODUCTDESC",
                "SECTOR_DESCRIPTION",
            ]
        ].copy()

        output.columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "CustomerFullName",
            "CustSectorCode",
            "Customer_Occupation",
            "CustomerPosition",
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
                        "CustSectorCode": pd.NA,
                        "Customer_Occupation": pd.NA,
                        "CustomerPosition": pd.NA,
                        "ProductDesc": pd.NA,
                        "Sector_Description": pd.NA,
                    }
                ]
            )

        # 📁 Optional export
        if mode == "full":
            file_path = f"missing_title_or_position_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
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
                    "CustSectorCode": pd.NA,
                    "Customer_Occupation": pd.NA,
                    "CustomerPosition": pd.NA,
                    "ProductDesc": pd.NA,
                    "Sector_Description": pd.NA,
                }
            ]
        )
