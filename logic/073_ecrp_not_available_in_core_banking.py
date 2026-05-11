import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


def normalize_text(value: str) -> str:
    return str(value).strip().upper().replace(".", "").replace("/", "").replace(" ", "")


@register_logic(
    name="073_ecrp_not_available_in_core_banking",
    description="Flags customers whose eCRP risk profile is blank, null, or matches known nil combinations.",
    category="Compliance & Screening",
)
def logic_073_ecrp_not_available_in_core_banking(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    try:
        # 🔍 Locate merged KYC file
        kyc_key = next((k for k in dataframes if "merged_file.xlsx" in k.lower()), None)
        if not kyc_key:
            raise ValueError("KYC profile file not found.")
        df_main = dataframes[kyc_key].copy()
        df_main.columns = (
            df_main.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )

        required_cols = [
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "CUSTOMER_NUMBER",
            "CUSTOMERFULLNAME",
            "CNIC_NUMBER",
            "POLITICALFIGURE",
            "KYCRISK",
            "ECRP_RISK",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        missing = [col for col in required_cols if col not in df_main.columns]
        if missing:
            raise ValueError(f"Merged file missing required columns: {missing}")

        df = df_main[required_cols].copy()

        # 🔍 Normalize eCRP_RISK field
        df["ECRP_RISK_NORMALIZED"] = df["ECRP_RISK"].fillna("").apply(normalize_text)

        # 🔍 Load NILL_COMBINATIONS reference from uploaded files
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

        nill_set_normalized = {normalize_text(val) for val in raw_nill_values}

        # 🧠 Apply contradiction logic
        df_filtered = df[
            df["ECRP_RISK_NORMALIZED"].isin(nill_set_normalized)
            & df["ACCOUNT_NUMBER"].notna()
        ].copy()

        # 📤 Prepare output
        output = df_filtered[
            [
                "CUSTOMER_NUMBER",
                "ACCOUNT_NUMBER",
                "TITLEOFACCOUNT",
                "PRODUCTDESC",
                "SECTOR_DESCRIPTION",
            ]
        ].copy()

        output.columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
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
                        "ProductDesc": pd.NA,
                        "SECTOR_DESCRIPTION": pd.NA,
                    }
                ]
            )

        # 📁 Optional export
        if mode == "full":
            file_path = f"ecrp_not_available_in_core_banking_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
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
                    "ProductDesc": pd.NA,
                    "SECTOR_DESCRIPTION": pd.NA,
                }
            ]
        )
