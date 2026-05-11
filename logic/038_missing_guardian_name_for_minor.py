import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="038_missing_guardian_name_for_minor",
    description="Flags minor customers (sector code 1006) whose guardian name is missing in KYC profile.",
    category="Compliance & Screening",
)
def logic_038_missing_guardian_name_for_minor(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    try:
        # Load KYC file
        kyc_key = next((k for k in dataframes if "merged_file.xlsx" in k.lower()), None)
        kyc_candidates = [kyc_key] if kyc_key else list(dataframes.keys())
        df_kyc = None

        for key in kyc_candidates:
            df = dataframes[key].copy()
            df.columns = (
                df.columns.str.strip()
                .str.upper()
                .str.replace(" ", "_")
                .str.replace("-", "_")
            )
            required_cols = [
                "ACCOUNT_NUMBER",
                "TITLEOFACCOUNT",
                "CUSTOMER_NUMBER",
                "CUSTOMERFULLNAME",
                "CUSTSECTORCODE",
                "RELATIONWITHMINOR",
                "PRODUCTDESC",
                "SECTOR_DESCRIPTION",
            ]
            if all(col in df.columns for col in required_cols):
                df_kyc = df
                break

        if df_kyc is None:
            raise ValueError("No file contains required KYC columns.")

        # Clean and filter
        df_kyc["CUSTSECTORCODE_CLEAN"] = pd.to_numeric(
            df_kyc["CUSTSECTORCODE"], errors="coerce"
        )

        guardian_name_col = (
            "GUARDIAN_NAME2" if "GUARDIAN_NAME2" in df_kyc.columns else None
        )
        if guardian_name_col:
            df_kyc["GUARDIAN_NAME2_CLEAN"] = (
                df_kyc[guardian_name_col].fillna("").astype(str).str.strip().str.lower()
            )
        else:
            df_kyc["GUARDIAN_NAME2_CLEAN"] = ""

        # Flag minors with missing guardian name
        df_flagged = df_kyc[
            (df_kyc["CUSTSECTORCODE_CLEAN"] == 1006)
            & (df_kyc["GUARDIAN_NAME2_CLEAN"] == "")
        ].copy()

        output_cols = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "CUSTOMERFULLNAME",
            "CUSTSECTORCODE",
            guardian_name_col if guardian_name_col else None,
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        output_cols = [
            col for col in output_cols if col
        ]  # drop None if guardian missing

        output = df_flagged[output_cols].copy()

        rename_map = {
            "CUSTOMER_NUMBER": "Customer_Number",
            "ACCOUNT_NUMBER": "Account_Number",
            "TITLEOFACCOUNT": "TitleOfAccount",
            "CUSTOMERFULLNAME": "CustomerFullName",
            "CUSTSECTORCODE": "CustSectorCode",
            "PRODUCTDESC": "ProductDesc",
            "SECTOR_DESCRIPTION": "Sector_Description",
        }
        if guardian_name_col:
            rename_map[guardian_name_col] = "Guardian_Name"

        output = output.rename(columns=rename_map).reset_index(drop=True)

        # Fallback for empty output
        if output.empty:
            output = pd.DataFrame([{col: pd.NA for col in rename_map.values()}])

        # Export control
        if mode == "full":
            file_path = (
                f"missing_guardian_name_for_minor_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            )
            output.to_excel(file_path, index=False)

        return output

    except Exception as e:
        print(f"❌ Error in logic_038_missing_guardian_name_for_minor: {e}")
        return pd.DataFrame(
            [
                {
                    "Customer_Number": pd.NA,
                    "Account_Number": pd.NA,
                    "TitleOfAccount": pd.NA,
                    "CustomerFullName": pd.NA,
                    "CustSectorCode": pd.NA,
                    "Guardian_Name": pd.NA,
                    "ProductDesc": pd.NA,
                    "Sector_Description": pd.NA,
                }
            ]
        )
