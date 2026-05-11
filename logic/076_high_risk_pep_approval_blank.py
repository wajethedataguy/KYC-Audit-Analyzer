import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="076_high_risk_pep_approval_check",
    description="Flags high-risk customers where ApprovalObtained is blank or marked as 'No'.",
    category="Compliance & Screening",
)
def logic_076_high_risk_pep_approval_check(
    dataframes: dict, mode: str = "full"
) -> pd.DataFrame:
    try:
        # 🔍 Locate merged KYC file
        kyc_key = next((k for k in dataframes if "merged_file.xlsx" in k.lower()), None)
        if not kyc_key:
            raise ValueError("Merged KYC file not found.")
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
            "KYCRISK",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
            "APPROVALOBTAINED",  # <-- use this column now
        ]
        missing = [col for col in required_cols if col not in df_main.columns]
        if missing:
            raise ValueError(f"Missing required columns in merged file: {missing}")

        df = df_main.copy()

        # 🔍 Normalize fields
        df["KYCRISK_CLEAN"] = df["KYCRISK"].astype(str).str.strip().str.lower()
        df["APPROVAL_CLEAN"] = (
            df["APPROVALOBTAINED"]
            .astype(str)
            .str.strip()
            .str.lower()
            .replace(
                {
                    "": "",
                    "nan": "",
                    "none": "",
                    "null": "",
                    "n": "no",
                    "0": "no",
                    "no": "no",
                    "false": "no",
                    "not approved": "no",
                    "pending": "no",
                    "n/a": "no",
                    "na": "no",
                }
            )
        )

        # 🧠 Apply contradiction logic: High risk + approval blank or "no"
        mask = df["KYCRISK_CLEAN"].str.contains("high", na=False) & df[
            "APPROVAL_CLEAN"
        ].isin(["", "no"])
        df_filtered = df.loc[mask].copy()
        df_filtered["CONTRADICTION_REASON"] = (
            "High risk approvals were left blank or marked as 'No' in KYC profiles of high risk customers."
        )

        # 📤 Prepare output
        final_cols = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "KYCRISK",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
            "CONTRADICTION_REASON",
        ]
        output = df_filtered[final_cols].reset_index(drop=True)

        # 🧯 Handle empty output
        if output.empty:
            output = pd.DataFrame([{col: pd.NA for col in final_cols}])

        # 📁 Optional export
        if mode == "full" and not output.empty:
            file_path = (
                f"high_risk_pep_approval_check_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            )
            output.to_excel(file_path, index=False)
        elif mode == "ui" and not output.empty:
            output = output[
                [
                    "CUSTOMER_NUMBER",
                    "ACCOUNT_NUMBER",
                    "TITLEOFACCOUNT",
                    "KYCRISK",
                    "PRODUCTDESC",
                    "SECTOR_DESCRIPTION",
                ]
            ]

        return output

    except Exception as e:
        print(f"❌ Error in logic_076_high_risk_pep_approval_check: {e}")
        return pd.DataFrame(
            [
                {
                    col: pd.NA
                    for col in [
                        "CUSTOMER_NUMBER",
                        "ACCOUNT_NUMBER",
                        "TITLEOFACCOUNT",
                        "KYCRISK",
                        "PRODUCTDESC",
                        "SECTOR_DESCRIPTION",
                        "CONTRADICTION_REASON",
                    ]
                }
            ]
        )
