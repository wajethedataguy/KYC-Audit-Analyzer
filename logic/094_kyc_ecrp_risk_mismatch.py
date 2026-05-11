import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="094_kyc_ecrp_Risk_mismatch",
    description="Flags customers whose KYC risk level does not match eCRP risk level.",
    category="CDD & EDD Review",
)
def logic_094_kyc_ecrp_risk_mismatch(dataframes: dict, mode="full") -> pd.DataFrame:
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

        # 🔧 Define risk standardization function
        def standardize_risk(risk):
            if pd.isna(risk):
                return ""
            risk = str(risk).lower().replace(".", "").strip()
            if any(tag in risk for tag in ["mandatory", "mdtry", "mn"]):
                return "high"
            elif "medium" in risk:
                return "medium"
            elif "low" in risk:
                return "low"
            return risk

        # 🧠 Apply standardization
        df_kyc["KYCRISK_CLEAN"] = df_kyc["KYCRISK"].apply(standardize_risk)
        df_kyc["ECRP_RISK_CLEAN"] = df_kyc["ECRP_RISK"].apply(standardize_risk)

        # 🧠 Apply contradiction logic
        contradiction_mask = (
            (df_kyc["KYCRISK_CLEAN"] != df_kyc["ECRP_RISK_CLEAN"])
            & (df_kyc["KYCRISK_CLEAN"] != "")
            & (df_kyc["ECRP_RISK_CLEAN"] != "")
        )

        # 📤 Final output
        output_cols = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "KYCRISK",
            "ECRP_RISK",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        output = (
            df_kyc[contradiction_mask][output_cols]
            .drop_duplicates()
            .reset_index(drop=True)
        )

        # 🧯 Handle empty output
        if output.empty:
            output = pd.DataFrame([{col: pd.NA for col in output_cols}])

        # 📁 Optional export
        if mode == "full":
            file_path = f"kyc_ecrp_risk_mismatch_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            output.to_excel(file_path, index=False)
        elif mode == "ui":
            output = output[
                [
                    "CUSTOMER_NUMBER",
                    "ACCOUNT_NUMBER",
                    "TITLEOFACCOUNT",
                    "KYCRISK",
                    "ECRP_RISK",
                    "PRODUCTDESC",
                    "SECTOR_DESCRIPTION",
                ]
            ]

        return output

    except Exception as e:
        print(f"❌ Error in logic_094_kyc_ecrp_risk_mismatch: {e}")
        return pd.DataFrame(
            [
                {
                    col: pd.NA
                    for col in [
                        "CUSTOMER_NUMBER",
                        "ACCOUNT_NUMBER",
                        "TITLEOFACCOUNT",
                        "KYCRISK",
                        "ECRP_RISK",
                        "PRODUCTDESC",
                        "SECTOR_DESCRIPTION",
                    ]
                }
            ]
        )
