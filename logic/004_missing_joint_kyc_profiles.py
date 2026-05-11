import pandas as pd
import unicodedata
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="004_missing_joint_kyc_profiles",
    description="Flags personal accounts with joint indicators in title but no joint holder declared, matching Power Query logic.",
    category="Customer Name Filter",
)
def logic_004_missing_joint_kyc_profiles(dataframes: dict, mode="full") -> pd.DataFrame:
    try:
        # Load merged file
        merged_key = next(
            (k for k in dataframes if "merged_file.xlsx" in k.lower()), None
        )
        df = dataframes.get(merged_key)
        if isinstance(df, tuple):
            df = next(iter(df[0].values())) if df[0] else pd.DataFrame()
        if df is None or df.empty:
            raise ValueError("Merged file not found or empty.")

        # Normalize columns
        df.columns = (
            df.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )

        # Required columns
        required = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "CUSTSECTORCODE",
            "JOINT_HOLDER",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        missing = [col for col in required if col not in df.columns]
        if missing:
            raise ValueError(f"Missing required columns: {missing}")

        # Normalize fields
        df["CUSTSECTORCODE"] = pd.to_numeric(df["CUSTSECTORCODE"], errors="coerce")
        df["JOINT_HOLDER"] = df["JOINT_HOLDER"].fillna("").astype(str).str.strip()
        df["TITLEOFACCOUNT"] = (
            df["TITLEOFACCOUNT"]
            .fillna("")
            .astype(str)
            .apply(
                lambda x: unicodedata.normalize(
                    "NFKC", x.strip().upper().replace("\\", "/")
                )
            )
        )

        # Joint indicators (inclusive)
        joint_indicators = [
            "/",
            " AND ",
            " OR ",
            " & ",
            " S/O ",
            " D/O ",
        ]

        # Exclusion patterns (exclusive)
        exclusion_patterns = [" S/O ", " D/O ", " W/O "]

        # Apply Power Query logic
        df_filtered = df[
            (df["CUSTSECTORCODE"] <= 1005)
            & (df["JOINT_HOLDER"] == "")
            & df["TITLEOFACCOUNT"].apply(
                lambda title: any(p in title for p in joint_indicators)
            )
            & df["TITLEOFACCOUNT"].apply(
                lambda title: all(p not in title for p in exclusion_patterns)
            )
        ].copy()

        # Deduplicate to match Power Query behavior
        df_filtered = df_filtered.drop_duplicates(
            subset=["ACCOUNT_NUMBER", "TITLEOFACCOUNT"]
        )

        # Output
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
            "Sector_Description",
        ]
        output = output.reset_index(drop=True)

        if output.empty:
            output = pd.DataFrame([{col: pd.NA for col in output.columns}])

        if mode == "full":
            file_path = (
                f"missing_joint_kyc_profiles_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            )
            output.to_excel(file_path, index=False)
        elif mode == "ui":
            output = output[
                [
                    "Customer_Number",
                    "Account_Number",
                    "TitleOfAccount",
                    "ProductDesc",
                    "Sector_Description",
                ]
            ]

        return output

    except Exception as e:
        print(f"❌ Error in logic_004_missing_joint_kyc_profiles: {e}")
        return pd.DataFrame(
            [
                {
                    col: pd.NA
                    for col in [
                        "Customer_Number",
                        "Account_Number",
                        "TitleOfAccount",
                        "ProductDesc",
                        "Sector_Description",
                    ]
                }
            ]
        )
