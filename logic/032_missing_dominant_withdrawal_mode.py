import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="032_missing_dominant_withdrawal_mode",
    description="Flags customers whose Dominant Mode of Withdrawal is blank in KYC profile.",
    category="Compliance & Screening",
)
def logic_032_missing_dominant_withdrawal_mode(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    try:
        # Locate merged file
        merged_key = next(
            (k for k in dataframes if "merged_file.xlsx" in k.lower()), None
        )
        if not merged_key:
            raise ValueError("Merged file not found.")

        df = dataframes[merged_key].copy()
        df.columns = (
            df.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )

        required_cols = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "DOMINANT_MODE_OF_WITHDRAWAL",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        missing = [col for col in required_cols if col not in df.columns]
        if missing:
            raise ValueError(f"Missing required columns: {missing}")

        # Normalize and trim (string view)
        df["DOMINANT_WITHDRAWAL_CLEAN"] = (
            df["DOMINANT_MODE_OF_WITHDRAWAL"].astype(str).str.strip()
        )

        # Treat as missing if:
        #   - original value is NaN OR
        #   - stripped text is empty
        mask_missing = df["DOMINANT_MODE_OF_WITHDRAWAL"].isna() | (
            df["DOMINANT_WITHDRAWAL_CLEAN"] == ""
        )

        df_filtered = df[mask_missing].copy()

        # Prepare output
        output = df_filtered[
            [
                "CUSTOMER_NUMBER",
                "ACCOUNT_NUMBER",
                "TITLEOFACCOUNT",
                "DOMINANT_MODE_OF_WITHDRAWAL",
                "PRODUCTDESC",
                "SECTOR_DESCRIPTION",
            ]
        ].copy()

        output.columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "Dominant_Mode_of_Withdrawal",
            "ProductDesc",
            "SECTOR_DESCRIPTION",
        ]
        output = output.reset_index(drop=True)

        # Fallback row if empty
        if output.empty:
            output = pd.DataFrame(
                [
                    {
                        "Customer_Number": pd.NA,
                        "Account_Number": pd.NA,
                        "TitleOfAccount": pd.NA,
                        "Dominant_Mode_of_Withdrawal": pd.NA,
                        "ProductDesc": pd.NA,
                        "SECTOR_DESCRIPTION": pd.NA,
                    }
                ]
            )

        # Optional export
        if mode == "full":
            file_path = (
                f"missing_dominant_withdrawal_mode_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            )
            output.to_excel(file_path, index=False)

        return output

    except Exception as e:
        print(f"❌ Exception occurred: {e}")
        fallback_columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "Dominant_Mode_of_Withdrawal",
            "ProductDesc",
            "SECTOR_DESCRIPTION",
        ]
        return pd.DataFrame(columns=fallback_columns)
