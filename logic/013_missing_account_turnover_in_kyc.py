import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="013_missing_account_turnover_in_kyc",
    description="Flags accounts where turnover was not fed in KYC profile (blank, zero, or null) in both ACCOUNT_TURNOVER and KYC_Ann_TO_Corporate.",
    category="Data Completeness",
)
def logic_013_missing_account_turnover_in_kyc(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    try:
        merged_key = next(
            (k for k in dataframes if "merged_file.xlsx" in k.lower()), None
        )
        if merged_key is None:
            raise ValueError("Merged file not found in input dataframes.")
        df = dataframes[merged_key].copy()

        df.columns = (
            df.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )

        required_columns = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "PURPOSE_OF_ACCOUNT",
            "STATUS_OF_ACCOUNT",
            "ACCOUNT_TURNOVER",
            "KYC_ANN_TO_CORPORATE",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        missing = [col for col in required_columns if col not in df.columns]
        if missing:
            raise ValueError(f"Missing required columns: {missing}")

        # Normalize both turnover fields
        def normalize(val):
            return str(val).strip().upper().replace("  ", " ") if pd.notna(val) else ""

        df["TURNOVER_1"] = df["ACCOUNT_TURNOVER"].apply(normalize)
        df["TURNOVER_2"] = df["KYC_ANN_TO_CORPORATE"].apply(normalize)

        # Define valid labels and invalid indicators
        valid_labels = {"BELOW 1M", "1M TO 5M", "5M TO 10M", "ABOVE 10M", "ABOVE 50M"}
        invalid_values = {"", "0.00", "NAN", "NONE", "NULL"}

        # Flag rows where both fields are missing or invalid
        df_filtered = df[
            (
                ~df["TURNOVER_1"].isin(valid_labels)
                & df["TURNOVER_1"].isin(invalid_values)
            )
            & (
                ~df["TURNOVER_2"].isin(valid_labels)
                & df["TURNOVER_2"].isin(invalid_values)
            )
        ].copy()

        # Final output
        output = df_filtered[
            [
                "CUSTOMER_NUMBER",
                "ACCOUNT_NUMBER",
                "TITLEOFACCOUNT",
                "PURPOSE_OF_ACCOUNT",
                "STATUS_OF_ACCOUNT",
                "ACCOUNT_TURNOVER",
                "KYC_ANN_TO_CORPORATE",
                "PRODUCTDESC",
                "SECTOR_DESCRIPTION",
            ]
        ].copy()

        output.columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "Purpose_of_Account",
            "Status_Of_Account",
            "Account_Turnover",
            "KYC_Ann_TO_Corporate",
            "ProductDesc",
            "Sector_Description",
        ]

        output = output.dropna(how="all")

        output = output[
            ~(
                output[["Customer_Number", "Account_Number"]]
                .astype(str)
                .apply(lambda x: x.str.strip().str.upper())
                .isin(["", "NAN", "NONE", "NULL"])
                .any(axis=1)
            )
        ].reset_index(drop=True)

        if output.empty:
            print("✅ No missing turnover found. Returning preview-friendly empty row.")
            output = pd.DataFrame(
                [
                    {
                        "Customer_Number": "",
                        "Account_Number": "",
                        "TitleOfAccount": "",
                        "Purpose_of_Account": "",
                        "Status_Of_Account": "",
                        "Account_Turnover": "",
                        "KYC_Ann_TO_Corporate": "",
                        "ProductDesc": "",
                        "Sector_Description": "",
                    }
                ]
            )

        if mode == "full":
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_path = f"missing_account_turnover_in_kyc_{timestamp}.xlsx"
            output.to_excel(file_path, index=False)
            print(f"📁 Saved output to: {file_path}")

        return output

    except Exception as e:
        print(f"❌ Exception occurred: {e}")
        return pd.DataFrame(
            columns=[
                "Customer_Number",
                "Account_Number",
                "TitleOfAccount",
                "Purpose_of_Account",
                "Status_Of_Account",
                "Account_Turnover",
                "KYC_Ann_TO_Corporate",
                "ProductDesc",
                "Sector_Description",
            ]
        )
