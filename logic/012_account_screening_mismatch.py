import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="012_unsc_screening_mismatch",
    description="Flags accounts where UNSC screening was required but not properly recorded in KYC (blank or 'No').",
    category="Compliance",
)
def logic_012_unsc_screening_mismatch(dataframes: dict, mode="full") -> pd.DataFrame:
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
            "UNSC_SCREENING",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        missing = [col for col in required_columns if col not in df.columns]
        if missing:
            raise ValueError(f"Missing required columns: {missing}")

        df["SCREENING_CLEAN"] = df["UNSC_SCREENING"].astype(str).str.strip().str.lower()

        df_filtered = df[
            (df["SCREENING_CLEAN"].isin(["", "no", "nan"]))
            & df["CUSTOMER_NUMBER"].notna()
        ].copy()

        df_filtered["Mismatch_Reason"] = (
            "UNSC screening required but not properly recorded in KYC (blank or 'No')"
        )

        output = df_filtered[
            [
                "CUSTOMER_NUMBER",
                "ACCOUNT_NUMBER",
                "TITLEOFACCOUNT",
                "PURPOSE_OF_ACCOUNT",
                "UNSC_SCREENING",
                "PRODUCTDESC",
                "SECTOR_DESCRIPTION",
            ]
        ].copy()

        output.columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "Purpose_of_Account",
            "UNSC_Screening",
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

        # ✅ Ensure preview-friendly empty row if no records found
        if output.empty:
            print(
                "✅ No UNSC screening mismatches found. Returning preview-friendly empty row."
            )
            output = pd.DataFrame(
                [
                    {
                        "Customer_Number": "",
                        "Account_Number": "",
                        "TitleOfAccount": "",
                        "Purpose_of_Account": "",
                        "UNSC_Screening": "",
                        "ProductDesc": "",
                        "Sector_Description": "",
                    }
                ]
            )

        # 📁 Controlled export
        if mode == "full":
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_path = f"unsc_screening_mismatch_{timestamp}.xlsx"
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
                "UNSC_Screening",
                "ProductDesc",
                "Sector_Description",
            ]
        )
