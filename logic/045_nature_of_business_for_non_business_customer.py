import pandas as pd
import re
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="045_nature_of_business_for_non_business_customer",
    description="Flags non-business customers whose NatureOfBusiness is populated with a non-nil, non-generic value.",
    category="Compliance & Screening",
)
def logic_045_nature_of_business_for_non_business_customer(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    required_cols = [
        "ACCOUNT_NUMBER",
        "TITLEOFACCOUNT",
        "CUSTOMER_NUMBER",
        "CUSTOMERFULLNAME",
        "CUSTSECTORCODE",
        "CUSTOMER_OCCUPATION",
        "NATUREOFBUSINESS",
        "PRODUCTDESC",
        "SECTOR_DESCRIPTION",
    ]

    try:
        # 🔍 Identify KYC file
        kyc_key = next((k for k in dataframes if "merged_file.xlsx" in k.lower()), None)
        if not kyc_key:
            raise ValueError("KYC profile file not found.")

        df = dataframes[kyc_key].copy()
        df.columns = (
            df.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )

        # ✅ Check required columns
        if missing := [col for col in required_cols if col not in df.columns]:
            raise ValueError(f"Missing required columns: {missing}")

        # 🔍 Load NILL_COMBINATIONS reference
        raw_nill_values = set()
        for df_nill in dataframes.values():
            if isinstance(df_nill, dict):
                for sheet in df_nill.values():
                    if (
                        isinstance(sheet, pd.DataFrame)
                        and "NILL_COMBINATIONS" in sheet.columns
                    ):
                        raw_nill_values.update(
                            sheet["NILL_COMBINATIONS"].dropna().astype(str).str.strip()
                        )
            elif (
                isinstance(df_nill, pd.DataFrame)
                and "NILL_COMBINATIONS" in df_nill.columns
            ):
                raw_nill_values.update(
                    df_nill["NILL_COMBINATIONS"].dropna().astype(str).str.strip()
                )

        # 🔧 Strong normalization function
        def normalize_nill(value: str) -> str:
            val = str(value).upper().strip()
            # Remove all non-alphanumeric characters
            val = re.sub(r"[^A-Z0-9]", "", val)
            return val

        # Normalize reference set
        nill_set = {normalize_nill(val) for val in raw_nill_values}
        nill_set.add("NA")  # explicit safeguard

        # 🔧 Normalize fields
        df["CUSTOMER_OCCUPATION_CLEAN"] = (
            df["CUSTOMER_OCCUPATION"].fillna("").astype(str).str.lower().str.strip()
        )
        df["NATUREOFBUSINESS_CLEAN"] = (
            df["NATUREOFBUSINESS"].fillna("").apply(normalize_nill)
        )

        # Direct NIL membership flag
        df["IS_NATURE_IN_NILL"] = df["NATUREOFBUSINESS_CLEAN"].isin(nill_set)

        # Business-type occupations to exclude
        excluded_occupations = {"business", "landlord", "others"}

        def is_valid_nature(val: str, is_in_nill: bool) -> bool:
            if is_in_nill:
                return False
            if val == "":
                return False
            if val.isdigit():
                return False
            return True

        df["IS_NATURE_VALID"] = df.apply(
            lambda row: is_valid_nature(
                row["NATUREOFBUSINESS_CLEAN"], row["IS_NATURE_IN_NILL"]
            ),
            axis=1,
        )

        # 🧠 Flagging logic
        df["IS_NON_BUSINESS_OCCUPATION"] = ~df["CUSTOMER_OCCUPATION_CLEAN"].isin(
            excluded_occupations
        )
        df["IS_OCCUPATION_PRESENT"] = df["CUSTOMER_OCCUPATION_CLEAN"] != ""

        df_flagged = df[
            df["IS_NON_BUSINESS_OCCUPATION"]
            & df["IS_OCCUPATION_PRESENT"]
            & df["IS_NATURE_VALID"]
        ].copy()

        # 📤 Prepare output (Customer_Number always first)
        output = df_flagged[
            [
                "CUSTOMER_NUMBER",
                "ACCOUNT_NUMBER",
                "TITLEOFACCOUNT",
                "CUSTOMERFULLNAME",
                "CUSTSECTORCODE",
                "CUSTOMER_OCCUPATION",
                "NATUREOFBUSINESS",
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
            "NatureOfBusiness",
            "ProductDesc",
            "Sector_Description",
        ]
        output = output.reset_index(drop=True)

        # 🧯 Fallback row if no matches
        if output.empty:
            output = pd.DataFrame([{col: pd.NA for col in output.columns}])

        # 📁 Optional export
        if mode == "full":
            file_path = f"nature_of_business_for_non_business_customer_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            output.to_excel(file_path, index=False)

        return output

    except Exception as e:
        print(f"❌ Error: {e}")
        return pd.DataFrame(
            [
                {
                    "Customer_Number": pd.NA,
                    "Account_Number": pd.NA,
                    "TitleOfAccount": pd.NA,
                    "CustomerFullName": pd.NA,
                    "CustSectorCode": pd.NA,
                    "Customer_Occupation": pd.NA,
                    "NatureOfBusiness": pd.NA,
                    "ProductDesc": pd.NA,
                    "Sector_Description": pd.NA,
                }
            ]
        )
