import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="043_business_name_for_non_business_customer",
    description="Flags non-business customers whose NameOfBusiness is populated with a non-nil, non-generic value.",
    category="Compliance & Screening",
)
def logic_043_business_name_for_non_business_customer(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    required_cols = [
        "ACCOUNT_NUMBER",
        "TITLEOFACCOUNT",
        "CUSTOMER_NUMBER",
        "CUSTOMERFULLNAME",
        "CUSTSECTORCODE",
        "CUSTOMER_OCCUPATION",
        "NAMEOFBUSINESS",
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

        # 🔧 Normalize NIL values (N.A, N/A, N-A → NA)
        def normalize_nill(value: str) -> str:
            val = str(value).upper().strip()
            val = val.replace(".", "").replace("-", "").replace("/", "")
            return val

        nill_set = {normalize_nill(val) for val in raw_nill_values}
        nill_set.update({"NA"})  # safeguard

        # 🔧 Normalize fields
        df["CUSTOMER_OCCUPATION_CLEAN"] = (
            df["CUSTOMER_OCCUPATION"].fillna("").astype(str).str.lower().str.strip()
        )
        df["NAMEOFBUSINESS_CLEAN"] = (
            df["NAMEOFBUSINESS"].fillna("").astype(str).str.upper().str.strip()
        )
        df["NAMEOFBUSINESS_CLEAN"] = df["NAMEOFBUSINESS_CLEAN"].apply(normalize_nill)

        # Direct NIL membership flag
        df["IS_NAME_IN_NILL"] = df["NAMEOFBUSINESS_CLEAN"].isin(nill_set)

        excluded_occupations = {"business", "landlord", "others"}

        def is_valid_name(name_clean: str, is_in_nill: bool) -> bool:
            if is_in_nill:
                return False
            name_clean = str(name_clean).strip().upper()
            if name_clean == "":
                return False
            if name_clean.isdigit():
                return False
            special_chars = "!@#$%^&*()_+=-[]{}|\\:;\"'<>,.?/~`"
            if name_clean and all(ch in special_chars for ch in name_clean):
                return False
            return True

        df["IS_NAME_VALID"] = df.apply(
            lambda row: is_valid_name(
                row["NAMEOFBUSINESS_CLEAN"], row["IS_NAME_IN_NILL"]
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
            & df["IS_NAME_VALID"]
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
                "NAMEOFBUSINESS",
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
            "NameOfBusiness",
            "ProductDesc",
            "Sector_Description",
        ]
        output = output.reset_index(drop=True)

        # 🧯 Fallback row if no matches
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
                        "NameOfBusiness": pd.NA,
                        "ProductDesc": pd.NA,
                        "Sector_Description": pd.NA,
                    }
                ]
            )

        # 📁 Optional export
        if mode == "full":
            file_path = f"business_name_for_non_business_customer_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            output.to_excel(file_path, index=False)

        return output

    except Exception:
        return pd.DataFrame(
            [
                {
                    "Customer_Number": pd.NA,
                    "Account_Number": pd.NA,
                    "TitleOfAccount": pd.NA,
                    "CustomerFullName": pd.NA,
                    "CustSectorCode": pd.NA,
                    "Customer_Occupation": pd.NA,
                    "NameOfBusiness": pd.NA,
                    "ProductDesc": pd.NA,
                    "Sector_Description": pd.NA,
                }
            ]
        )
