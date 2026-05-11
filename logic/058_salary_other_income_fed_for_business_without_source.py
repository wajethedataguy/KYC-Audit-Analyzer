import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


def normalize_text(value: str) -> str:
    return str(value).strip().upper().replace(".", "").replace("/", "").replace(" ", "")


@register_logic(
    name="058_salary_other_income_fed_for_business_without_source",
    description="Flags business customers where Salary_Other_Income is fed without source justification.",
    category="Compliance & Screening",
)
def logic_058_salary_other_income_fed_for_business_without_source(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    try:
        # 🔍 Locate merged KYC file
        kyc_key = next((k for k in dataframes if "merged_file.xlsx" in k.lower()), None)
        if not kyc_key:
            raise ValueError("KYC profile file not found.")
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
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "CUSTOMER_NUMBER",
            "CUSTOMERFULLNAME",
            "CUSTSECTORCODE",
            "CUSTOMER_OCCUPATION",
            "SALARY_OTHER_INCOME",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        missing = [col for col in required_cols if col not in df_main.columns]
        if missing:
            raise ValueError(f"Merged file missing required columns: {missing}")

        df = df_main[required_cols].copy()

        # 🔁 Force occupation to Business if CustSectorCode >= 1100
        df["CUSTSECTORCODE_NUM"] = pd.to_numeric(df["CUSTSECTORCODE"], errors="coerce")
        df.loc[df["CUSTSECTORCODE_NUM"] >= 1100, "CUSTOMER_OCCUPATION"] = "Business"

        # 🔍 Load Nill_Combinations reference
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

        # 🔍 Normalize fields
        df["CUSTOMER_OCCUPATION_CLEAN"] = (
            df["CUSTOMER_OCCUPATION"].astype(str).str.strip()
        )
        df["CUSTOMER_OCCUPATION_NORMALIZED"] = df["CUSTOMER_OCCUPATION_CLEAN"].apply(
            normalize_text
        )
        df["SALARY_OTHER_INCOME_CLEAN"] = (
            df["SALARY_OTHER_INCOME"].astype(str).str.strip()
        )
        df["SALARY_OTHER_INCOME_NORMALIZED"] = df["SALARY_OTHER_INCOME_CLEAN"].apply(
            normalize_text
        )

        # 🔧 Attempt numeric conversion
        def parse_income(val):
            try:
                return float(val)
            except:
                return None

        df["SALARY_OTHER_INCOME_NUMERIC"] = df["SALARY_OTHER_INCOME_CLEAN"].apply(
            parse_income
        )

        # 🧠 Apply contradiction logic
        df_filtered = df[
            (~df["SALARY_OTHER_INCOME_NORMALIZED"].isin(nill_set_normalized))
            & (df["CUSTOMER_OCCUPATION_NORMALIZED"] == "BUSINESS")
            & (df["SALARY_OTHER_INCOME_NUMERIC"] > 0)
        ].copy()

        # 📤 Prepare output
        output = df_filtered[
            [
                "CUSTOMER_NUMBER",
                "ACCOUNT_NUMBER",
                "TITLEOFACCOUNT",
                "CUSTOMERFULLNAME",
                "PRODUCTDESC",
                "SECTOR_DESCRIPTION",
                "SALARY_OTHER_INCOME",
            ]
        ].copy()

        output.columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "CustomerFullName",
            "ProductDesc",
            "SECTOR_DESCRIPTION",
            "SALARY_OTHER_INCOME",
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
                        "CustomerFullName": pd.NA,
                        "ProductDesc": pd.NA,
                        "SECTOR_DESCRIPTION": pd.NA,
                        "SALARY_OTHER_INCOME": pd.NA,
                    }
                ]
            )

        # 📁 Optional export
        if mode == "full":
            file_path = f"salary_other_income_fed_for_business_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
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
                    "CustomerFullName": pd.NA,
                    "ProductDesc": pd.NA,
                    "SECTOR_DESCRIPTION": pd.NA,
                    "SALARY_OTHER_INCOME": pd.NA,
                }
            ]
        )
