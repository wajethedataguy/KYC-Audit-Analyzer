import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="024_actual_credit_turnover_exceeds_income_for_nonbusiness_profiles",
    description="Flags top 15 non-business customers where actual credit turnover exceeds declared salary/income.",
    category="Compliance & Screening",
)
def logic_024_actual_credit_turnover_exceeds_income_for_nonbusiness_profiles(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    try:
        # 🔍 Locate merged file
        merged_key = next(
            (k for k in dataframes if "merged_file.xlsx" in k.lower()), None
        )
        if merged_key is None:
            raise ValueError("Merged file not found in input dataframes.")
        df_main = dataframes[merged_key].copy()

        # 🔧 Normalize columns
        df_main.columns = (
            df_main.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )

        required_main = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "SALARY_OTHER_INCOME",
            "CUSTOMER_OCCUPATION",
            "CUSTSECTORCODE",
        ]
        missing_main = [col for col in required_main if col not in df_main.columns]
        if missing_main:
            raise ValueError(f"Missing required columns in merged file: {missing_main}")

        # 🔁 Search fallback for TOTAL_CREDIT_TOTAL, PRODUCTDESC, SECTOR_DESCRIPTION
        fallback_df = None
        for key, df_other in dataframes.items():
            if key == merged_key:
                continue
            df_other.columns = (
                df_other.columns.str.strip()
                .str.upper()
                .str.replace(" ", "_")
                .str.replace("-", "_")
            )
            if "ACCOUNT_NUM" in df_other.columns and (
                "TOTAL_CREDIT_TOTAL" in df_other.columns
                or "PRODUCTDESC" in df_other.columns
                or "SECTOR_DESCRIPTION" in df_other.columns
            ):
                fallback_df = df_other.copy()
                break

        # 🔁 Merge fallback if found
        if fallback_df is not None:
            df_main["ACCOUNT_NUMBER"] = df_main["ACCOUNT_NUMBER"].astype(str)
            fallback_df["ACCOUNT_NUM"] = fallback_df["ACCOUNT_NUM"].astype(str)

            df_main = df_main.merge(
                fallback_df[
                    ["ACCOUNT_NUM"]
                    + [
                        col
                        for col in [
                            "TOTAL_CREDIT_TOTAL",
                            "PRODUCTDESC",
                            "SECTOR_DESCRIPTION",
                        ]
                        if col in fallback_df.columns
                    ]
                ],
                left_on="ACCOUNT_NUMBER",
                right_on="ACCOUNT_NUM",
                how="left",
                suffixes=("", "_FALLBACK"),
            )

            for col in ["TOTAL_CREDIT_TOTAL", "PRODUCTDESC", "SECTOR_DESCRIPTION"]:
                fallback_col = f"{col}_FALLBACK"
                if fallback_col in df_main.columns:
                    df_main[col] = df_main[col].combine_first(df_main[fallback_col])

        # 🔧 Clean numeric fields
        df_main["SALARY_INCOME_CLEAN"] = (
            df_main["SALARY_OTHER_INCOME"]
            .astype(str)
            .str.replace(r"[^\d.]", "", regex=True)
        )
        df_main["CREDIT_TURNOVER_CLEAN"] = (
            df_main["TOTAL_CREDIT_TOTAL"]
            .astype(str)
            .str.replace(r"[^\d.]", "", regex=True)
        )

        df_main["SALARY_INCOME_NUM"] = pd.to_numeric(
            df_main["SALARY_INCOME_CLEAN"], errors="coerce"
        )
        df_main["CREDIT_TURNOVER_NUM"] = pd.to_numeric(
            df_main["CREDIT_TURNOVER_CLEAN"], errors="coerce"
        )

        # ✅ Normalize occupation and filter non-business profiles
        df_main["OCCUPATION_NORMALIZED"] = (
            df_main["CUSTOMER_OCCUPATION"].astype(str).str.lower().str.strip()
        )
        non_business_roles = {"salaried", "housewife", "retired", "student"}

        df_filtered = df_main[
            df_main["OCCUPATION_NORMALIZED"].isin(non_business_roles)
            & df_main["SALARY_INCOME_NUM"].notnull()
            & (df_main["SALARY_INCOME_NUM"] > 0)
            & (df_main["CREDIT_TURNOVER_NUM"] > df_main["SALARY_INCOME_NUM"])
        ].copy()

        # 🚫 Apply joint account exclusion ONLY for CustSectorCode 1000/1005
        joint_patterns = [
            "/",
            "\\",
            " AND ",
            " OR ",
            " & ",
            " S/O ",
            " S\\O ",
            " D/O ",
            " D\\O ",
            " W/O ",
            " W\\O ",
        ]
        pattern_regex = "|".join([p.replace("\\", "\\\\") for p in joint_patterns])

        mask_sector = (
            df_filtered["CUSTSECTORCODE"]
            .astype(str)
            .isin(["1000", "1000.0", "1005", "1005.0"])
        )
        mask_joint = (
            df_filtered["TITLEOFACCOUNT"]
            .astype(str)
            .str.upper()
            .str.contains(pattern_regex, na=False)
        )

        df_filtered = df_filtered[~(mask_sector & mask_joint)].copy()

        # ✅ Calculate difference
        df_filtered["DIFFERENCE"] = (
            df_filtered["CREDIT_TURNOVER_NUM"] - df_filtered["SALARY_INCOME_NUM"]
        )

        # ✅ Sort and select top 15
        df_top15 = df_filtered.sort_values(by="DIFFERENCE", ascending=False).head(15)

        # 📤 Prepare output
        output = df_top15[
            [
                "CUSTOMER_NUMBER",
                "ACCOUNT_NUMBER",
                "TITLEOFACCOUNT",
                "PRODUCTDESC",
                "SECTOR_DESCRIPTION",
                "CREDIT_TURNOVER_NUM",
                "SALARY_INCOME_NUM",
                "DIFFERENCE",
            ]
        ].copy()

        output.columns = [
            "Customer ID",
            "Account Number",
            "Account Title",
            "ProductDesc",
            "SECTOR_DESCRIPTION",
            "Account Annual Actual Credit Turnover",
            "Salary & Other Income",
            "Difference",
        ]

        output = output.reset_index(drop=True)

        if output.empty:
            print(
                "✅ No credit turnover vs income contradictions found. Returning preview-friendly empty row."
            )
            output = pd.DataFrame(
                [
                    {
                        "Customer ID": "",
                        "Account Number": "",
                        "Account Title": "",
                        "ProductDesc": "",
                        "SECTOR_DESCRIPTION": "",
                        "Account Annual Actual Credit Turnover": "",
                        "Salary & Other Income": "",
                        "Difference": "",
                    }
                ]
            )

        if mode == "full":
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_path = f"top15_credit_turnover_exceeds_income_{timestamp}.xlsx"
            output.to_excel(file_path, index=False)
            print(f"📁 Saved output to: {file_path}")

        return output

    except Exception as e:
        print(f"❌ Exception occurred: {e}")
        return pd.DataFrame(
            [
                {
                    "Customer ID": "",
                    "Account Number": "",
                    "Account Title": "",
                    "ProductDesc": "",
                    "SECTOR_DESCRIPTION": "",
                    "Account Annual Actual Credit Turnover": "",
                    "Salary & Other Income": "",
                    "Difference": "",
                }
            ]
        )
