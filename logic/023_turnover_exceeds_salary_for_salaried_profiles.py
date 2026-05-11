import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="023_turnover_exceeds_salary_for_salaried_profiles",
    description="Flags salaried customers where account turnover is disproportionately higher than declared salary/income.",
    category="Compliance & Screening",
)
def logic_023_turnover_exceeds_salary_for_salaried_profiles(
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

        required_columns = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "ACCOUNT_TURNOVER",
            "ACCOUNT_TURNOVER_IN_NUMBERS",
            "SALARY_OTHER_INCOME",
            "CUSTOMER_OCCUPATION",
            "CUSTSECTORCODE",
        ]
        missing = [col for col in required_columns if col not in df_main.columns]
        if missing:
            raise ValueError(f"Missing required columns: {missing}")

        # 🔧 Clean numeric fields
        df_main["ACCOUNT_TURNOVER_CLEAN"] = (
            df_main["ACCOUNT_TURNOVER_IN_NUMBERS"]
            .astype(str)
            .str.replace(r"[^\d.]", "", regex=True)
        )
        df_main["SALARY_INCOME_CLEAN"] = (
            df_main["SALARY_OTHER_INCOME"]
            .astype(str)
            .str.replace(r"[^\d.]", "", regex=True)
        )

        df_main["ACCOUNT_TURNOVER_NUM"] = pd.to_numeric(
            df_main["ACCOUNT_TURNOVER_CLEAN"], errors="coerce"
        )
        df_main["SALARY_INCOME_NUM"] = pd.to_numeric(
            df_main["SALARY_INCOME_CLEAN"], errors="coerce"
        )

        # ✅ Filter profiles with valid salary
        df_main = df_main[
            df_main["SALARY_INCOME_NUM"].notna() & (df_main["SALARY_INCOME_NUM"] > 0)
        ]

        # ✅ Allowed occupations: Salaried, House Wife, Student, Retired
        allowed_occupations = {"salaried", "house wife", "student", "retired"}
        occupation_clean = (
            df_main["CUSTOMER_OCCUPATION"].astype(str).str.strip().str.lower()
        )
        occupation_mask = occupation_clean.isin(allowed_occupations)

        # ✅ Turnover > 2x salary for selected occupations
        df_filtered = df_main[
            occupation_mask
            & (df_main["ACCOUNT_TURNOVER_NUM"] > df_main["SALARY_INCOME_NUM"] * 2)
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

        # 🔁 Search fallback files for ProductDesc and Sector_Description using Account_Num
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
                "PRODUCTDESC" in df_other.columns
                or "SECTOR_DESCRIPTION" in df_other.columns
            ):
                fallback_df = df_other.copy()
                break

        # 🔁 Merge fallback if found
        if fallback_df is not None and not df_filtered.empty:
            df_filtered["ACCOUNT_NUMBER"] = df_filtered["ACCOUNT_NUMBER"].astype(str)
            fallback_df["ACCOUNT_NUM"] = fallback_df["ACCOUNT_NUM"].astype(str)

            df_filtered = df_filtered.merge(
                fallback_df[
                    ["ACCOUNT_NUM"]
                    + [
                        col
                        for col in ["PRODUCTDESC", "SECTOR_DESCRIPTION"]
                        if col in fallback_df.columns
                    ]
                ],
                left_on="ACCOUNT_NUMBER",
                right_on="ACCOUNT_NUM",
                how="left",
                suffixes=("", "_FALLBACK"),
            )

            if "PRODUCTDESC_FALLBACK" in df_filtered.columns:
                df_filtered["PRODUCTDESC"] = df_filtered["PRODUCTDESC"].combine_first(
                    df_filtered["PRODUCTDESC_FALLBACK"]
                )
            if "SECTOR_DESCRIPTION_FALLBACK" in df_filtered.columns:
                df_filtered["SECTOR_DESCRIPTION"] = df_filtered[
                    "SECTOR_DESCRIPTION"
                ].combine_first(df_filtered["SECTOR_DESCRIPTION_FALLBACK"])

        # 📤 Prepare output
        output = df_filtered[
            [
                "CUSTOMER_NUMBER",
                "ACCOUNT_NUMBER",
                "TITLEOFACCOUNT",
                "ACCOUNT_TURNOVER_IN_NUMBERS",
                "SALARY_OTHER_INCOME",
                "PRODUCTDESC",
                "SECTOR_DESCRIPTION",
            ]
        ].copy()

        output.columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "Account_Turnover_in_Numbers",
            "Salary_Other_Income",
            "ProductDesc",
            "SECTOR_DESCRIPTION",
        ]

        output = output.reset_index(drop=True)

        if output.empty:
            print(
                "✅ No turnover vs salary contradictions found. Excel file not created."
            )
            output = pd.DataFrame(
                [
                    {
                        "Customer_Number": "",
                        "Account_Number": "",
                        "TitleOfAccount": "",
                        "Account_Turnover_in_Numbers": "",
                        "Salary_Other_Income": "",
                        "ProductDesc": "",
                        "SECTOR_DESCRIPTION": "",
                    }
                ]
            )

        if mode == "full":
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_path = f"turnover_exceeds_salary_salaried_profiles_{timestamp}.xlsx"
            output.to_excel(file_path, index=False)
            print(f"📁 Saved output to: {file_path}")

        return output

    except Exception as e:
        print(f"❌ Exception occurred: {e}")
        return pd.DataFrame(
            [
                {
                    "Customer_Number": "",
                    "Account_Number": "",
                    "TitleOfAccount": "",
                    "Account_Turnover_in_Numbers": "",
                    "Salary_Other_Income": "",
                    "ProductDesc": "",
                    "SECTOR_DESCRIPTION": "",
                }
            ]
        )
