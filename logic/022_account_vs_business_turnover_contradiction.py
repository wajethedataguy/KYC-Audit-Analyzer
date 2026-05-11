import re
import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="022_account_vs_business_turnover_contradiction",
    description="Flags business customers where declared account turnover exceeds declared business turnover.",
    category="Compliance & Screening",
)
def logic_022_account_vs_business_turnover_contradiction(
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

        # Ensure required columns exist
        required_main = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "ACCOUNT_TURNOVER_IN_NUMBERS",
            "BUSINESSTURNOVER",
            "CUSTOMER_OCCUPATION",
            "CUSTSECTORCODE",
        ]
        if "KYC_ANN_TO_CORPORATE" not in df_main.columns:
            df_main["KYC_ANN_TO_CORPORATE"] = ""

        missing_main = [col for col in required_main if col not in df_main.columns]
        if missing_main:
            raise ValueError(f"Missing required columns in merged file: {missing_main}")

        # 🔧 Clean numeric fields for individuals
        df_main["ACCOUNT_TURNOVER_NUM"] = pd.to_numeric(
            df_main["ACCOUNT_TURNOVER_IN_NUMBERS"]
            .astype(str)
            .str.replace(r"[^\d.]", "", regex=True),
            errors="coerce",
        )
        df_main["BUSINESSTURNOVER_NUM"] = pd.to_numeric(
            df_main["BUSINESSTURNOVER"]
            .astype(str)
            .str.replace(r"[^\d.]", "", regex=True),
            errors="coerce",
        )

        # 🔧 Parse corporate bands dynamically
        def parse_corporate_band(val: str):
            if not val:
                return None
            s = str(val).strip().lower().replace(",", "")
            s = " ".join(s.split())  # normalize spaces

            if s.startswith("below"):
                return 0
            m_above = re.search(r"above\s+(\d+)\s*m", s)
            if m_above:
                return float(m_above.group(1)) * 1_000_000
            m_range = re.search(r"(\d+)\s*m\s*(?:to|-)\s*(\d+)\s*m", s)
            if m_range:
                return float(m_range.group(1)) * 1_000_000
            m_single = re.search(r"(\d+)\s*m", s)
            if m_single:
                return float(m_single.group(1)) * 1_000_000
            return None

        df_main["CORPORATE_TURNOVER_NUM"] = df_main["KYC_ANN_TO_CORPORATE"].apply(
            parse_corporate_band
        )

        # ❌ Exclude corporates with "Above 50M"
        df_main = df_main[
            ~df_main["KYC_ANN_TO_CORPORATE"]
            .astype(str)
            .str.strip()
            .str.lower()
            .str.startswith("above 50m")
        ]

        # ✅ Apply contradiction logic separately for individuals and corporates
        contradiction_individual = df_main[
            df_main["ACCOUNT_TURNOVER_NUM"].notnull()
            & df_main["BUSINESSTURNOVER_NUM"].notnull()
            & (df_main["ACCOUNT_TURNOVER_NUM"] > df_main["BUSINESSTURNOVER_NUM"])
        ]

        contradiction_corporate = df_main[
            df_main["CORPORATE_TURNOVER_NUM"].notnull()
            & df_main["BUSINESSTURNOVER_NUM"].notnull()
            & (df_main["CORPORATE_TURNOVER_NUM"] > df_main["BUSINESSTURNOVER_NUM"])
        ]

        df_filtered = pd.concat(
            [contradiction_individual, contradiction_corporate], ignore_index=True
        )

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

        # 📤 Prepare output (drop BusinessTurnover_Num)
        output = df_filtered[
            [
                "CUSTOMER_NUMBER",
                "ACCOUNT_NUMBER",
                "TITLEOFACCOUNT",
                "ACCOUNT_TURNOVER_IN_NUMBERS",
                "KYC_ANN_TO_CORPORATE",
                "CORPORATE_TURNOVER_NUM",
                "BUSINESSTURNOVER",
                "PRODUCTDESC",
                "SECTOR_DESCRIPTION",
            ]
        ].copy()

        output.columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "Account_Turnover_in_Numbers",
            "KYC_Ann_TO_Corporate",
            "Corporate_Turnover_Num",
            "BusinessTurnover",
            "ProductDesc",
            "SECTOR_DESCRIPTION",
        ]

        output = output.reset_index(drop=True)

        if output.empty:
            print(
                "✅ No account vs business turnover contradictions found. Returning preview-friendly empty row."
            )
            output = pd.DataFrame(
                [
                    {
                        "Customer_Number": "",
                        "Account_Number": "",
                        "TitleOfAccount": "",
                        "Account_Turnover_in_Numbers": "",
                        "KYC_Ann_TO_Corporate": "",
                        "Corporate_Turnover_Num": "",
                        "BusinessTurnover": "",
                        "ProductDesc": "",
                        "SECTOR_DESCRIPTION": "",
                    }
                ]
            )

        if mode == "full":
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_path = f"account_vs_business_turnover_contradiction_{timestamp}.xlsx"
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
                    "KYC_Ann_TO_Corporate": "",
                    "Corporate_Turnover_Num": "",
                    "BusinessTurnover": "",
                    "ProductDesc": "",
                    "SECTOR_DESCRIPTION": "",
                }
            ]
        )
