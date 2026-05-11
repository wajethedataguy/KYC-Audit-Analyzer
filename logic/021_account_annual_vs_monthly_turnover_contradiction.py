import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="021_account_annual_vs_monthly_turnover_contradiction",
    description="Flags customers where annual turnover label is equal to or lower than monthly turnover label, or numeric contradiction when monthly numeric exceeds annual band.",
    category="Compliance & Screening",
)
def logic_021_account_annual_vs_monthly_turnover_contradiction(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    try:
        # Load merged KYC file
        merged_key = next(
            (k for k in dataframes if "merged_file.xlsx" in k.lower()), None
        )
        if not merged_key:
            raise ValueError("Merged file not found.")
        df_ind = dataframes[merged_key].copy()
        df_ind.columns = (
            df_ind.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )

        # Detect corporate file by columns
        df_corp = None
        for key, df_candidate in dataframes.items():
            cols = (
                df_candidate.columns.str.strip()
                .str.upper()
                .str.replace(" ", "_")
                .str.replace("-", "_")
            )
            if {
                "KYC_ANN_TO_CORPORATE",
                "MON_TOVER_CRG",
                "CUST_SECTOR_CODE",
                "PRODUCTDESC",
                "SECTOR_DESCRIPTION",
            }.issubset(set(cols)):
                df_corp = df_candidate.copy()
                df_corp.columns = cols
                break
        if df_corp is None:
            raise ValueError("Corporate file with required columns not found.")

        label_order = {
            "below 1m": 1,
            "below 10m": 1,
            "1m to 5m": 2,
            "5m to 10m": 3,
            "10m to 50m": 4,
            "10m to 50m+": 4,
            "10m to 50m plus": 4,
            "above 10m": 5,
            "above 50m": 6,
        }

        # Map annual label to numeric max
        def label_to_max(val):
            val = str(val).lower().strip()
            if val == "1m to 5m":
                return 5_000_000
            if val == "5m to 10m":
                return 10_000_000
            if val in ["10m to 50m", "10m to 50m+", "10m to 50m plus"]:
                return 50_000_000
            return None

        low_labels = {"below 1m", "below 10m"}

        def normalize_numeric(val):
            try:
                val = str(val).replace(",", "").replace("PKR", "").strip()
                return float(val)
            except:
                return None

        # ✅ Individual contradiction logic
        df_ind = df_ind[df_ind["CUSTSECTORCODE"] <= 1100].copy()
        df_ind["ANNUAL_LABEL"] = (
            df_ind["ACCOUNT_TURNOVER"].astype(str).str.lower().str.strip()
        )
        df_ind["MONTHLY_LABEL"] = (
            df_ind["ACCOUNT_MONTHLY_TURNOVER"].astype(str).str.lower().str.strip()
        )
        df_ind["ANNUAL_RANK"] = df_ind["ANNUAL_LABEL"].map(label_order)
        df_ind["MONTHLY_RANK"] = df_ind["MONTHLY_LABEL"].map(label_order)
        df_ind["ANNUAL_NUM"] = df_ind["ACCOUNT_TURNOVER_IN_NUMBERS"].apply(
            normalize_numeric
        )
        df_ind["MONTHLY_NUM"] = df_ind["EXPECTED_MONTHLY_TURNOVER_GT"].apply(
            normalize_numeric
        )
        df_ind["ANNUAL_MAX"] = df_ind["ANNUAL_LABEL"].apply(label_to_max)

        # Drop only when BOTH annual and monthly are low-band
        df_ind = df_ind[
            ~(
                df_ind["ANNUAL_LABEL"].isin(low_labels)
                & df_ind["MONTHLY_LABEL"].isin(low_labels)
            )
        ].copy()

        # Narrative contradictions (label vs label)
        narrative_ind = df_ind[
            df_ind["ANNUAL_RANK"].notnull()
            & df_ind["MONTHLY_RANK"].notnull()
            & (df_ind["ANNUAL_RANK"] <= df_ind["MONTHLY_RANK"])
            & ~(
                df_ind["ANNUAL_LABEL"].isin(["above 10m", "above 50m"])
                | df_ind["MONTHLY_LABEL"].isin(["above 10m", "above 50m"])
            )
        ].copy()

        # Numeric contradictions (annual above 10M vs monthly numeric)
        numeric_ind = df_ind[
            df_ind["ANNUAL_LABEL"].isin(["above 10m", "above 50m"])
            & df_ind["ANNUAL_NUM"].notnull()
            & df_ind["MONTHLY_NUM"].notnull()
            & (df_ind["ANNUAL_NUM"] <= df_ind["MONTHLY_NUM"])
        ].copy()

        # NEW: Label vs numeric contradictions
        label_vs_num = df_ind[
            df_ind["ANNUAL_MAX"].notnull()
            & df_ind["MONTHLY_NUM"].notnull()
            & (df_ind["MONTHLY_NUM"] > df_ind["ANNUAL_MAX"])
        ].copy()

        df_ind_combined = pd.concat(
            [narrative_ind, numeric_ind, label_vs_num], ignore_index=True
        )

        # ✅ Corporate contradiction logic (unchanged)
        df_corp["CUST_SECTOR_CODE"] = pd.to_numeric(
            df_corp["CUST_SECTOR_CODE"], errors="coerce"
        )
        df_corp = df_corp[df_corp["CUST_SECTOR_CODE"] > 1100].copy()

        df_corp["KYC_ANN_TO_CORPORATE"] = (
            df_corp["KYC_ANN_TO_CORPORATE"].astype(str).str.lower().str.strip()
        )
        df_corp["MON_TOVER_CRG"] = (
            df_corp["MON_TOVER_CRG"].astype(str).str.lower().str.strip()
        )
        df_corp["ANNUAL_RANK"] = df_corp["KYC_ANN_TO_CORPORATE"].map(label_order)
        df_corp["MONTHLY_RANK"] = df_corp["MON_TOVER_CRG"].map(label_order)
        df_corp["ANNUAL_NUM"] = df_corp["KYC_ANN_TO_CORPORATE"].apply(normalize_numeric)
        df_corp["MONTHLY_NUM"] = df_corp["MON_TOVER_CRG"].apply(normalize_numeric)
        df_corp["ANNUAL_MAX"] = df_corp["KYC_ANN_TO_CORPORATE"].apply(label_to_max)

        df_corp = df_corp[
            ~(
                df_corp["KYC_ANN_TO_CORPORATE"].isin(low_labels)
                & df_corp["MON_TOVER_CRG"].isin(low_labels)
            )
        ].copy()

        narrative_corp = df_corp[
            df_corp["ANNUAL_RANK"].notnull()
            & df_corp["MONTHLY_RANK"].notnull()
            & (df_corp["ANNUAL_RANK"] <= df_corp["MONTHLY_RANK"])
            & ~(
                df_corp["KYC_ANN_TO_CORPORATE"].isin(["above 10m", "above 50m"])
                | df_corp["MON_TOVER_CRG"].isin(["above 10m", "above 50m"])
            )
        ].copy()

        numeric_corp = df_corp[
            df_corp["KYC_ANN_TO_CORPORATE"].isin(["above 10m", "above 50m"])
            & df_corp["ANNUAL_NUM"].notnull()
            & df_corp["MONTHLY_NUM"].notnull()
            & (df_corp["ANNUAL_NUM"] <= df_corp["MONTHLY_NUM"])
        ].copy()

        label_vs_num_corp = df_corp[
            df_corp["ANNUAL_MAX"].notnull()
            & df_corp["MONTHLY_NUM"].notnull()
            & (df_corp["MONTHLY_NUM"] > df_corp["ANNUAL_MAX"])
        ].copy()

        df_corp_combined = pd.concat(
            [narrative_corp, numeric_corp, label_vs_num_corp], ignore_index=True
        )

        # ✅ Standardize output columns
        df_ind_output = df_ind_combined[
            [
                "CUSTOMER_NUMBER",
                "ACCOUNT_NUMBER",
                "TITLEOFACCOUNT",
                "PURPOSE_OF_ACCOUNT",
                "ACCOUNT_TURNOVER",
                "ACCOUNT_MONTHLY_TURNOVER",
                "PRODUCTDESC",
                "SECTOR_DESCRIPTION",
            ]
        ].copy()

        df_corp_output = df_corp_combined[
            [
                "CUSTOMER_NUM",
                "ACCOUNT_NUM",
                "TITLE_OF_ACCOUNT",
                "PURPOSE",
                "KYC_ANN_TO_CORPORATE",
                "MON_TOVER_CRG",
                "PRODUCTDESC",
                "SECTOR_DESCRIPTION",
            ]
        ].copy()

        # Rename columns consistently
        df_ind_output.columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "Purpose_of_Account",
            "Account_Turnover",
            "Account_Monthly_Turnover",
            "ProductDesc",
            "SECTOR_DESCRIPTION",
        ]

        df_corp_output.columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "Purpose_of_Account",
            "Account_Turnover",
            "Account_Monthly_Turnover",
            "ProductDesc",
            "SECTOR_DESCRIPTION",
        ]

        # Combine individual and corporate outputs
        output = pd.concat([df_ind_output, df_corp_output], ignore_index=True)

        if output.empty:
            print("✅ No contradictions found. Showing preview-friendly empty row.")
            output = pd.DataFrame(
                [
                    {
                        "Customer_Number": "",
                        "Account_Number": "",
                        "TitleOfAccount": "",
                        "Purpose_of_Account": "",
                        "Account_Turnover": "",
                        "Account_Monthly_Turnover": "",
                        "ProductDesc": "",
                        "SECTOR_DESCRIPTION": "",
                    }
                ]
            )

        # Controlled export
        if mode == "full":
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_path = f"annual_vs_monthly_turnover_contradiction_{timestamp}.xlsx"
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
                    "Purpose_of_Account": "",
                    "Account_Turnover": "",
                    "Account_Monthly_Turnover": "",
                    "ProductDesc": "",
                    "SECTOR_DESCRIPTION": "",
                }
            ]
        )
