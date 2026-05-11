import pandas as pd
import re
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="018_monthly_turnover_contradiction",
    description="Flags customers where primary & secondary account expected monthly credit turnovers are contradictory.",
    category="Compliance & Screening",
)
def logic_018_monthly_turnover_contradiction(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    try:
        # 🔍 Locate merged file
        merged_key = next(
            (k for k in dataframes if "merged_file.xlsx" in k.lower()), None
        )
        if merged_key is None:
            raise ValueError("Merged file not found in input dataframes.")

        df = dataframes[merged_key].copy()

        # 🔧 Normalize column names
        df.columns = (
            df.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )

        # 🔧 ID cleaner
        def clean_id(val):
            try:
                num = float(val)
                return str(int(num)) if num.is_integer() else str(val).strip()
            except:
                return str(val).strip()

        # 🔧 Robust numeric normalizer
        def normalize_numeric(val):
            if pd.isna(val):
                return 0.0
            s = re.sub(r"[^0-9.]", "", str(val))
            try:
                return float(s) if s else 0.0
            except:
                return 0.0

        # 🔧 Blank checker
        def is_blank(val):
            return (
                val is None
                or (isinstance(val, float) and pd.isna(val))
                or (isinstance(val, str) and val.strip() == "")
            )

        # Apply cleaning
        df["CUSTOMER_NUMBER"] = df["CUSTOMER_NUMBER"].fillna("").apply(clean_id)
        df["ACCOUNT_NUMBER"] = df["ACCOUNT_NUMBER"].fillna("").apply(clean_id)

        # ---------------------------------------------------------
        #  CORE FILTER: handle retail + corporate
        # ---------------------------------------------------------
        def has_contradiction(row):
            contradiction = False

            # --- Retail check ---
            val_retail_raw = row.get("ACCOUNT_MONTHLY_TURNOVERGT1050M", None)
            retail_label_raw = row.get("ACCOUNT_MONTHLY_TURNOVER", None)

            if not (is_blank(val_retail_raw) or is_blank(retail_label_raw)):
                val_retail = normalize_numeric(val_retail_raw)
                retail_label = str(retail_label_raw).lower().strip()
                if (
                    retail_label not in ["above 10m", "above 50m", "below 10m"]
                    and val_retail > 0
                ):
                    contradiction = True

            # --- Corporate check ---
#            corp_label_raw = row.get("ACCOUNT_MONTHLY_TURNOVER", None)
#            corp_val_raw = row.get("ACCOUNT_MONTHLY_TURNOVERGT1050M", None)

#            if not (is_blank(corp_label_raw) or is_blank(corp_val_raw)):
#                corp_label = str(corp_label_raw).lower().strip()
#                corp_val = normalize_numeric(corp_val_raw)
#                if corp_label == "10m to 50m" and corp_val > 0:
#                    contradiction = True

            return contradiction

        df["IsContradiction"] = df.apply(has_contradiction, axis=1)
        df_filtered = df[df["IsContradiction"]].copy()

        # Metadata
        df_filtered["SourceLogicName"] = "018_monthly_turnover_contradiction"
        df_filtered["Logic_Version"] = "018_v11.0"

        # 📤 Prepare output
        output_cols = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "PURPOSE_OF_ACCOUNT",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
            "ACCOUNT_MONTHLY_TURNOVER",
            "Account_Monthly_TurnoverGT1050M",
            "MON_TOVER_CRG",
            "MON_TOVER_CORP",
        ]

        for col in output_cols:
            if col not in df_filtered.columns:
                df_filtered[col] = ""

        output = df_filtered[output_cols].copy()

        # Rename columns for consistency
        rename_map = {
            "CUSTOMER_NUMBER": "Customer_Number",
            "ACCOUNT_NUMBER": "Account_Number",
            "TITLEOFACCOUNT": "TitleOfAccount",
            "PURPOSE_OF_ACCOUNT": "Purpose_of_Account",
            "PRODUCTDESC": "ProductDesc",
            "SECTOR_DESCRIPTION": "Sector_Description",
            "ACCOUNT_MONTHLY_TURNOVER": "Account_Monthly_Turnover",
            "MON_TOVER_CRG": "MON_TOVER_CRG",
            "MON_TOVER_CORP": "MON_TOVER_CORP",
            "Account_Monthly_TurnoverGT1050M": "Expected_Monthly_Turnover_GT",
        }
        output.rename(columns=rename_map, inplace=True)

        output = output.reset_index(drop=True)

        # Excel export
        if mode == "full" and not output.empty:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_path = f"monthly_turnover_contradiction_{timestamp}.xlsx"
            output.to_excel(file_path, index=False)
            print(f"📁 Saved output to: {file_path}")
        elif mode == "full":
            print("✅ No contradictions found. Creating preview-friendly row.")
            output = pd.DataFrame([{col: "" for col in rename_map.values()}])

        return output

    except Exception as e:
        print(f"❌ Exception occurred: {e}")
        return pd.DataFrame(
            [
                {
                    col: ""
                    for col in [
                        "Customer_Number",
                        "Account_Number",
                        "TitleOfAccount",
                        "Purpose_of_Account",
                        "Account_Monthly_Turnover",
                        "Expected_Monthly_Turnover_GT",
                        "MON_TOVER_CRG",
                        "MON_TOVER_CORP",
                        "ProductDesc",
                        "Sector_Description",
                    ]
                }
            ]
        )
