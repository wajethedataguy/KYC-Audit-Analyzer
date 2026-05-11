import pandas as pd
import re
from datetime import datetime
from dateutil.parser import parse
from KYC_Viewer.utils import register_logic


@register_logic(
    name="101_employment_date_for_non_salaried_flag",
    description="Flags KYC records where employment date was fed for non-salaried individuals.",
    category="CDD & EDD Review",
)
def logic_101_employment_date_for_non_salaried_flag(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    try:
        # Load KYC file
        kyc_key = next((k for k in dataframes if "merged_file.xlsx" in k.lower()), None)
        df_kyc = dataframes.get(kyc_key)
        if isinstance(df_kyc, tuple):
            df_kyc = next(iter(df_kyc[0].values())) if df_kyc[0] else pd.DataFrame()
        if df_kyc is None or df_kyc.empty:
            raise ValueError("KYC file not found or empty.")

        # Normalize column names
        df_kyc.columns = (
            df_kyc.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )

        # ── Load NILL_COMBINATIONS from any other uploaded file ───────────
        # These values are treated as "no meaningful data entered" and are
        # EXCLUDED from flagging (e.g. "N/A", "NA", "Nill.", "nil" etc.)
        raw_nill_values = set()
        for k, df_nill in dataframes.items():
            if k == kyc_key:
                continue

            # Handle tuple (same structure as KYC file loading)
            if isinstance(df_nill, tuple):
                df_nill = next(iter(df_nill[0].values())) if df_nill[0] else pd.DataFrame()

            if isinstance(df_nill, dict):
                for sheet in df_nill.values():
                    if not isinstance(sheet, pd.DataFrame):
                        continue
                    sheet.columns = (
                        sheet.columns.str.strip()
                        .str.upper()
                        .str.replace(" ", "_")
                        .str.replace("-", "_")
                    )
                    if "NILL_COMBINATIONS" in sheet.columns:
                        raw_nill_values.update(
                            sheet["NILL_COMBINATIONS"]
                            .dropna()
                            .astype(str)
                            .str.strip()
                            .str.lower()
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
                        df_nill["NILL_COMBINATIONS"]
                        .dropna()
                        .astype(str)
                        .str.strip()
                        .str.lower()
                    )

        nill_set = {str(v).strip().lower() for v in raw_nill_values}

        # ── Normalize KYC fields ──────────────────────────────────────────
        df_kyc["CUSTOMER_OCCUPATION_CLEAN"] = (
            df_kyc["CUSTOMER_OCCUPATION"].fillna("").astype(str).str.strip().str.lower()
        )
        df_kyc["EMPLOYEDSINCE_RAW"] = (
            df_kyc["EMPLOYEDSINCE"].fillna("").astype(str).str.strip().str.lower()
        )

        # ── Allowed occupations (these are never flagged) ─────────────────
        allowed_occupations = {"salaried", "others", ""}

        # ── Hard-skip values (truly empty — not meaningful data) ──────────
        # These are separate from nill_set: they are pandas read artefacts
        # or genuinely empty cells which carry no information at all.
        hard_skip = {"", "nan", "none", "null"}

        # ── Contradiction mask ────────────────────────────────────────────
        # Flag if:
        #   1. Occupation is NOT in the allowed list (non-salaried)
        #   2. EmployedSince has a value (not blank / not a pandas artefact)
        #   3. EmployedSince is NOT in the NILL list (e.g. not "N/A", "Nill." etc.)
        #
        # Result:
        #   ✅ Flagged  : "2025-04-30", "2 Years", "Jan 2020"  (real data)
        #   ❌ Excluded : "N/A", "NA", "Nill.", "nil", "n/a"   (nill list)
        #   ❌ Excluded : ""  (blank)

        contradiction_mask = (
            ~df_kyc["CUSTOMER_OCCUPATION_CLEAN"].isin(allowed_occupations)
            & ~df_kyc["EMPLOYEDSINCE_RAW"].isin(hard_skip)
            & ~df_kyc["EMPLOYEDSINCE_RAW"].isin(nill_set)
        )

        # ── Build output ──────────────────────────────────────────────────
        output_cols = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "CUSTOMERFULLNAME",
            "CUSTSECTORCODE",
            "CUSTOMER_OCCUPATION",
            "EMPLOYEDSINCE",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        output = (
            df_kyc[contradiction_mask][output_cols]
            .drop_duplicates()
            .reset_index(drop=True)
        )

        if output.empty:
            output = pd.DataFrame([{col: pd.NA for col in output_cols}])

        # Rename to friendly names
        output = output.rename(
            columns={
                "CUSTOMER_NUMBER":    "Customer_Number",
                "ACCOUNT_NUMBER":     "Account_Number",
                "TITLEOFACCOUNT":     "TitleOfAccount",
                "CUSTOMERFULLNAME":   "CustomerFullName",
                "CUSTSECTORCODE":     "CustSectorCode",
                "CUSTOMER_OCCUPATION":"Customer_Occupation",
                "EMPLOYEDSINCE":      "EmployedSince",
                "PRODUCTDESC":        "ProductDesc",
                "SECTOR_DESCRIPTION": "Sector_Description",
            }
        )

        # Export / UI mode
        if mode == "full":
            file_path = f"employment_date_for_non_salaried_flag_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            output.to_excel(file_path, index=False)
        elif mode == "ui":
            output = output[[
                "Customer_Number",
                "Account_Number",
                "TitleOfAccount",
                "CustomerFullName",
                "CustSectorCode",
                "Customer_Occupation",
                "EmployedSince",
                "ProductDesc",
                "Sector_Description",
            ]]

        return output

    except Exception as e:
        print(f"❌ Error in logic_101_employment_date_for_non_salaried_flag: {e}")
        return pd.DataFrame([{
            "Customer_Number":    pd.NA,
            "Account_Number":     pd.NA,
            "TitleOfAccount":     pd.NA,
            "CustomerFullName":   pd.NA,
            "CustSectorCode":     pd.NA,
            "Customer_Occupation":pd.NA,
            "EmployedSince":      pd.NA,
            "ProductDesc":        pd.NA,
            "Sector_Description": pd.NA,
        }])