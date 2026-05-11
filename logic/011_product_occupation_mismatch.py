import pandas as pd
import unicodedata
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="011_product_occupation_mismatch",
    description="Flags accounts with product-occupation mismatch based on product codes and occupation keywords.",
    category="Purpose & Occupation Filter",
)
def logic_011_product_occupation_mismatch(
    dataframes: dict, mode="full"
) -> pd.DataFrame:
    expected_columns = {
        "Account_Number",
        "TitleOfAccount",
        "Customer_Number",
        "Customer_Occupation",
        "ProductCode",
        "ProductDesc",
        "Sector_Description",
    }

    exclusion_terms = ["/", "\\", " \\", " AND ", " OR ", " & ", "& "]

    def _normalize(text):
        return unicodedata.normalize("NFKC", str(text).strip().upper())

    def _deduplicate_columns(df):
        seen = {}
        new_cols = []
        for col in df.columns:
            col_clean = str(col).strip()
            if col_clean in seen:
                seen[col_clean] += 1
                col_clean = f"{col_clean}_{seen[col_clean]}"
            else:
                seen[col_clean] = 0
            new_cols.append(col_clean)
        df.columns = new_cols
        return df

    def _normalize_code(code):
        try:
            return str(int(float(str(code).strip())))
        except Exception:
            return str(code).strip()

    def _is_mismatch(row):
        occ = row["Customer_Occupation"]
        code = row["ProductCode"]

        # Skip empty or 'OTH' occupations
        if not occ or occ.strip() == "" or occ.startswith("OTH"):
            return False

        # Rule: 1150 → occupation must contain BUS
        if code == "1150":
            return "BUS" not in occ

        # Rule: 1011 or 6012 → occupation must contain SAL
        if code in {"1011", "6012"}:
            return "SAL" not in occ

        return False

    def _is_clean_title(title):
        title = str(title).upper()
        return all(term not in title for term in exclusion_terms)

    try:
        collected_dfs = []
        for df in dataframes.values():
            df = df.copy()
            df.columns = (
                df.columns.str.strip()
                .str.upper()
                .str.replace(" ", "_")
                .str.replace("-", "_")
            )

            rename_map = {}
            for col in df.columns:
                col_upper = col.upper()
                if col_upper in [
                    "ACCOUNT_NUMBER",
                    "ACCOUNT_NO",
                    "ACCOUNT_NO_",
                    "ACC_NO",
                    "ACCT_NUMBER",
                ]:
                    rename_map[col] = "Account_Number"
                elif "TITLE" in col_upper:
                    rename_map[col] = "TitleOfAccount"
                elif "CUSTOMER" in col_upper and (
                    "NUMBER" in col_upper or "NUM" in col_upper
                ):
                    rename_map[col] = "Customer_Number"
                elif "OCCUPATION" in col_upper:
                    rename_map[col] = "Customer_Occupation"
                elif "PRODUCTCODE" in col_upper or col_upper == "PRODUCT":
                    rename_map[col] = "ProductCode"
                elif "PRODUCTDESC" in col_upper or "PRODDESC" in col_upper:
                    rename_map[col] = "ProductDesc"
                elif "SECTOR" in col_upper:
                    rename_map[col] = "Sector_Description"

            df.rename(columns=rename_map, inplace=True)
            available_cols = expected_columns.intersection(df.columns)
            if available_cols:
                df_subset = df[list(available_cols)].copy()
                df_subset = _deduplicate_columns(df_subset)
                collected_dfs.append(df_subset)

        filtered_dfs = [
            df.dropna(axis=1, how="all")
            for df in collected_dfs
            if not df.dropna(axis=1, how="all").empty
        ]

        if not filtered_dfs:
            raise ValueError("No files contain required columns.")

        merged_df = pd.concat(filtered_dfs, ignore_index=True)

        # Ensure all expected columns exist
        for col in expected_columns:
            if col not in merged_df.columns:
                merged_df[col] = ""

        # Normalize key fields
        merged_df["Customer_Occupation"] = (
            merged_df["Customer_Occupation"].fillna("").astype(str).apply(_normalize)
        )
        merged_df["TitleOfAccount"] = (
            merged_df["TitleOfAccount"].fillna("").astype(str).apply(_normalize)
        )
        merged_df["ProductDesc"] = (
            merged_df["ProductDesc"].fillna("").astype(str).apply(_normalize)
        )
        merged_df["ProductCode"] = (
            merged_df["ProductCode"].fillna("").apply(_normalize_code)
        )

        # Apply mismatch rules
        merged_df["IsMismatch"] = merged_df.apply(_is_mismatch, axis=1)

        df_flagged = merged_df[merged_df["IsMismatch"]].copy()
        df_flagged = df_flagged[
            df_flagged["TitleOfAccount"].apply(_is_clean_title)
        ].copy()

        # Ensure Account_Number exists before filtering
        if "Account_Number" not in df_flagged.columns:
            df_flagged["Account_Number"] = ""

        df_flagged = df_flagged[
            df_flagged["Account_Number"].notna()
            & (df_flagged["Account_Number"].astype(str).str.strip() != "")
        ].copy()

        df_flagged = df_flagged.drop_duplicates(
            subset=["Account_Number", "Customer_Occupation", "ProductCode"]
        )

        final_columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "ProductCode",
            "Customer_Occupation",
            "ProductDesc",
            "Sector_Description",
        ]

        # Ensure all final columns exist
        for col in final_columns:
            if col not in df_flagged.columns:
                df_flagged[col] = ""

        output = df_flagged[final_columns].copy()

        # Title-case occupation and title for readability
        for col in ["Customer_Occupation", "TitleOfAccount"]:
            output[col] = output[col].astype(str).str.title()

        if output.empty:
            output = pd.DataFrame([{col: pd.NA for col in final_columns}])

        output = output.astype(object).where(pd.notna(output), None)

        if mode == "full" and not output.isna().all(axis=1).iloc[0]:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_path = f"product_occupation_mismatch_{timestamp}.xlsx"
            output.to_excel(file_path, index=False)
            print(f"📁 Saved output to: {file_path}")
        elif mode == "full":
            print("✅ No mismatches found. Excel file not created.")

        return output

    except Exception as e:
        print(f"❌ Error in logic_011_product_occupation_mismatch: {e}")
        return pd.DataFrame(
            [
                {
                    col: pd.NA
                    for col in [
                        "Customer_Number",
                        "Account_Number",
                        "TitleOfAccount",
                        "ProductCode",
                        "Customer_Occupation",
                        "ProductDesc",
                        "Sector_Description",
                    ]
                }
            ]
        )
