import pandas as pd
from datetime import datetime
from KYC_Viewer.utils import register_logic


@register_logic(
    name="003_joint_name_misassignment",
    description="Flags personal accounts where customer name contains joint indicators but no joint holder is declared.",
    category="Customer Name Filter",
)
def logic_003_joint_name_misassignment(dataframes: dict, mode="full") -> pd.DataFrame:
    try:
        # Load merged KYC file
        merged_key = next(
            (k for k in dataframes if "merged_file.xlsx" in k.lower()), None
        )
        df = dataframes.get(merged_key)
        if isinstance(df, tuple):
            df = next(iter(df[0].values())) if df[0] else pd.DataFrame()
        if df is None or df.empty:
            raise ValueError("Merged file not found or empty.")

        # Normalize column names
        df.columns = (
            df.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_")
            .str.replace("-", "_")
        )

        # Validate required columns
        required_columns = [
            "COB_DATE",
            "CUSTOMER_NUMBER",
            "CUSTOMERFULLNAME",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "ACCOUNT_OPEN_DT",
            "ACCOUNT_TURNOVER",
            "ACCOUNT_TURNOVER_IN_NUMBERS",
            "CUSTSECTORCODE",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        missing = [col for col in required_columns if col not in df.columns]
        if missing:
            raise ValueError(f"Missing required columns: {missing}")

        # Normalize sector code
        df["CUSTSECTORCODE"] = pd.to_numeric(df["CUSTSECTORCODE"], errors="coerce")

        # Define joint name patterns
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
        exclude_patterns = [" S/O ", " S\\O ", " D/O ", " D\\O ", " W/O ", " W\\O "]

        # Apply filters
        pattern_filter = df["CUSTOMERFULLNAME"].apply(
            lambda name: any(p in str(name).upper() for p in joint_patterns)
        )
        exclusion_filter = df["CUSTOMERFULLNAME"].apply(
            lambda name: not any(p in str(name).upper() for p in exclude_patterns)
        )

        contradiction_mask = (
            (df["CUSTSECTORCODE"] <= 1005) & pattern_filter & exclusion_filter
        )

        df_flagged = df[contradiction_mask].copy()
        df_flagged["Mismatch_Reason"] = (
            "Customer name contains joint indicators but no joint holder declared"
        )
        df_flagged["SourceLogicName"] = "003_joint_name_misassignment"

        # Output columns
        output_cols = [
            "CUSTOMER_NUMBER",
            "ACCOUNT_NUMBER",
            "TITLEOFACCOUNT",
            "CUSTOMERFULLNAME",
            "PRODUCTDESC",
            "SECTOR_DESCRIPTION",
        ]
        output = df_flagged[output_cols].drop_duplicates().reset_index(drop=True)
        output.columns = [
            "Customer_Number",
            "Account_Number",
            "TitleOfAccount",
            "CustomerFullName",
            "ProductDesc",
            "Sector_Description",
        ]

        if output.empty:
            output = pd.DataFrame([{col: pd.NA for col in output.columns}])

        if mode == "full":
            file_path = f"joint_name_misassignment_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            output.to_excel(file_path, index=False)
        elif mode == "ui":
            output = output[
                [
                    "Customer_Number",
                    "Account_Number",
                    "TitleOfAccount",
                    "CustomerFullName",
                    "ProductDesc",
                    "Sector_Description",
                ]
            ]

        return output

    except Exception as e:
        print(f"❌ Error in logic_003_joint_name_misassignment: {e}")
        return pd.DataFrame(
            [
                {
                    col: pd.NA
                    for col in [
                        "Customer_Number",
                        "Account_Number",
                        "TitleOfAccount",
                        "CustomerFullName",
                        "ProductDesc",
                        "Sector_Description",
                    ]
                }
            ]
        )
