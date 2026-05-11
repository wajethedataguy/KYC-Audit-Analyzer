from KYC_Viewer.utils import logic_registry
from KYC_Viewer.exporter import save_excel_output


def generate_breach_table(
    selected_logic_name, dataframes: dict, mode="pdf_only", report_mode=False
):
    """
    Execute a single logic and return its DataFrame.
    If report_mode=False → save individual Excel output.
    If report_mode=True → skip saving, only collect results.
    """
    logic_entry = logic_registry.get(selected_logic_name)
    if logic_entry and "function" in logic_entry:
        df = logic_entry["function"](dataframes, mode=mode)

        if mode == "full" and df is not None and not df.empty:
            if not report_mode:
                save_excel_output(selected_logic_name, df)

        return df

    print("❌ Selected logic not found.")
    return None
