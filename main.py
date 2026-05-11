# ─────────────────────────────────────────────────────────────
# MODE CONFIGURATION
# ─────────────────────────────────────────────────────────────
MODE = "pdf_only"  # Options: "full", "pdf_only"

# ─────────────────────────────────────────────────────────────
# STEP 1: Load Logic Modules and Metadata
# ─────────────────────────────────────────────────────────────
from KYC_Viewer.logic_loader import load_all_logics
from KYC_Viewer.logic_metadata import logic_metadata

load_all_logics()

# ─────────────────────────────────────────────────────────────
# STEP 2: Load Registry and UI Renderer
# ─────────────────────────────────────────────────────────────
from KYC_Viewer.utils import logic_registry
from KYC_Viewer.ui import render_ui
from KYC_Viewer.breach_runner import (
    generate_breach_table,
)  # ✅ import from breach_runner

# ─────────────────────────────────────────────────────────────
# STEP 3: Define resource_path for dynamic asset resolution
# ─────────────────────────────────────────────────────────────
import os, sys


def resource_path(relative_path):
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


# ─────────────────────────────────────────────────────────────
# STEP 4: Launch the UI (User will upload files manually)
# ─────────────────────────────────────────────────────────────
if __name__ == "__main__":
    ICON_PATH = resource_path("assets/analysis_13530196.ico")
    render_ui(icon_path=ICON_PATH)
