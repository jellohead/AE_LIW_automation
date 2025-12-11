# constants.py
# Config file for AE LIW automation

from pathlib import Path

# This is .../src/AE_LIW_automation/config/<this_file>.py
THIS_FILE = Path(__file__).resolve()

# Package root: .../src/AE_LIW_automation
PACKAGE_ROOT = THIS_FILE.parent.parent

# Resources directory inside the package
RESOURCES_DIR = PACKAGE_ROOT / "resources"

# Output directory inside the package
OUTPUT_DIR = PACKAGE_ROOT / "output"

# Now define your paths using RESOURCES_DIR
PPTX_INPUT_FILE = RESOURCES_DIR / "9978 LIW Report FY2025 Q2 2025 .pptx"
PPTX_OUTPUT_FILE = OUTPUT_DIR / "AE_LIW_updated.pptx"
DATASET_FILE_PATH = RESOURCES_DIR / "10100 Low Income Weatherization FY25Q3.sav"
EXCEL_FILE = OUTPUT_DIR / "AE_LIW_Excel_output.xlsx"


# Date constants
CURRENT_MONTH_TEXT = 'December'
CURRENT_YEAR = '2025'
REPORTING_YEAR = '2025'
REPORTING_PERIOD = 'Q3'
REPORTING_QUARTER = 'Jul - Sep 2025'
PREVIOUS_PERIOD = 'Q2'