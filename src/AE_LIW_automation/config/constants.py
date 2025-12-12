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
# 2025 Q3 Report
# PPTX_INPUT_FILE = RESOURCES_DIR / "9978 LIW Report FY2025 Q2 2025 .pptx"
# PPTX_OUTPUT_FILE = OUTPUT_DIR / "AE_LIW_updated.pptx"
# DATASET_FILE_PATH = RESOURCES_DIR / "10100 Low Income Weatherization FY25Q3.sav"
# EXCEL_FILE = OUTPUT_DIR / "AE_LIW_Excel_output.xlsx"

# 2025 Q1 Report
# PPTX_INPUT_FILE = RESOURCES_DIR / "9826 LIW Report FY2024 Q4 02122024.pptx"
# PPTX_OUTPUT_FILE = OUTPUT_DIR / "AE_LIW_updated_FY25_Q1.pptx"
# DATASET_FILE_PATH = RESOURCES_DIR / "9930 Low Income Weatherization FY25Q1.sav"
# EXCEL_FILE = OUTPUT_DIR / "AE_LIW_Excel_output.xlsx"

# 2025 Q2 Report
PPTX_INPUT_FILE = OUTPUT_DIR / "AE_LIW_updated_FY25_Q1.pptx"
PPTX_OUTPUT_FILE = OUTPUT_DIR / "AE_LIW_updated_FY25_Q2.pptx"
DATASET_FILE_PATH = RESOURCES_DIR / "9978 Low Income Weatherization FY25Q2.sav"
EXCEL_FILE = OUTPUT_DIR / "AE_LIW_Excel_output.xlsx"





# Date constants 2025 Q3
# CURRENT_MONTH_TEXT = 'December'
# CURRENT_YEAR = '2025'
# REPORTING_YEAR = '2025'
# REPORTING_PERIOD = 'Q3'
# REPORTING_QUARTER = 'Jul - Sep 2025'
# PREVIOUS_PERIOD = 'Q2'

# # Date constants 2025 Q1
# CURRENT_MONTH_TEXT = 'January'
# CURRENT_YEAR = '2025'
# REPORTING_YEAR = '2025'
# REPORTING_PERIOD = 'Q1'
# REPORTING_QUARTER = 'Oct - Dec 2024'
# PREVIOUS_PERIOD = 'Q4'

# Date constants 2025 Q2
CURRENT_MONTH_TEXT = 'April'
CURRENT_YEAR = '2025'
REPORTING_YEAR = '2025'
REPORTING_PERIOD = 'Q2'
REPORTING_QUARTER = 'Jan - Mar 2025'
PREVIOUS_PERIOD = 'Q1'