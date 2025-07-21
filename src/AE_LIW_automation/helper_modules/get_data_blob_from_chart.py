# get_data_blob_from_chart.py
# pull blob with chart data from PowerPoint slide

import io
from openpyxl import load_workbook

from typing import Tuple
from pptx.chart.chart import Chart
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from AE_LIW_automation.config import EXCEL_FILE


def get_data_blob_from_chart(chart: object) -> Tuple[Workbook, Worksheet]:
    """
    Pulls the blob with chart data from PowerPoint slide

    Parameters
    ----------
    chart : pptx.chart.chart.Chart
        PowerPoint chart object

    Returns
    -------
    workbook : openpyxl.workbook.workbook.Workbook
        Openpyxl workbook object loaded from the blob with chart data
    worksheet : openpyxl.worksheet.worksheet.Worksheet
        Openpyxl worksheet object loaded from the blob with chart data
    """
    xlsx_blob = chart.part.chart_workbook.xlsx_part.blob
    workbook = load_workbook(io.BytesIO(xlsx_blob))
    workbook.save(EXCEL_FILE)
    worksheet = workbook.active
    print(f'Workbook loaded and worksheet = {worksheet.title}')

    return workbook, worksheet