# slide_15.py

import logging
import pandas as pd
from pptx.chart.data import CategoryChartData
from AE_LIW_automation.config import REPORTING_PERIOD, REPORTING_YEAR, CURRENT_MONTH_TEXT, CURRENT_YEAR
from AE_LIW_automation.helper_modules import (get_chart_object_by_name, get_chart_categories, get_chart_series_data,
                                              get_data_blob_from_chart, get_chart_object)


logger = logging.getLogger(__name__)

def slide_15_updater(df, prs) -> object:
    slide_index = 14

    msg = f"Updating slide {slide_index + 1}"
    width = 40
    print(f"\n{'=' * width}\n{' ' + msg + ' ':=^{width}}\n{'=' * width}\n")
    logger.info(msg)

    question = 'Q11'
    chart_name = 'Content Placeholder 10'

    slide = prs.slides[slide_index]
    chart = get_chart_object_by_name(slide, chart_name)

    # pull blob with chart data out of slide
    workbook, worksheet = get_data_blob_from_chart(chart)
    data = list(worksheet.values)

    # create new dataframe from old chart data
    slide_df = pd.DataFrame(data)
    # set dataframe column labels to first row values then drop first row
    slide_df.columns = slide_df.iloc[0]
    slide_df.drop(slide_df.index[0], inplace=True)
    # set dataframe index labels to values in first column
    slide_df.set_index(slide_df.columns[0], inplace=True)
    # drop the oldest quarter of data
    slide_df = slide_df.iloc[1:].copy()

    # generate new quarter data
    new_quarter_df = pd.DataFrame(index=[f'{REPORTING_PERIOD} {REPORTING_YEAR}\n(N={len(df)})'],
                                  columns=['8', '9', '10', 'sum of displayed values'])
    question_value_counts = df[question].dropna().value_counts(normalize=True).sort_index()
    new_quarter_df['8'] =  question_value_counts.get(8, 0)
    new_quarter_df['9'] =  question_value_counts.get(9, 0)
    new_quarter_df['10'] =  question_value_counts.get(10, 0)
    new_quarter_df['sum of displayed values'] = (
                new_quarter_df['8'].values + new_quarter_df['9'].values + new_quarter_df['10'].values)

    existing_df = pd.concat([slide_df, new_quarter_df])
    # replace NaN with None
    existing_df = existing_df.astype(object).where(pd.notna(existing_df), None)
    # replace '0' values with 'None' to prevent 0% showing up on chart
    existing_df = existing_df.astype(object).replace(0, None)

    # update chart data
    new_chart_data = CategoryChartData()
    new_chart_data.categories = existing_df.index
    for idx, values in existing_df.items():
        new_chart_data.add_series(idx, existing_df[idx].values, number_format='0%')
    chart.replace_data(new_chart_data)

