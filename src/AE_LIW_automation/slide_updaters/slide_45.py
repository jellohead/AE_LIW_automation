# slide_45.py
import logging

import numpy as np
import pandas as pd
from pptx.chart.data import CategoryChartData

from AE_LIW_automation.helper_modules import get_data_blob_from_chart
from src.AE_LIW_automation.config import REPORTING_PERIOD, REPORTING_YEAR, CURRENT_MONTH_TEXT, CURRENT_YEAR
from src.AE_LIW_automation.helper_modules import get_chart_object_by_name, get_chart_categories, get_chart_series_data


logger = logging.getLogger(__name__)


def slide_45_updater(df, meta, df_labeled, prs) -> object:
    slide_index = 44

    msg = f"Updating slide {slide_index + 1}"
    width = 40
    print(f"\n{'=' * width}\n{' ' + msg + ' ':=^{width}}\n{'=' * width}\n")
    logger.info(msg)

    chart_name = 'Content Placeholder 8'
    question_list = [f'Q29_{i}' for i in range(1, 9)]
    last_rows_list = ['All other']
    label_sub_dict = {'Other Mention': 'All other',
                      'Austin Energy’s website': 'Austin Energy\'s website',
                      }

    slide = prs.slides[slide_index]
    chart = get_chart_object_by_name(slide, 'Content Placeholder 8')

    # pull old chart data blob
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
    slide_df = slide_df.iloc[:, 1:].copy()

    # generate new quarter data
    current_quarter_col_name = f'{REPORTING_PERIOD} {REPORTING_YEAR}\n(N={len(df)})'
    current_quarter_chart_series = pd.Series(name=current_quarter_col_name)
    for question in question_list:
        row_label = (meta.column_names_to_labels[question].split('? ', 1)[1].strip())
        current_quarter_chart_series[row_label] = df_labeled[question].value_counts(normalize=True).get('Checked', 0)

    current_quarter_chart_series.rename(index=label_sub_dict, inplace=True)

    # combine old and new data into a dataframe
    combined_df = pd.concat([slide_df, current_quarter_chart_series], axis=1)

    # reorder dataframe to start with last rows and sort remaining rows
    last_rows_mask = combined_df.index.isin(last_rows_list)
    last_rows_df = combined_df[last_rows_mask]
    # sort last rows to match order in last_rows_list
    last_rows_df = last_rows_df.reindex(last_rows_list)

    # pull all rows that are not part of the last rows dataframe
    combined_df = combined_df[~last_rows_mask].sort_values(by=current_quarter_col_name, ascending=False)

    # concat both dataframes into a single dataframe and clean up the data
    combined_df_sorted = pd.concat([combined_df, last_rows_df]).replace({np.nan: 0})
    # drop rows that are all 0 values
    combined_df_sorted = combined_df_sorted[(combined_df_sorted != 0).any(axis=1)]

    # update chart data
    new_chart_data = CategoryChartData()
    new_chart_data.categories = list(combined_df_sorted.index)
    for column in combined_df_sorted.columns:
        new_chart_data.add_series(column, combined_df_sorted[column], number_format='0%')
    chart.replace_data(new_chart_data)