# slide_43.py
# This file contains the functions for updating the slide 43 of the PowerPoint file
import logging

import numpy as np
import pandas as pd
from pptx.chart.data import CategoryChartData

from AE_LIW_automation.helper_modules import get_data_blob_from_chart
from src.AE_LIW_automation.config import REPORTING_PERIOD, REPORTING_YEAR, CURRENT_MONTH_TEXT, CURRENT_YEAR
from src.AE_LIW_automation.helper_modules import get_chart_object_by_name, get_chart_categories, get_chart_series_data


logger = logging.getLogger(__name__)

# TODO slide 43 and possibly 44 not working

def slide_43_updater(df, meta, df_labeled, prs) -> object:
    slide_index = 42

    msg = f"Updating slide {slide_index + 1}"
    width = 40
    print(f"\n{'=' * width}\n{' ' + msg + ' ':=^{width}}\n{'=' * width}\n")
    logger.info(msg)

    question = 'Q1'
    last_rows_list = ['All other', 'Do not know', 'None']
    chart_name = 'Content Placeholder 8'

    slide = prs.slides[slide_index]
    chart = get_chart_object_by_name(slide, chart_name)
    # old_categories = get_chart_categories(chart)


    # pull blob with chart data out of side
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
    slide_df = slide_df.iloc[:,:3].copy()

    # existing_data = []
    # for item in data:
    #     # drop oldest column of data and remove extra columns that are filled with None values
    #     new_item = item [:4]
    #     if any(new_item):
    #         existing_data.append(new_item)
    #
    # existing_data_df = pd.DataFrame(data=existing_data)
    # existing_data_df.index = existing_data_df.iloc[:, 0]
    # existing_data_df.drop([0], axis=1, inplace=True)
    # existing_data_df.columns = existing_data_df.iloc[0]
    # existing_data_df = existing_data_df.iloc[1:]

    # generate new quarter data
    current_quarter_col_name = f'{REPORTING_PERIOD} {REPORTING_YEAR}\n(N={len(df)})'
    current_quarter_chart_data = df_labeled[question].value_counts(normalize=True).sort_index()
    current_quarter_chart_data.rename(current_quarter_col_name, inplace=True)
    # current_quarter_chart_data_df = current_quarter_chart_data.to_frame(name=current_quarter_col_name)

    # combine old and new data into a dataframe
    combined_df = pd.concat([current_quarter_chart_data, slide_df], axis=1)
    combined_df.replace({np.nan: None}, inplace=True)

    # reorder dataframe to start with last rows and sort remaining rows
    last_rows_mask = combined_df.index.isin(last_rows_list)
    last_rows_df = combined_df[last_rows_mask]
    # sort last rows to match order in last_rows_list
    last_rows_df = last_rows_df.reindex(last_rows_list)

    # pull all rows that are not part of the last rows dataframe
    combined_df = combined_df[~last_rows_mask].sort_values(by=current_quarter_col_name, ascending=True)

    # concat both dataframes into a single dataframe and clean up the data
    combined_df_sorted = pd.concat([last_rows_df, combined_df]).fillna(0)

    combined_df_sorted = combined_df_sorted[(combined_df_sorted !=0).any(axis=1)]

    # update chart data
    new_chart_data = CategoryChartData()
    new_chart_data.categories = list(combined_df_sorted.index)
    for column in combined_df_sorted.columns:
        new_chart_data.add_series(column, combined_df_sorted[column], number_format='0%')
    chart.replace_data(new_chart_data)