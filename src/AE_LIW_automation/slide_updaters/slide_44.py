# slide_44.py
# This file contains the functions for updating the slide 44 of the PowerPoint file
import logging

import numpy as np
import pandas as pd
from pptx.chart.data import CategoryChartData

from AE_LIW_automation.helper_modules import get_data_blob_from_chart
from src.AE_LIW_automation.config import REPORTING_PERIOD, REPORTING_YEAR, CURRENT_MONTH_TEXT, CURRENT_YEAR
from src.AE_LIW_automation.helper_modules import get_chart_object_by_name, get_chart_categories, get_chart_series_data


logger = logging.getLogger(__name__)


def slide_44_updater(df, meta, df_labeled, prs) -> object:
    slide_index = 43
    print(
        f'\n================================\n======= Updating slide {slide_index + 1} =======\n================================\n')
    logger.info(f'Updating slide {slide_index + 1}')
    slide = prs.slides[slide_index]
    chart = get_chart_object_by_name(slide, 'Content Placeholder 8')
    old_categories = get_chart_categories(chart)
    question_list = ['Q2_1', 'Q2_2', 'Q2_3', 'Q2_4', 'Q2_5', 'Q2_6', 'Q2_7', 'Q2_8', 'Q2_9', 'Q2_10', 'Q2_11', 'Q2_12']
    last_rows_list = ['Other Mention']


    # pull old chart data blob
    workbook, worksheet = get_data_blob_from_chart(chart)
    old_data = list(worksheet.values)
    existing_data = []
    for item in old_data:
        # drop oldest column of data and remove extra columns that are filled with None values
        new_item = item[:1] + item[2:5]
        if any(new_item):
            existing_data.append(new_item)

    existing_data_df = pd.DataFrame(data=existing_data)
    existing_data_df.index = existing_data_df.iloc[:, 0]
    existing_data_df.drop([0], axis=1, inplace=True)
    existing_data_df.columns = existing_data_df.iloc[0]
    existing_data_df = existing_data_df.iloc[1:]

    # generate new quarter data
    new_key = f'{REPORTING_PERIOD} {REPORTING_YEAR}\n(N={len(df)})'
    current_quarter_chart_data_df = pd.DataFrame(columns=[new_key])
    for question in question_list:
        row_label = (meta.column_names_to_labels[question].split('? ', 1)[1].strip())
        current_quarter_chart_data_df.loc[row_label]= df_labeled[question].value_counts(normalize=True).get('Checked', 0)
    # current_quarter_chart_data_df = current_quarter_chart_data.to_frame(name=new_key)

    # combine old and new data into a dataframe
    combined_df = pd.concat([existing_data_df ,current_quarter_chart_data_df], axis=1)
    merged_df = pd.merge(existing_data_df, current_quarter_chart_data_df, left_index=True, right_index=True, how='outer')
    combined_df.replace({np.nan: None}, inplace=True)

    # reorder dataframe to start with last rows and sort remaining rows
    last_rows_mask = combined_df.index.isin(last_rows_list)
    last_rows_df = combined_df[last_rows_mask]
    # sort last rows to match order in last_rows_list
    last_rows_df = last_rows_df.reindex(last_rows_list)

    # pull all rows that are not part of the last rows dataframe
    combined_df = combined_df[~last_rows_mask].sort_values(by=new_key, na_position='first', ascending=True)

    # concat both dataframes into a single dataframe and clean up the data
    combined_df_sorted = pd.concat([last_rows_df, combined_df]).replace({np.nan: None}).dropna(how='all')

    # update chart data
    new_chart_data = CategoryChartData()
    new_chart_data.categories = list(combined_df_sorted.index)
    for column in combined_df_sorted.columns:
        new_chart_data.add_series(column, combined_df_sorted[column], number_format='0%')
    chart.replace_data(new_chart_data)