# slide_21.py

import logging
import pandas as pd
from pptx.chart.data import CategoryChartData
from AE_LIW_automation.config import REPORTING_PERIOD, REPORTING_YEAR, CURRENT_MONTH_TEXT, CURRENT_YEAR
from AE_LIW_automation.helper_modules import get_chart_object_by_name, get_chart_categories, get_chart_series_data, get_chart_object


logger = logging.getLogger(__name__)

def slide_21_updater(df, prs) -> object:
    slide_index = 20
    print(
        f'\n================================\n======= Updating slide {slide_index + 1} =======\n================================\n')
    logger.info(f'Updating slide {slide_index + 1}')

    slide = prs.slides[slide_index]
    chart_name = 'Content Placeholder 10'
    chart = get_chart_object_by_name(slide, chart_name)
    question_list = ['Q3_r5', 'Q3_r3']
    old_categories = get_chart_categories(chart)
    existing_series_data_dict = get_chart_series_data(chart)
    # existing_series_data_dict['sum of displayed values'] = existing_series_data_dict.pop('')

    # create a new categories list by dropping oldest labels and inserting new category labels
    current_quarter_category = f'{REPORTING_PERIOD} {REPORTING_YEAR}\n(N={len(df)})'
    new_categories_list =

    # split existing_series_data_dict into two dictionaries
    existing_series_data_dict_left = {k: v[1:4] for k, v in existing_series_data_dict.items()}
    existing_series_data_dict_right = {k: v[5:] for k, v in existing_series_data_dict.items()}

    # append new quarter data to the end list of dictionary items
    for question in question_list:


    existing_df = pd.DataFrame(index=old_categories, data=existing_series_data_dict)[1:]

    new_quarter_df = pd.DataFrame(index=[f'{REPORTING_PERIOD} {REPORTING_YEAR}\n(N={len(df)})'], columns=existing_df.columns)
    question_value_counts = df[question].dropna().value_counts(normalize=True).sort_index()
    new_quarter_df['8'] =  question_value_counts.get(8, 0)
    new_quarter_df['9'] =  question_value_counts.get(9, 0)
    new_quarter_df['10'] =  question_value_counts.get(10, 0)
    new_quarter_df['sum of displayed values'] = new_quarter_df['8'].values + new_quarter_df['9'].values + new_quarter_df['10'].values

    existing_df = pd.concat([existing_df, new_quarter_df])
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

