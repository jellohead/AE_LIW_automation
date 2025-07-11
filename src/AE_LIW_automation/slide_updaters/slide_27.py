# slide_27.py

import logging
from pptx.chart.data import CategoryChartData
from AE_LIW_automation.config import REPORTING_PERIOD, REPORTING_YEAR, CURRENT_MONTH_TEXT, CURRENT_YEAR
from AE_LIW_automation.helper_modules import get_chart_object_by_name, get_chart_categories, get_chart_series_data


logger = logging.getLogger(__name__)


# TODO: refactor


def slide_27_updater(df, prs) -> object:
    slide_index = 26
    print(
        f'\n================================\n======= Updating slide {slide_index + 1} =======\n================================\n')
    logger.info(f'Updating slide {slide_index + 1}')

    slide = prs.slides[slide_index]
    chart_name = 'Content Placeholder 10'
    chart = get_chart_object_by_name(slide, chart_name)
    question_left = 'Q7_r3'
    question_right = 'Q7_r5'
    old_categories = get_chart_categories(chart)
    existing_series_data_dict = get_chart_series_data(chart)

    # create a new categories list by dropping oldest labels and inserting new category labels
    current_quarter_category = f'{REPORTING_PERIOD} {REPORTING_YEAR}\n(N={len(df)})'
    new_categories_list = [j for i, j in enumerate(old_categories) if i not in [0, 4]]
    new_categories_list_copy = new_categories_list[:]
    new_categories_list_copy.insert(3, current_quarter_category)
    new_categories_list_copy.insert(7, current_quarter_category)

    # split existing_series_data_dict into two dictionaries, dropping oldest data from values
    existing_series_data_dict_left = {k: v[1:4] for k, v in existing_series_data_dict.items()}
    existing_series_data_dict_right = {k: v[5:] for k, v in existing_series_data_dict.items()}

    # append new quarter data to the end list of dictionary items
    new_quarter_dict_left = {}
    new_quarter_dict_right = {}

    question_value_counts = df[question_left].dropna().value_counts(normalize=True).sort_index()
    new_quarter_dict_left['8'] = question_value_counts.get(8, 0)
    existing_series_data_dict_left['8'].append(new_quarter_dict_left['8'])
    new_quarter_dict_left['9'] = question_value_counts.get(9, 0)
    existing_series_data_dict_left['9'].append(new_quarter_dict_left['9'])
    new_quarter_dict_left['10'] = question_value_counts.get(10, 0)
    existing_series_data_dict_left['10'].append(new_quarter_dict_left['10'])
    new_quarter_dict_left['sum of displayed values'] = new_quarter_dict_left['8'] + new_quarter_dict_left['9'] + \
                                                new_quarter_dict_left['10']
    existing_series_data_dict_left['sum of displayed values'].append(new_quarter_dict_left['sum of displayed values'])

    question_value_counts = df[question_right].dropna().value_counts(normalize=True).sort_index()
    new_quarter_dict_right['8'] = question_value_counts.get(8, 0)
    existing_series_data_dict_right['8'].append(new_quarter_dict_right['8'])
    new_quarter_dict_right['9'] = question_value_counts.get(9, 0)
    existing_series_data_dict_right['9'].append(new_quarter_dict_right['9'])
    new_quarter_dict_right['10'] = question_value_counts.get(10, 0)
    existing_series_data_dict_right['10'].append(new_quarter_dict_right['10'])
    new_quarter_dict_right['sum of displayed values'] = new_quarter_dict_right['8'] + new_quarter_dict_right['9'] + \
                                                       new_quarter_dict_right['10']
    existing_series_data_dict_right['sum of displayed values'].append(new_quarter_dict_right['sum of displayed values'])

    # combine left and right series data dicts
    combined_series_data_dict = {key: existing_series_data_dict_left[key] + existing_series_data_dict_right[key] for key in existing_series_data_dict_left}
    # replace 0 with None to prevent showing 0% values on charts
    for key, value in combined_series_data_dict.items():
        combined_series_data_dict[key] = [None if v == 0 else v for v in value]

    # update chart data
    new_chart_data = CategoryChartData()
    new_chart_data.categories = new_categories_list_copy
    for k, v in combined_series_data_dict.items():
        new_chart_data.add_series(k, v, number_format='0%')

    chart.replace_data(new_chart_data)

