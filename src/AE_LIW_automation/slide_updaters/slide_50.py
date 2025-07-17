# slide_50.py
# This file contains the functions for updating the slide 50 of the Powerpoint file

import logging
from pptx.chart.data import CategoryChartData
from AE_LIW_automation.config import REPORTING_PERIOD, REPORTING_YEAR, CURRENT_MONTH_TEXT, CURRENT_YEAR
from AE_LIW_automation.helper_modules import get_chart_object_by_name, get_chart_categories, get_chart_series_data


logger = logging.getLogger(__name__)

def slide_50_updater(meta, df, df_labeled, prs) -> object:
    slide_index = 49
    print(
        f'\n================================\n======= Updating slide {slide_index + 1} =======\n================================\n')
    logger.info(f'Updating slide {slide_index + 1}')
    slide = prs.slides[slide_index]
    question = 'D11'
    chart = get_chart_object_by_name(slide, 'Chart 6')
    old_categories = get_chart_categories(chart)

    # remove extra leading and trailing spaces from category labels
    old_categories_cleaned = []
    for category in old_categories:
        old_categories_cleaned.append(category.strip())

    # get the order of the old chart categories
    category_label_order = []
    d11_value_labels = meta.variable_value_labels.get(question)
    for category in old_categories_cleaned:
        for k, v in d11_value_labels.items():
            if category in v:
                category_label_order.append(k)

    print(f'{category_label_order = }')

    existing_series_data = get_chart_series_data(chart)
    print(f'{existing_series_data = }')

    current_quarter_chart_data = df[question].value_counts(normalize=True).sort_index()
    print(f'{type(current_quarter_chart_data) = } {current_quarter_chart_data = }')

    # reindex the current quarter data to match the order of the old chart data
    current_quarter_chart_data = current_quarter_chart_data.reindex(index=category_label_order)
    print(f'{current_quarter_chart_data = }')

    # drop oldest data series and append new quarter series data
    existing_series_data = dict(list(existing_series_data.items())[1:])
    new_key = f'{REPORTING_PERIOD} {REPORTING_YEAR}\n(N={len(df)})'
    new_value = current_quarter_chart_data.values.tolist()
    existing_series_data[new_key] = new_value

    # update chart data
    new_chart_data = CategoryChartData()
    new_chart_data.categories = old_categories
    for k, v in existing_series_data.items():
        new_chart_data.add_series(k, v, number_format='0%')
    chart.replace_data(new_chart_data)

