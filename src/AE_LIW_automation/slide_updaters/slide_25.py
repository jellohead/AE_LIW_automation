# slide_25.py
# This file contains the functions for updating the slide 6 of the Powerpoint file

from pptx.chart.data import CategoryChartData
from AE_LIW_automation.config import REPORTING_PERIOD, REPORTING_YEAR, CURRENT_MONTH_TEXT, CURRENT_YEAR
from AE_LIW_automation.helper_modules import get_chart_object_by_name, get_chart_categories, get_chart_series_data


def slide_25_updater(df, prs) -> object:
    print('updating slide 25')
    slide_index = 24
    slide = prs.slides[slide_index]
    chart = get_chart_object_by_name(slide, 'Chart 6')
    old_categories = get_chart_categories(chart)
    existing_series_data = get_chart_series_data(chart)

    # current_quarter_chart_data = df['Q5'][df['Q5'].isin([1, 2])]
    # print(f'{current_quarter_chart_data = }')
    print(f'{len(df['Q5']) = }')
    current_quarter_chart_data = df['Q5'].dropna().value_counts(normalize=True).sort_index()
    print(f'{current_quarter_chart_data = }')
    # current_quarter_chart_data = current_quarter_chart_data.loc[[1, 2]]

    # drop oldest data series and append new quarter series data
    existing_series_data = dict(list(existing_series_data.items())[1:])
    new_key = f'{REPORTING_PERIOD} {REPORTING_YEAR}\n(n={len(df['Q5'].dropna())})'
    new_value = current_quarter_chart_data.values.tolist()
    existing_series_data[new_key] = new_value

    # update chart data
    new_chart_data = CategoryChartData()
    new_chart_data.categories = old_categories
    for k, v in existing_series_data.items():
        new_chart_data.add_series(k, v, number_format='0%')
    chart.replace_data(new_chart_data)

