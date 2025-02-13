# slide_7.py
# This file contains the functions for updating the slide 7 of the Powerpoint file

from pptx.chart.data import CategoryChartData
from config import REPORTING_PERIOD, REPORTING_YEAR, CURRENT_MONTH_TEXT, CURRENT_YEAR
from helper_modules import get_chart_object_by_name, get_chart_categories, get_chart_series_data, get_chart_object


def slide_7_updater(df, prs) -> object:
    print('updating slide 7')
    slide_index = 6
    slide = prs.slides[slide_index]
    # chart = get_chart_object_by_name(slide, 'Content Placeholder 10')
    chart = get_chart_object(slide)
    print(f'{chart = }')
    old_categories = get_chart_categories(chart)
    print(f'{old_categories = }')
    existing_series_data = get_chart_series_data(chart)
    print(f'{existing_series_data = }')

    current_quarter_chart_data = df['Q19'].dropna().value_counts(normalize=True).sort_index()

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

