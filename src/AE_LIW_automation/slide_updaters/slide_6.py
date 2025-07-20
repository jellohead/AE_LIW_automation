# slide_6.py
# This file contains the functions for updating the slide 6 of the PowerPoint file
import logging
from pptx.chart.data import CategoryChartData
from AE_LIW_automation.config import REPORTING_PERIOD, REPORTING_YEAR
from AE_LIW_automation.helper_modules import get_chart_object_by_name, get_chart_categories, get_chart_series_data


logger = logging.getLogger(__name__)


def slide_6_updater(df, prs) -> object:
    print('updating slide 6')
    slide_index = 5
    print(
        f'\n================================\n======= Updating slide {slide_index + 1} =======\n================================\n')
    logger.info(f'Updating slide {slide_index + 1}')
    slide = prs.slides[slide_index]
    chart = get_chart_object_by_name(slide, 'Chart 6')
    question = 'Q17'
    old_categories = get_chart_categories(chart)
    existing_series_data = get_chart_series_data(chart)

    current_quarter_chart_data = df[question].dropna().value_counts(normalize=True).sort_index()

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

