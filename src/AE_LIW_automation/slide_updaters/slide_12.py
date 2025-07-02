# slide_12.py
import logging
from pptx.chart.data import CategoryChartData
from AE_LIW_automation.config import REPORTING_PERIOD, REPORTING_YEAR, CURRENT_MONTH_TEXT, CURRENT_YEAR
from AE_LIW_automation.helper_modules import get_chart_object_by_name, get_chart_categories, get_chart_series_data


logger = logging.getLogger(__name__)


def slide_12_updater(df, prs) -> object:
    slide_index = 11
    print(
        f'\n================================\n======= Updating slide {slide_index + 1} =======\n================================\n')
    logger.info(f'Updating slide {slide_index + 1}')
    slide = prs.slides[slide_index]
    chart = get_chart_object_by_name(slide, 'Chart 6')
    question_list = ['Q14_r5', 'Q14_r6']
    expected_value_labels = [1, 2, 3, 4, 5]
    old_categories = get_chart_categories(chart)
    existing_series_data = get_chart_series_data(chart)

    # drop oldest data series and convert it to a dictionary
    existing_series_data = dict(list(existing_series_data.items())[1:])

    current_quarter_chart_data =[]
    for question in question_list:
        # generate new quarter data series, removing responses of don't know
        current_quarter_chart_data_new = list(
            df[question]
            .value_counts()
            .reindex(expected_value_labels, fill_value=0)
            .drop(labels=5, errors='ignore')
            .sort_index()
        )
        current_quarter_chart_data += current_quarter_chart_data_new

    # append new quarter data to existing data
    new_key = f'{REPORTING_PERIOD} {REPORTING_YEAR}\n(N={len(df)})'
    existing_series_data[new_key] = current_quarter_chart_data

    # update chart data
    new_chart_data = CategoryChartData()
    new_chart_data.categories = old_categories
    for k, v in existing_series_data.items():
        new_chart_data.add_series(k, v)
    chart.replace_data(new_chart_data)

