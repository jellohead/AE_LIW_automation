# slide_10.py
import logging
from pptx.chart.data import CategoryChartData
from AE_LIW_automation.config import REPORTING_PERIOD, REPORTING_YEAR, CURRENT_MONTH_TEXT, CURRENT_YEAR
from AE_LIW_automation.helper_modules import get_chart_object_by_name, get_chart_categories, get_chart_series_data


logger = logging.getLogger(__name__)


def slide_10_updater(meta, df, df_labeled, prs) -> object:
    slide_index = 9
    print(
        f'\n================================\n======= Updating slide {slide_index + 1} =======\n================================\n')
    logger.info(f'Updating slide {slide_index + 1}')
    slide = prs.slides[slide_index]
    chart = get_chart_object_by_name(slide, 'Chart 6')
    question_list = ['Q14_r1', 'Q14_r2']
    old_categories = get_chart_categories(chart)
    existing_series_data = get_chart_series_data(chart)

    # drop oldest data series and convert it to a dictionary
    existing_series_data = dict(list(existing_series_data.items())[1:])

    # generate new quarter data series, ignoring responses of don't know
    current_quarter_chart_data_left = list(df[question_list[0]].value_counts().drop(labels=5, errors='ignore').sort_index())
    current_quarter_chart_data_right = list(df[question_list[1]].value_counts().drop(labels=5, errors='ignore').sort_index())

    # append new quarter data to existing data
    new_key = f'{REPORTING_PERIOD} {REPORTING_YEAR}\n(N={len(df)})'
    existing_series_data[new_key] = list(current_quarter_chart_data_left) + list(current_quarter_chart_data_right)

    # update chart data
    new_chart_data = CategoryChartData()
    new_chart_data.categories = old_categories
    for k, v in existing_series_data.items():
        new_chart_data.add_series(k, v)
    chart.replace_data(new_chart_data)

