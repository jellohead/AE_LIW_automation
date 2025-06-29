# slide_14.py

from pptx.chart.data import CategoryChartData
from AE_LIW_automation.config import REPORTING_PERIOD, REPORTING_YEAR, CURRENT_MONTH_TEXT, CURRENT_YEAR
from AE_LIW_automation.helper_modules import get_chart_object_by_name, get_chart_categories, get_chart_series_data



def slide_14_updater(df, prs) -> object:
    print('updating slide 14')
    slide_index = 13
    slide = prs.slides[slide_index]
    chart = get_chart_object_by_name(slide, 'Chart 6')
    old_categories = get_chart_categories(chart)
    existing_series_data = get_chart_series_data(chart)

    current_quarter_chart_data = df['Q14_r9'].value_counts().sort_index()

    # define expected index range of 1 to 4
    expected_index = range(1, 5)
    # populate data ensuring each index has a value, filling in 0 where no value results from query
    current_quarter_chart_data = current_quarter_chart_data.reindex(expected_index, fill_value=0)
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
        new_chart_data.add_series(k, v)
    chart.replace_data(new_chart_data)

