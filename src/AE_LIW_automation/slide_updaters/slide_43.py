# slide_43.py
# This file contains the functions for updating the slide 43 of the PowerPoint file
import pandas as pd
from pptx.chart.data import CategoryChartData
from config import REPORTING_PERIOD, REPORTING_YEAR, CURRENT_MONTH_TEXT, CURRENT_YEAR
from helper_modules import get_chart_object_by_name, get_chart_categories, get_chart_series_data


# TODO slide 43 and possibly 44 not working

def slide_43_updater(df, meta, df_labeled, prs) -> object:
    print('updating slide 43')
    slide_index = 42
    slide = prs.slides[slide_index]
    chart = get_chart_object_by_name(slide, 'Content Placeholder 8')
    old_categories = get_chart_categories(chart)
    print(f'{old_categories = }')
    print(f'{meta.variable_value_labels['Q1'] = }')
    existing_series_data = get_chart_series_data(chart)

    current_quarter_chart_data = df_labeled['Q1'].value_counts(normalize=True).sort_index()
    print(f'{type(current_quarter_chart_data) = }\n{current_quarter_chart_data = }')

    # map current_quarter_chart_data indexes to meta.variable_value_labels values

    # drop oldest data series and append new quarter series data
    existing_series_data = dict(list(existing_series_data.items())[:-1])
    # existing_series_data = dict(list(existing_series_data.items()))
    print(f'{existing_series_data = }')
    combined_df = pd.DataFrame(existing_series_data, index=old_categories)
    # combined_df.set_index(old_categories, inplace=True)
    # combined_df.set_index(old_categories,  inplace=True)
    print(f'{combined_df = }')

    new_key = f'{REPORTING_PERIOD} {REPORTING_YEAR}\n(N={len(df)})'
    new_value = current_quarter_chart_data.values.tolist()
    current_quarter_df = pd.DataFrame(current_quarter_chart_data.values,
                                      index=current_quarter_chart_data.keys(),
                                      columns=[new_key])
    print(f'{current_quarter_df = }')

    # combined_df.merge(current_quarter_df, left_index=True, right_index=True)
    df_merged = current_quarter_df.merge(combined_df, left_index=True, right_index=True, how='outer')
    print(f'{df_merged = }')


    # existing_series_data[new_key] = new_value

    # update chart data
    new_chart_data = CategoryChartData()
    new_chart_data.categories = old_categories
    for k, v in existing_series_data.items():
        new_chart_data.add_series(k, v, number_format='0%')
    chart.replace_data(new_chart_data)

