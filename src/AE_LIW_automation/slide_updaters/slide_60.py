# slide_60.py
# Demographics How would you describe your health chart

import logging
import numpy as np
from pptx.chart.data import CategoryChartData
from AE_LIW_automation.config import REPORTING_PERIOD, REPORTING_YEAR, CURRENT_MONTH_TEXT, CURRENT_YEAR
from AE_LIW_automation.helper_modules import get_chart_object_by_name, get_chart_categories, get_chart_series_data


logger = logging.getLogger(__name__)

def slide_60_updater(df, meta, df_labeled, prs) -> object:
    slide_index = 58

    msg = f"Updating slide {slide_index + 1}"
    width = 40
    print(f"\n{'=' * width}\n{' ' + msg + ' ':=^{width}}\n{'=' * width}\n")
    logger.info(msg)

    slide = prs.slides[slide_index]
    chart = get_chart_object_by_name(slide, 'Chart 6')
    question = 'D8'
    category_label_order = [5, 4, 3, 2, 1]
    old_categories = get_chart_categories(chart)

    existing_series_data = get_chart_series_data(chart)

    current_quarter_chart_data = df[question].dropna().value_counts(normalize=True).reindex(index=category_label_order).replace(np.nan, None)

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

