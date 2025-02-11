# slide_6.py
# This file contains the functions for updating the slide 6 of the Powerpoint file

from helper_modules import get_chart_object_by_name, get_chart_categories


def slide_6_updater(df, prs) -> object:
    print('updating slide 6')
    slide_index = 5
    slide = prs.slides[slide_index]

    chart = get_chart_object_by_name(slide, 'Chart 6')

    categories = get_chart_categories(chart)

    chart_data = chart.chart_data
    chart_data.categories = categories
