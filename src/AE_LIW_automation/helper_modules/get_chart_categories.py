from pptx.enum.chart import XL_CHART_TYPE


def get_chart_categories(chart):
    '''
    Provide a chart object and return the list of categories from the slide.

    :param chart: Chart object
    :return: List of categories from the slide
    '''
    if chart.chart_type in [
        XL_CHART_TYPE.BAR_CLUSTERED,
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        XL_CHART_TYPE.LINE,
        XL_CHART_TYPE.LINE_MARKERS,
        XL_CHART_TYPE.RADAR_MARKERS,
    ]:
        # Access the category data from the first plot
        if chart.plots:
            plot = chart.plots[0]

            # Get categories using list comprehension
            return [category for category in plot.categories]
    # return empty list if chart.chart_type is False
    return []
