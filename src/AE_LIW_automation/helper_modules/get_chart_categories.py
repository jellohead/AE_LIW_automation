import logging
from pptx.enum.chart import XL_CHART_TYPE


logger = logging.getLogger(__name__)


def get_chart_categories(chart):
    '''
    Provide a chart object and return the list of categories from the slide.

    :param chart: Chart object
    :return: List of categories from the slide
    '''
    if chart.chart_type in [
        XL_CHART_TYPE.AREA,
        XL_CHART_TYPE.AREA_STACKED,
        XL_CHART_TYPE.AREA_STACKED_100,
        XL_CHART_TYPE.BAR_CLUSTERED,
        XL_CHART_TYPE.BAR_STACKED,
        XL_CHART_TYPE.BAR_STACKED_100,
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        XL_CHART_TYPE.COLUMN_STACKED,
        XL_CHART_TYPE.COLUMN_STACKED_100,
        XL_CHART_TYPE.LINE,
        XL_CHART_TYPE.LINE_MARKERS,
        XL_CHART_TYPE.LINE_MARKERS_STACKED,
        XL_CHART_TYPE.LINE_MARKERS_STACKED_100,
        XL_CHART_TYPE.LINE_STACKED,
        XL_CHART_TYPE.LINE_STACKED_100,
        XL_CHART_TYPE.RADAR,
        XL_CHART_TYPE.RADAR_MARKERS,
        XL_CHART_TYPE.RADAR_FILLED,
        XL_CHART_TYPE.STOCK_HLC,
        XL_CHART_TYPE.STOCK_OHLC,
        XL_CHART_TYPE.STOCK_VHLC,
        XL_CHART_TYPE.STOCK_VOHLC,
    ]:
        # Access the category data from the first plot
        if chart.plots:
            plot = chart.plots[0]

            # Get categories using list comprehension
            return [category for category in plot.categories]
    # return empty list if chart.chart_type is False
    logger.info('No categories found, returning empty list')
    return []
