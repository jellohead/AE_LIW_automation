# get_chart_series_data.py
# Pull series name and values from an existing embedded Excel chart

def get_chart_series_data(chart) -> dict:
    """
    Get series name and values from an existing embedded Excel chart
    :param chart: object
    :return: dict
    """

    series_data = {}
    for series in chart.plots[0].series:
        series_name = series.name
        series_values = [pt for pt in series.values]
        series_data[series_name] = series_values
    return series_data