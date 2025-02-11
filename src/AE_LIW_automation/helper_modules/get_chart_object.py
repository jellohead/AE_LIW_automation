def get_chart_object(slide: object) -> object:
    '''
    Takes a slide object and returns a chart object.
    :param slide: slide object
    :return: chart object
    '''
    # Iterate through shapes on the slide to find the chart shape
    chart_shape = None
    for shape in slide.shapes:
        if shape.has_chart:
            chart_shape = shape
            break

    # Check if a chart shape was found
    if chart_shape is not None:
        chart = chart_shape.chart
        return chart

    else:
        print("No embedded chart found on the specified slide.")