import logging


logger = logging.getLogger(__name__)

def get_chart_object_by_name(slide:object,chart_name:str) -> object:
    '''
    Provide a slide object and chart_name and returns the corresponding chart object from the powerpoint slide.
    :param slide: slide object
    :param chart_name: string (from PowerPoint Selection Pane)
    :return: chart object
    '''
    for shape in slide.shapes:
        if shape.name == chart_name and shape.has_chart:
            return shape.chart

    logger.info(f'Chart shape {chart_name} not found, returning empty chart object')