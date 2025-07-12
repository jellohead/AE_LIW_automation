import logging


logger = logging.getLogger(__name__)


def get_table_object(slide: object) -> object:
    '''
    Provide a slide object and returns the corresponding table object from the PowerPoint slide.
    :param slide:
    :return: object
    '''

    table_shape: object = None
    for shape in slide.shapes:
        if shape.has_table:
            logging.info(f'Table name: {shape.name}')
            table_shape = shape.table
            return table_shape
