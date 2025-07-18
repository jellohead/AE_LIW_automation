import logging


logger = logging.getLogger(__name__)


def get_table_shape_by_name(slide: object, table_name: None) -> object:
    '''
    Provide a slide object and returns the corresponding table object from the PowerPoint slide.
    If a table name is provided, it will return the corresponding table object.
    When a table name is not provided, it will return the first table object from the slide.
    :param slide:
    :param table_name: (default is None)
    :return: object
    '''

    # table_shape: object = None
    for shape in slide.shapes:
        if shape.has_table:
            if table_name is not None and shape.name == table_name:
                # if shape.name == table_name:
                    logging.info(f'Shape with "{shape.name}" table name found on this slide.\nTable name: {shape.name}')
                    return shape
            else:
                logging.info(f'Table name not provided.\n{shape.name} was found on this slide\nTable name: {shape.name}')
                return shape
