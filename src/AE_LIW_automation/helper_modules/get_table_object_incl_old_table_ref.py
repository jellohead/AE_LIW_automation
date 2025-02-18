def get_table_object_incl_old_table_ref(slide):
    '''
    Provide a slide object and returns the corresponding table object from the PowerPoint slide.
    :param slide:
    :return: object
    '''

    # table_shape: object = None
    for shape in slide.shapes:
        if shape.has_table:
            print(f'{dir(shape) = }')
            table_shape = shape.table
            shape_xml_element = shape._element
            return table_shape, shape_xml_element
