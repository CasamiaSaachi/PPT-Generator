from pptx.util import Inches


def get_layout_positions(product_count, template_number):
    
    """
    Returns a tuple: (positions, image_size, text_height)
    positions: list of (left, top) tuples for placing product images
    """
    spacing_x = Inches(0.5)
    spacing_y = Inches(0.3)
    if product_count == 1:
        if template_number == 1:
            positions = [
            (Inches(1), Inches(4.0)), # x axis and y axis
                        
        ]
        elif template_number == 2:
            positions = [
            (Inches(10.335), Inches(3.85)),

        ]
        elif template_number == 3:
            positions = [
            (Inches(20.335), Inches(3.75)),
        ]
        else:
            raise ValueError("Template must be 1, 2, or 3")
        image_size = Inches(5)
        text_height = Inches(1.5)
    
    elif product_count == 2:
        if template_number == 1:
            positions = [
            (Inches(5.55), Inches(4.1)),
            (Inches(16.114), Inches(4.1)),
            
        ]
        elif template_number == 2:
            positions = [
            (Inches(3.55), Inches(2.0)),
            (Inches(18.114), Inches(6.0)),
            
        ]
        elif template_number == 3:
            positions = [
            (Inches(3.55), Inches(6.0)),
            (Inches(18.114), Inches(2.0)),

        ]
        else:
            raise ValueError("Template must be 1, 2, or 3")
        image_size = Inches(5)
        text_height = Inches(1.5)
    
    elif product_count == 3:
        if template_number == 1:
            positions = [
            (Inches(1.5), Inches(4.1)),
            (Inches(10.835), Inches(4.1)),
            (Inches(20.170), Inches(4.1)),
            
        ]
        elif template_number == 2:
            positions = [
            (Inches(1.5), Inches(1.5)),
            (Inches(10.835), Inches(4.1)),
            (Inches(20.170), Inches(7.0)),

        ]
        elif template_number == 3:
            positions = [
            (Inches(1.5), Inches(7.0)),
            (Inches(10.835), Inches(4.25)),
            (Inches(20.170), Inches(1.5)),

        ]
        else:
            raise ValueError("Template must be 1, 2, or 3")
        image_size = Inches(5)
        text_height = Inches(1.5)

    elif product_count == 4:
        if template_number == 1:
            positions = [
            (Inches(3.9), Inches(1.5)),
            (Inches(18.7), Inches(1.5)),
            (Inches(3.9), Inches(7.6)),
            (Inches(18.7), Inches(7.6)),
        ]
        elif template_number == 2:
            positions = [
            (Inches(1.9), Inches(1.5)),
            (Inches(14.4), Inches(1.5)),
            (Inches(8.4), Inches(7.5)),
            (Inches(20.4), Inches(7.5)),
        ]
        elif template_number == 3:
            positions = [
            (Inches(1.9), Inches(7.5)),
            (Inches(14.4), Inches(7.5)),
            (Inches(8.4), Inches(1.5)),
            (Inches(20.4), Inches(1.5)),
        ]
        else:
            raise ValueError("Template must be 1, 2, or 3")
        image_size = Inches(4.5)
        text_height = Inches(1.5)
    
    elif product_count == 5:
        if template_number == 1:
            positions = [
            (Inches(0.83), Inches(1.5)),
            (Inches(10.835), Inches(1.5)),
            (Inches(20.83), Inches(1.5)),
            (Inches(5.83), Inches(7.75)),
            (Inches(15.83), Inches(7.75)),
        ]
        elif template_number == 2:
            positions = [
            (Inches(0.83), Inches(7.75)),
            (Inches(10.835), Inches(7.75)),
            (Inches(20.83), Inches(7.75)),
            (Inches(5.83), Inches(1.5)),
            (Inches(15.83), Inches(1.5)),
        ]
        elif template_number == 3:  #this also looks like template 2 but with different positions
            positions = [
            (Inches(1.25), Inches(1.5)),
            (Inches(10.835), Inches(1.5)),
            (Inches(20.420), Inches(1.5)),
            (Inches(5.55), Inches(7.75)),
            (Inches(16.114), Inches(7.75)),
        ]
        else:
            raise ValueError("Template must be 1, 2, or 3")
        image_size = Inches(4.5)
        text_height = Inches(1.5)

    elif product_count == 6:
        if template_number == 1:
            positions = [
            (Inches(1.83), Inches(1.5)),
            (Inches(10.835), Inches(1.5)),
            (Inches(19.83), Inches(1.5)),
            (Inches(1.83), Inches(7.75)),
            (Inches(10.835), Inches(7.75)),
            (Inches(19.83), Inches(7.75)),
        ]
        elif template_number == 2:
            positions = [
            (Inches(1), Inches(1.1)),
            (Inches(7), Inches(1.1)),
            (Inches(13), Inches(1.1)),
            (Inches(1), Inches(8.1)),
            (Inches(7), Inches(8.1)),
            (Inches(13), Inches(8.1)),
        ]
        elif template_number == 3:
            positions = [
            (Inches(1), Inches(1.1)),
            (Inches(7), Inches(1.1)),
            (Inches(13), Inches(1.1)),
            (Inches(1), Inches(8.1)),
            (Inches(7), Inches(8.1)),
            (Inches(13), Inches(8.1)),
        ]
        else:
            raise ValueError("Template must be 1, 2, or 3")
        image_size = Inches(4.5)
        text_height = Inches(1.5)
    
    elif product_count == 7:
        if template_number == 1:
            positions = [
            (Inches(0.83), Inches(1.5)),
            (Inches(7.16), Inches(1.5)),
            (Inches(13.49), Inches(1.5)),
            (Inches(19.83), Inches(1.5)),
            (Inches(4.835), Inches(7.75)),
            (Inches(10.835), Inches(7.75)),
            (Inches(16.835), Inches(7.75)),
        ]
        elif template_number == 2:
            positions = [
            (Inches(1), Inches(1.25)),
            (Inches(7), Inches(1.25)),
            (Inches(13), Inches(1.25)),
            (Inches(19), Inches(1.25)),
            (Inches(1), Inches(7.75)),
            (Inches(7), Inches(7.75)),
            (Inches(13), Inches(7.75)),
        ]
        elif template_number == 3:
            positions = [
            (Inches(1), Inches(1.1)),
            (Inches(7), Inches(1.1)),
            (Inches(13), Inches(1.1)),
            (Inches(19), Inches(1.1)),
            (Inches(1), Inches(8.1)),
            (Inches(7), Inches(8.1)),
            (Inches(13), Inches(8.1)),
        ]
        else:
            raise ValueError("Template must be 1, 2, or 3")
        image_size = Inches(3.5)
        text_height = Inches(1.5)


    
    else:
        raise ValueError("Supported product count is 1 to 7 only.")

    return positions, image_size, text_height




