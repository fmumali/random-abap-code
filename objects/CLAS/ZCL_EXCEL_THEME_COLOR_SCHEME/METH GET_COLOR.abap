  METHOD get_color.
    DATA: lo_color_children TYPE REF TO if_ixml_node_list.
    DATA: lo_color_iterator TYPE REF TO if_ixml_node_iterator.
    DATA: lo_color_element TYPE REF TO if_ixml_element.
    CHECK io_object  IS NOT INITIAL.

    lo_color_children = io_object->get_children( ).
    lo_color_iterator = lo_color_children->create_iterator( ).
    lo_color_element ?= lo_color_iterator->get_next( ).
    IF lo_color_element IS BOUND.
      CASE lo_color_element->get_name( ).
        WHEN c_srgbcolor.
          rv_color-srgb = lo_color_element->get_attribute( name = c_val ).
        WHEN c_syscolor.
          rv_color-syscolor-val = lo_color_element->get_attribute( name = c_val ).
          rv_color-syscolor-lastclr = lo_color_element->get_attribute( name = c_lastclr ).
      ENDCASE.
    ENDIF.
  ENDMETHOD.                    "get_color