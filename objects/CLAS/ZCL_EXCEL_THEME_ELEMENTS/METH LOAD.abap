  METHOD load.
    DATA: lo_elements_children TYPE REF TO if_ixml_node_list.
    DATA: lo_elements_iterator TYPE REF TO if_ixml_node_iterator.
    DATA: lo_elements_element TYPE REF TO if_ixml_element.
    CHECK io_elements IS NOT INITIAL.

    lo_elements_children = io_elements->get_children( ).
    lo_elements_iterator = lo_elements_children->create_iterator( ).
    lo_elements_element ?= lo_elements_iterator->get_next( ).
    WHILE lo_elements_element IS BOUND.
      CASE lo_elements_element->get_name( ).
        WHEN c_color_scheme.
          color_scheme->load( io_color_scheme = lo_elements_element ).
        WHEN c_font_scheme.
          font_scheme->load( io_font_scheme = lo_elements_element ).
        WHEN c_fmt_scheme.
          fmt_scheme->load( io_fmt_scheme = lo_elements_element ).
      ENDCASE.
      lo_elements_element ?= lo_elements_iterator->get_next( ).
    ENDWHILE.
  ENDMETHOD.                    "load