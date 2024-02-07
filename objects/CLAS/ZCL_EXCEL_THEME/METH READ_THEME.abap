  METHOD read_theme.
    DATA: lo_node_theme TYPE REF TO if_ixml_element.
    DATA: lo_theme_children TYPE REF TO if_ixml_node_list.
    DATA: lo_theme_iterator TYPE REF TO if_ixml_node_iterator.
    DATA: lo_theme_element TYPE REF TO if_ixml_element.
    CHECK io_theme_xml IS NOT INITIAL.

    lo_node_theme  = io_theme_xml->get_root_element( )."   find_from_name( name = c_theme ).
    IF lo_node_theme IS BOUND.
      name = lo_node_theme->get_attribute( name = c_theme_name ).
      lo_theme_children = lo_node_theme->get_children( ).
      lo_theme_iterator = lo_theme_children->create_iterator( ).
      lo_theme_element ?= lo_theme_iterator->get_next( ).
      WHILE lo_theme_element IS BOUND.
        CASE lo_theme_element->get_name( ).
          WHEN c_theme_elements.
            elements->load( io_elements = lo_theme_element ).
          WHEN c_theme_object_def.
            objectdefaults->load( io_object_def = lo_theme_element ).
          WHEN c_theme_extra_color.
            extclrschemelst->load( io_extra_color = lo_theme_element ).
          WHEN c_theme_extlst.
            extlst->load( io_extlst = lo_theme_element ).
        ENDCASE.
        lo_theme_element ?= lo_theme_iterator->get_next( ).
      ENDWHILE.
    ENDIF.
  ENDMETHOD.                    "read_theme