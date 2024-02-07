  METHOD load.
    DATA: lo_scheme_children TYPE REF TO if_ixml_node_list.
    DATA: lo_scheme_iterator TYPE REF TO if_ixml_node_iterator.
    DATA: lo_scheme_element TYPE REF TO if_ixml_element.
    CHECK io_color_scheme  IS NOT INITIAL.

    name = io_color_scheme->get_attribute( name = c_name ).
    lo_scheme_children = io_color_scheme->get_children( ).
    lo_scheme_iterator = lo_scheme_children->create_iterator( ).
    lo_scheme_element ?= lo_scheme_iterator->get_next( ).
    WHILE lo_scheme_element IS BOUND.
      CASE lo_scheme_element->get_name( ).
        WHEN c_dark1.
          dark1 = me->get_color( lo_scheme_element ).
        WHEN c_dark2.
          dark2 = me->get_color( lo_scheme_element ).
        WHEN c_light1.
          light1 = me->get_color( lo_scheme_element ).
        WHEN c_light2.
          light2 = me->get_color( lo_scheme_element ).
        WHEN c_accent1.
          accent1 = me->get_color( lo_scheme_element ).
        WHEN c_accent2.
          accent2 = me->get_color( lo_scheme_element ).
        WHEN c_accent3.
          accent3 = me->get_color( lo_scheme_element ).
        WHEN c_accent4.
          accent4 = me->get_color( lo_scheme_element ).
        WHEN c_accent5.
          accent5 = me->get_color( lo_scheme_element ).
        WHEN c_accent6.
          accent6 = me->get_color( lo_scheme_element ).
        WHEN c_hlink.
          hlink = me->get_color( lo_scheme_element ).
        WHEN c_folhlink.
          folhlink = me->get_color( lo_scheme_element ).
      ENDCASE.
      lo_scheme_element ?= lo_scheme_iterator->get_next( ).
    ENDWHILE.
  ENDMETHOD.                    "load