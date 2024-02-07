  METHOD build_xml.
    DATA: lo_node TYPE REF TO if_ixml_node.
    DATA: lo_elements TYPE REF TO if_ixml_element.
    CHECK io_document IS BOUND.
    lo_elements ?= io_document->find_from_name_ns( name = zcl_excel_theme=>c_theme_elements ).
    IF lo_elements IS BOUND.

      IF fmt_scheme IS INITIAL.
        lo_node = parse_string( get_default_fmt( ) ).
        lo_elements->append_child( new_child = lo_node ).
      ELSE.
        lo_elements->append_child( new_child = fmt_scheme ).
      ENDIF.
    ENDIF.
  ENDMETHOD.                    "build_xml