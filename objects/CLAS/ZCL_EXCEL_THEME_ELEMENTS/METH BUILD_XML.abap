  METHOD build_xml.
    DATA: lo_theme_element TYPE REF TO if_ixml_element.
    DATA: lo_theme TYPE REF TO if_ixml_element.
    CHECK io_document IS BOUND.
    lo_theme ?= io_document->get_root_element( ).
    IF lo_theme IS BOUND.
      lo_theme_element ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix
                                                                 name   = c_theme_elements
                                                              parent = lo_theme ).

      color_scheme->build_xml( io_document = io_document ).
      font_scheme->build_xml( io_document = io_document ).
      fmt_scheme->build_xml( io_document = io_document ).
    ENDIF.
  ENDMETHOD.