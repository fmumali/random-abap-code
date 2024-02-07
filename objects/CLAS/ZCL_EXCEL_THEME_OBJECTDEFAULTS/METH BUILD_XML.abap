  METHOD build_xml.
    DATA: lo_theme TYPE REF TO if_ixml_element.
    DATA: lo_theme_objdef TYPE REF TO if_ixml_element.
    CHECK io_document IS BOUND.
    lo_theme ?= io_document->get_root_element( ).
    CHECK lo_theme IS BOUND.
    IF objectdefaults IS INITIAL.
      lo_theme_objdef ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix
                                                                name   = zcl_excel_theme=>c_theme_object_def
                                                                parent = lo_theme ).
    ELSE.
      lo_theme->append_child( new_child = objectdefaults ).
    ENDIF.
  ENDMETHOD.                    "build_xml