  METHOD build_xml.
    DATA: lo_scheme_element TYPE REF TO if_ixml_element.
    DATA: lo_font TYPE REF TO if_ixml_element.
    DATA: lo_latin TYPE REF TO if_ixml_element.
    DATA: lo_ea TYPE REF TO if_ixml_element.
    DATA: lo_cs TYPE REF TO if_ixml_element.
    DATA: lo_major TYPE REF TO if_ixml_element.
    DATA: lo_minor TYPE REF TO if_ixml_element.
    DATA: lo_elements TYPE REF TO if_ixml_element.
    FIELD-SYMBOLS: <font> TYPE t_font.
    CHECK io_document IS BOUND.
    lo_elements ?= io_document->find_from_name_ns( name = zcl_excel_theme=>c_theme_elements ).
    IF lo_elements IS BOUND.
      lo_scheme_element ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = zcl_excel_theme_elements=>c_font_scheme
                                                               parent = lo_elements ).
      lo_scheme_element->set_attribute( name = c_name value = font_scheme-name ).

      lo_major ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_major
                                                      parent = lo_scheme_element ).
      IF lo_major IS BOUND.
        lo_latin ?=  io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_latin
                                                         parent = lo_major ).
        lo_latin->set_attribute( name = c_typeface value = font_scheme-major-latin-typeface ).
        IF font_scheme-major-latin-panose IS NOT INITIAL.
          lo_latin->set_attribute( name = c_panose value = font_scheme-major-latin-panose ).
        ENDIF.
        IF font_scheme-major-latin-pitchfamily IS NOT INITIAL.
          lo_latin->set_attribute( name = c_pitchfamily value = font_scheme-major-latin-pitchfamily ).
        ENDIF.
        IF font_scheme-major-latin-charset IS NOT INITIAL.
          lo_latin->set_attribute( name = c_charset value = font_scheme-major-latin-charset ).
        ENDIF.

        lo_ea ?=  io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_ea
                                                         parent = lo_major ).
        lo_ea->set_attribute( name = c_typeface value = font_scheme-major-ea-typeface ).
        IF font_scheme-major-ea-panose IS NOT INITIAL.
          lo_ea->set_attribute( name = c_panose value = font_scheme-major-ea-panose ).
        ENDIF.
        IF font_scheme-major-ea-pitchfamily IS NOT INITIAL.
          lo_ea->set_attribute( name = c_pitchfamily value = font_scheme-major-ea-pitchfamily ).
        ENDIF.
        IF font_scheme-major-ea-charset IS NOT INITIAL.
          lo_ea->set_attribute( name = c_charset value = font_scheme-major-ea-charset ).
        ENDIF.

        lo_cs ?=  io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_cs
                                                      parent = lo_major ).
        lo_cs->set_attribute( name = c_typeface value = font_scheme-major-cs-typeface ).
        IF font_scheme-major-cs-panose IS NOT INITIAL.
          lo_cs->set_attribute( name = c_panose value = font_scheme-major-cs-panose ).
        ENDIF.
        IF font_scheme-major-cs-pitchfamily IS NOT INITIAL.
          lo_cs->set_attribute( name = c_pitchfamily value = font_scheme-major-cs-pitchfamily ).
        ENDIF.
        IF font_scheme-major-cs-charset IS NOT INITIAL.
          lo_cs->set_attribute( name = c_charset value = font_scheme-major-cs-charset ).
        ENDIF.

        LOOP AT font_scheme-major-fonts ASSIGNING <font>.
          IF <font>-script IS NOT INITIAL AND <font>-typeface IS NOT INITIAL.
            CLEAR lo_font.
            lo_font ?=  io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_font
                                                           parent = lo_major ).
            lo_font->set_attribute( name = c_script value = <font>-script ).
            lo_font->set_attribute( name = c_typeface value = <font>-typeface ).
          ENDIF.
        ENDLOOP.
        CLEAR: lo_latin, lo_ea, lo_cs, lo_font.
      ENDIF.

      lo_minor ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_minor
                                                      parent = lo_scheme_element ).
      IF lo_minor IS BOUND.
        lo_latin ?=  io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_latin
                                                         parent = lo_minor ).
        lo_latin->set_attribute( name = c_typeface value = font_scheme-minor-latin-typeface ).
        IF font_scheme-minor-latin-panose IS NOT INITIAL.
          lo_latin->set_attribute( name = c_panose value = font_scheme-minor-latin-panose ).
        ENDIF.
        IF font_scheme-minor-latin-pitchfamily IS NOT INITIAL.
          lo_latin->set_attribute( name = c_pitchfamily value = font_scheme-minor-latin-pitchfamily ).
        ENDIF.
        IF font_scheme-minor-latin-charset IS NOT INITIAL.
          lo_latin->set_attribute( name = c_charset value = font_scheme-minor-latin-charset ).
        ENDIF.

        lo_ea ?=  io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_ea
                                                         parent = lo_minor ).
        lo_ea->set_attribute( name = c_typeface value = font_scheme-minor-ea-typeface ).
        IF font_scheme-minor-ea-panose IS NOT INITIAL.
          lo_ea->set_attribute( name = c_panose value = font_scheme-minor-ea-panose ).
        ENDIF.
        IF font_scheme-minor-ea-pitchfamily IS NOT INITIAL.
          lo_ea->set_attribute( name = c_pitchfamily value = font_scheme-minor-ea-pitchfamily ).
        ENDIF.
        IF font_scheme-minor-ea-charset IS NOT INITIAL.
          lo_ea->set_attribute( name = c_charset value = font_scheme-minor-ea-charset ).
        ENDIF.

        lo_cs ?=  io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_cs
                                                      parent = lo_minor ).
        lo_cs->set_attribute( name = c_typeface value = font_scheme-minor-cs-typeface ).
        IF font_scheme-minor-cs-panose IS NOT INITIAL.
          lo_cs->set_attribute( name = c_panose value = font_scheme-minor-cs-panose ).
        ENDIF.
        IF font_scheme-minor-cs-pitchfamily IS NOT INITIAL.
          lo_cs->set_attribute( name = c_pitchfamily value = font_scheme-minor-cs-pitchfamily ).
        ENDIF.
        IF font_scheme-minor-cs-charset IS NOT INITIAL.
          lo_cs->set_attribute( name = c_charset value = font_scheme-minor-cs-charset ).
        ENDIF.

        LOOP AT font_scheme-minor-fonts ASSIGNING <font>.
          IF <font>-script IS NOT INITIAL AND <font>-typeface IS NOT INITIAL.
            CLEAR lo_font.
            lo_font ?=  io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_font
                                                           parent = lo_minor ).
            lo_font->set_attribute( name = c_script value = <font>-script ).
            lo_font->set_attribute( name = c_typeface value = <font>-typeface ).
          ENDIF.
        ENDLOOP.
      ENDIF.


    ENDIF.
  ENDMETHOD.                    "build_xml