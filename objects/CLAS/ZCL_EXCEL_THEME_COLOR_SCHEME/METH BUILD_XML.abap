  METHOD build_xml.
    DATA: lo_scheme_element TYPE REF TO if_ixml_element.
    DATA: lo_color TYPE REF TO if_ixml_element.
    DATA: lo_syscolor TYPE REF TO if_ixml_element.
    DATA: lo_srgb TYPE REF TO if_ixml_element.
    DATA: lo_elements TYPE REF TO if_ixml_element.

    CHECK io_document IS BOUND.
    lo_elements ?= io_document->find_from_name_ns( name   = zcl_excel_theme=>c_theme_elements ).
    IF lo_elements IS BOUND.
      lo_scheme_element ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix
                                                                  name   = zcl_excel_theme_elements=>c_color_scheme
                                                               parent = lo_elements ).
      lo_scheme_element->set_attribute( name = c_name value = name ).

      " Adding colors to scheme
      lo_color ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix
                                                                  name   = c_dark1
                                                      parent = lo_scheme_element ).
      IF lo_color IS BOUND.
        IF dark1-srgb IS NOT INITIAL.
          lo_srgb ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_srgbcolor
                                                         parent = lo_color ).
          lo_srgb->set_attribute( name = c_val value = dark1-srgb ).
        ELSE.
          lo_syscolor ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_syscolor
                                                          parent = lo_color ).
          lo_syscolor->set_attribute( name = c_val value = dark1-syscolor-val ).
          lo_syscolor->set_attribute( name = c_lastclr value = dark1-syscolor-lastclr ).
        ENDIF.
        CLEAR: lo_color, lo_srgb, lo_syscolor.
      ENDIF.

      lo_color ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_light1
                                                      parent = lo_scheme_element ).
      IF lo_color IS BOUND.
        IF light1-srgb IS NOT INITIAL.
          lo_srgb ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_srgbcolor
                                                         parent = lo_color ).
          lo_srgb->set_attribute( name = c_val value = light1-srgb ).
        ELSE.
          lo_syscolor ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_syscolor
                                                          parent = lo_color ).
          lo_syscolor->set_attribute( name = c_val value = light1-syscolor-val ).
          lo_syscolor->set_attribute( name = c_lastclr value = light1-syscolor-lastclr ).
        ENDIF.
        CLEAR: lo_color, lo_srgb, lo_syscolor.
      ENDIF.


      lo_color ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_dark2
                                                      parent = lo_scheme_element ).
      IF lo_color IS BOUND.
        IF dark2-srgb IS NOT INITIAL.
          lo_srgb ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_srgbcolor
                                                         parent = lo_color ).
          lo_srgb->set_attribute( name = c_val value = dark2-srgb ).
        ELSE.
          lo_syscolor ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_syscolor
                                                          parent = lo_color ).
          lo_syscolor->set_attribute( name = c_val value = dark2-syscolor-val ).
          lo_syscolor->set_attribute( name = c_lastclr value = dark2-syscolor-lastclr ).
        ENDIF.
        CLEAR: lo_color, lo_srgb, lo_syscolor.
      ENDIF.

      lo_color ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_light2
                                                      parent = lo_scheme_element ).
      IF lo_color IS BOUND.
        IF light2-srgb IS NOT INITIAL.
          lo_srgb ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_srgbcolor
                                                         parent = lo_color ).
          lo_srgb->set_attribute( name = c_val value = light2-srgb ).
        ELSE.
          lo_syscolor ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_syscolor
                                                          parent = lo_color ).
          lo_syscolor->set_attribute( name = c_val value = light2-syscolor-val ).
          lo_syscolor->set_attribute( name = c_lastclr value = light2-syscolor-lastclr ).
        ENDIF.
        CLEAR: lo_color, lo_srgb, lo_syscolor.
      ENDIF.


      lo_color ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_accent1
                                                      parent = lo_scheme_element ).
      IF lo_color IS BOUND.
        IF accent1-srgb IS NOT INITIAL.
          lo_srgb ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_srgbcolor
                                                         parent = lo_color ).
          lo_srgb->set_attribute( name = c_val value = accent1-srgb ).
        ELSE.
          lo_syscolor ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_syscolor
                                                          parent = lo_color ).
          lo_syscolor->set_attribute( name = c_val value = accent1-syscolor-val ).
          lo_syscolor->set_attribute( name = c_lastclr value = accent1-syscolor-lastclr ).
        ENDIF.
        CLEAR: lo_color, lo_srgb, lo_syscolor.
      ENDIF.


      lo_color ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_accent2
                                                      parent = lo_scheme_element ).
      IF lo_color IS BOUND.
        IF accent2-srgb IS NOT INITIAL.
          lo_srgb ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_srgbcolor
                                                         parent = lo_color ).
          lo_srgb->set_attribute( name = c_val value = accent2-srgb ).
        ELSE.
          lo_syscolor ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_syscolor
                                                          parent = lo_color ).
          lo_syscolor->set_attribute( name = c_val value = accent2-syscolor-val ).
          lo_syscolor->set_attribute( name = c_lastclr value = accent2-syscolor-lastclr ).
        ENDIF.
        CLEAR: lo_color, lo_srgb, lo_syscolor.
      ENDIF.


      lo_color ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_accent3
                                                      parent = lo_scheme_element ).
      IF lo_color IS BOUND.
        IF accent3-srgb IS NOT INITIAL.
          lo_srgb ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_srgbcolor
                                                         parent = lo_color ).
          lo_srgb->set_attribute( name = c_val value = accent3-srgb ).
        ELSE.
          lo_syscolor ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_syscolor
                                                          parent = lo_color ).
          lo_syscolor->set_attribute( name = c_val value = accent3-syscolor-val ).
          lo_syscolor->set_attribute( name = c_lastclr value = accent3-syscolor-lastclr ).
        ENDIF.
        CLEAR: lo_color, lo_srgb, lo_syscolor.
      ENDIF.


      lo_color ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_accent4
                                                      parent = lo_scheme_element ).
      IF lo_color IS BOUND.
        IF accent4-srgb IS NOT INITIAL.
          lo_srgb ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_srgbcolor
                                                         parent = lo_color ).
          lo_srgb->set_attribute( name = c_val value = accent4-srgb ).
        ELSE.
          lo_syscolor ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_syscolor
                                                          parent = lo_color ).
          lo_syscolor->set_attribute( name = c_val value = accent4-syscolor-val ).
          lo_syscolor->set_attribute( name = c_lastclr value = accent4-syscolor-lastclr ).
        ENDIF.
        CLEAR: lo_color, lo_srgb, lo_syscolor.
      ENDIF.


      lo_color ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_accent5
                                                      parent = lo_scheme_element ).
      IF lo_color IS BOUND.
        IF accent5-srgb IS NOT INITIAL.
          lo_srgb ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_srgbcolor
                                                         parent = lo_color ).
          lo_srgb->set_attribute( name = c_val value = accent5-srgb ).
        ELSE.
          lo_syscolor ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_syscolor
                                                          parent = lo_color ).
          lo_syscolor->set_attribute( name = c_val value = accent5-syscolor-val ).
          lo_syscolor->set_attribute( name = c_lastclr value = accent5-syscolor-lastclr ).
        ENDIF.
        CLEAR: lo_color, lo_srgb, lo_syscolor.
      ENDIF.


      lo_color ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_accent6
                                                      parent = lo_scheme_element ).
      IF lo_color IS BOUND.
        IF accent6-srgb IS NOT INITIAL.
          lo_srgb ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_srgbcolor
                                                         parent = lo_color ).
          lo_srgb->set_attribute( name = c_val value = accent6-srgb ).
        ELSE.
          lo_syscolor ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_syscolor
                                                          parent = lo_color ).
          lo_syscolor->set_attribute( name = c_val value = accent6-syscolor-val ).
          lo_syscolor->set_attribute( name = c_lastclr value = accent6-syscolor-lastclr ).
        ENDIF.
        CLEAR: lo_color, lo_srgb, lo_syscolor.
      ENDIF.

      lo_color ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_hlink
                                                      parent = lo_scheme_element ).
      IF lo_color IS BOUND.
        IF hlink-srgb IS NOT INITIAL.
          lo_srgb ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_srgbcolor
                                                         parent = lo_color ).
          lo_srgb->set_attribute( name = c_val value = hlink-srgb ).
        ELSE.
          lo_syscolor ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_syscolor
                                                          parent = lo_color ).
          lo_syscolor->set_attribute( name = c_val value = hlink-syscolor-val ).
          lo_syscolor->set_attribute( name = c_lastclr value = hlink-syscolor-lastclr ).
        ENDIF.
        CLEAR: lo_color, lo_srgb, lo_syscolor.
      ENDIF.

      lo_color ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_folhlink
                                                      parent = lo_scheme_element ).
      IF lo_color IS BOUND.
        IF folhlink-srgb IS NOT INITIAL.
          lo_srgb ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_srgbcolor
                                                         parent = lo_color ).
          lo_srgb->set_attribute( name = c_val value = folhlink-srgb ).
        ELSE.
          lo_syscolor ?= io_document->create_simple_element_ns( prefix = zcl_excel_theme=>c_theme_prefix  name   = c_syscolor
                                                          parent = lo_color ).
          lo_syscolor->set_attribute( name = c_val value = folhlink-syscolor-val ).
          lo_syscolor->set_attribute( name = c_lastclr value = folhlink-syscolor-lastclr ).
        ENDIF.
      ENDIF.


    ENDIF.
  ENDMETHOD.                    "build_xml