  METHOD create_xl_styles_color_node.
    DATA: lo_sub_element TYPE REF TO if_ixml_element,
          lv_value       TYPE string.

    CONSTANTS: lc_xml_attr_theme   TYPE string VALUE 'theme',
               lc_xml_attr_rgb     TYPE string VALUE 'rgb',
               lc_xml_attr_indexed TYPE string VALUE 'indexed',
               lc_xml_attr_tint    TYPE string VALUE 'tint'.

    "add node only if at least one attribute is set
    CHECK is_color-rgb IS NOT INITIAL OR
          is_color-indexed <> zcl_excel_style_color=>c_indexed_not_set OR
          is_color-theme <> zcl_excel_style_color=>c_theme_not_set OR
          is_color-tint IS NOT INITIAL.

    lo_sub_element = io_document->create_simple_element(
        name      = iv_color_elem_name
        parent    = io_parent ).

    IF is_color-rgb IS NOT INITIAL.
      lv_value = is_color-rgb.
      lo_sub_element->set_attribute_ns( name  = lc_xml_attr_rgb
                                        value = lv_value ).
    ENDIF.

    IF is_color-indexed <> zcl_excel_style_color=>c_indexed_not_set.
      lv_value = zcl_excel_common=>number_to_excel_string( is_color-indexed ).
      lo_sub_element->set_attribute_ns( name  = lc_xml_attr_indexed
                                        value = lv_value ).
    ENDIF.

    IF is_color-theme <> zcl_excel_style_color=>c_theme_not_set.
      lv_value = zcl_excel_common=>number_to_excel_string( is_color-theme ).
      lo_sub_element->set_attribute_ns( name  = lc_xml_attr_theme
                                        value = lv_value ).
    ENDIF.

    IF is_color-tint IS NOT INITIAL.
      lv_value = zcl_excel_common=>number_to_excel_string( is_color-tint ).
      lo_sub_element->set_attribute_ns( name  = lc_xml_attr_tint
                                        value = lv_value ).
    ENDIF.

    io_parent->append_child( new_child = lo_sub_element ).
  ENDMETHOD.