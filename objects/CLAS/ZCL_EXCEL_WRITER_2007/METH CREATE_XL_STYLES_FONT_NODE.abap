  METHOD create_xl_styles_font_node.

    CONSTANTS: lc_xml_node_b      TYPE string VALUE 'b',            "bold
               lc_xml_node_i      TYPE string VALUE 'i',            "italic
               lc_xml_node_u      TYPE string VALUE 'u',            "underline
               lc_xml_node_strike TYPE string VALUE 'strike',       "strikethrough
               lc_xml_node_sz     TYPE string VALUE 'sz',
               lc_xml_node_name   TYPE string VALUE 'name',
               lc_xml_node_rfont  TYPE string VALUE 'rFont',
               lc_xml_node_family TYPE string VALUE 'family',
               lc_xml_node_scheme TYPE string VALUE 'scheme',
               lc_xml_attr_val    TYPE string VALUE 'val'.

    DATA: lo_document     TYPE REF TO if_ixml_document,
          lo_element_font TYPE REF TO if_ixml_element,
          ls_font         TYPE zexcel_s_style_font,
          lo_sub_element  TYPE REF TO if_ixml_element,
          lv_value        TYPE string.

    lo_document = io_document.
    lo_element_font = io_parent.
    ls_font = is_font.

    IF ls_font-bold EQ abap_true.
      lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_b
                                                           parent = lo_document ).
      lo_element_font->append_child( new_child = lo_sub_element ).
    ENDIF.
    IF ls_font-italic EQ abap_true.
      lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_i
                                                           parent = lo_document ).
      lo_element_font->append_child( new_child = lo_sub_element ).
    ENDIF.
    IF ls_font-underline EQ abap_true.
      lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_u
                                                           parent = lo_document ).
      lv_value = ls_font-underline_mode.
      lo_sub_element->set_attribute_ns( name  = lc_xml_attr_val
                                        value = lv_value ).
      lo_element_font->append_child( new_child = lo_sub_element ).
    ENDIF.
    IF ls_font-strikethrough EQ abap_true.
      lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_strike
                                                           parent = lo_document ).
      lo_element_font->append_child( new_child = lo_sub_element ).
    ENDIF.
    "size
    lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_sz
                                                         parent = lo_document ).
    lv_value = ls_font-size.
    SHIFT lv_value RIGHT DELETING TRAILING space.
    SHIFT lv_value LEFT DELETING LEADING space.
    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_val
                                      value = lv_value ).
    lo_element_font->append_child( new_child = lo_sub_element ).
    "color
    create_xl_styles_color_node(
        io_document        = lo_document
        io_parent          = lo_element_font
        is_color           = ls_font-color ).

    "name
    IF iv_use_rtf = abap_false.
      lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_name
                                                           parent = lo_document ).
    ELSE.
      lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_rfont
                                                           parent = lo_document ).
    ENDIF.
    lv_value = ls_font-name.
    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_val
                                      value = lv_value ).
    lo_element_font->append_child( new_child = lo_sub_element ).
    "family
    lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_family
                                                         parent = lo_document ).
    lv_value = ls_font-family.
    SHIFT lv_value RIGHT DELETING TRAILING space.
    SHIFT lv_value LEFT DELETING LEADING space.
    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_val
                                      value = lv_value ).
    lo_element_font->append_child( new_child = lo_sub_element ).
    "scheme
    IF ls_font-scheme IS NOT INITIAL.
      lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_scheme
                                                           parent = lo_document ).
      lv_value = ls_font-scheme.
      lo_sub_element->set_attribute_ns( name  = lc_xml_attr_val
                                        value = lv_value ).
      lo_element_font->append_child( new_child = lo_sub_element ).
    ENDIF.

  ENDMETHOD.