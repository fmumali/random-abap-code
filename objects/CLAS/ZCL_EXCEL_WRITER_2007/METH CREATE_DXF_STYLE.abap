  METHOD create_dxf_style.

    CONSTANTS: lc_xml_node_dxf         TYPE string VALUE 'dxf',
               lc_xml_node_font        TYPE string VALUE 'font',
               lc_xml_node_b           TYPE string VALUE 'b',            "bold
               lc_xml_node_i           TYPE string VALUE 'i',            "italic
               lc_xml_node_u           TYPE string VALUE 'u',            "underline
               lc_xml_node_strike      TYPE string VALUE 'strike',       "strikethrough
               lc_xml_attr_val         TYPE string VALUE 'val',
               lc_xml_node_fill        TYPE string VALUE 'fill',
               lc_xml_node_patternfill TYPE string VALUE 'patternFill',
               lc_xml_attr_patterntype TYPE string VALUE 'patternType',
               lc_xml_node_fgcolor     TYPE string VALUE 'fgColor',
               lc_xml_node_bgcolor     TYPE string VALUE 'bgColor'.

    DATA: ls_styles_mapping     TYPE zexcel_s_styles_mapping,
          ls_cellxfs            TYPE zexcel_s_cellxfs,
          ls_style_cond_mapping TYPE zexcel_s_styles_cond_mapping,
          lo_sub_element        TYPE REF TO if_ixml_element,
          lo_sub_element_2      TYPE REF TO if_ixml_element,
          lv_index              TYPE i,
          ls_font               TYPE zexcel_s_style_font,
          lo_element_font       TYPE REF TO if_ixml_element,
          lv_value              TYPE string,
          ls_fill               TYPE zexcel_s_style_fill,
          lo_element_fill       TYPE REF TO if_ixml_element.

    CHECK iv_cell_style IS NOT INITIAL.

    READ TABLE me->styles_mapping INTO ls_styles_mapping WITH KEY guid = iv_cell_style.
    ADD 1 TO ls_styles_mapping-style. " the numbering starts from 0
    READ TABLE it_cellxfs INTO ls_cellxfs INDEX ls_styles_mapping-style.
    ADD 1 TO ls_cellxfs-fillid.       " the numbering starts from 0

    READ TABLE me->styles_cond_mapping INTO ls_style_cond_mapping WITH KEY style = ls_styles_mapping-style.
    IF sy-subrc EQ 0.
      ls_style_cond_mapping-guid  = iv_cell_style.
      APPEND ls_style_cond_mapping TO me->styles_cond_mapping.
    ELSE.
      ls_style_cond_mapping-guid  = iv_cell_style.
      ls_style_cond_mapping-style = ls_styles_mapping-style.
      ls_style_cond_mapping-dxf   = cv_dfx_count.
      APPEND ls_style_cond_mapping TO me->styles_cond_mapping.
      ADD 1 TO cv_dfx_count.

      " dxf node
      lo_sub_element = io_ixml_document->create_simple_element( name   = lc_xml_node_dxf
                                                                parent = io_ixml_document ).

      "Conditional formatting font style correction by Alessandro Iannacci START
      lv_index = ls_cellxfs-fontid + 1.
      READ TABLE it_fonts INTO ls_font INDEX lv_index.
      IF ls_font IS NOT INITIAL.
        lo_element_font = io_ixml_document->create_simple_element( name   = lc_xml_node_font
                                                              parent = io_ixml_document ).
        IF ls_font-bold EQ abap_true.
          lo_sub_element_2 = io_ixml_document->create_simple_element( name   = lc_xml_node_b
                                                               parent = io_ixml_document ).
          lo_element_font->append_child( new_child = lo_sub_element_2 ).
        ENDIF.
        IF ls_font-italic EQ abap_true.
          lo_sub_element_2 = io_ixml_document->create_simple_element( name   = lc_xml_node_i
                                                               parent = io_ixml_document ).
          lo_element_font->append_child( new_child = lo_sub_element_2 ).
        ENDIF.
        IF ls_font-underline EQ abap_true.
          lo_sub_element_2 = io_ixml_document->create_simple_element( name   = lc_xml_node_u
                                                               parent = io_ixml_document ).
          lv_value = ls_font-underline_mode.
          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_val
                                            value = lv_value ).
          lo_element_font->append_child( new_child = lo_sub_element_2 ).
        ENDIF.
        IF ls_font-strikethrough EQ abap_true.
          lo_sub_element_2 = io_ixml_document->create_simple_element( name   = lc_xml_node_strike
                                                               parent = io_ixml_document ).
          lo_element_font->append_child( new_child = lo_sub_element_2 ).
        ENDIF.
        "color
        create_xl_styles_color_node(
            io_document        = io_ixml_document
            io_parent          = lo_element_font
            is_color           = ls_font-color ).
        lo_sub_element->append_child( new_child = lo_element_font ).
      ENDIF.
      "---Conditional formatting font style correction by Alessandro Iannacci END


      READ TABLE it_fills INTO ls_fill INDEX ls_cellxfs-fillid.
      IF ls_fill IS NOT INITIAL.
        " fill properties
        lo_element_fill = io_ixml_document->create_simple_element( name   = lc_xml_node_fill
                                                                 parent = io_ixml_document ).
        "pattern
        lo_sub_element_2 = io_ixml_document->create_simple_element( name   = lc_xml_node_patternfill
                                                             parent = io_ixml_document ).
        lv_value = ls_fill-filltype.
        lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_patterntype
                                            value = lv_value ).
        " fgcolor
        create_xl_styles_color_node(
            io_document        = io_ixml_document
            io_parent          = lo_sub_element_2
            is_color           = ls_fill-fgcolor
            iv_color_elem_name = lc_xml_node_fgcolor ).

        IF  ls_fill-fgcolor-rgb IS INITIAL AND
          ls_fill-fgcolor-indexed EQ zcl_excel_style_color=>c_indexed_not_set AND
          ls_fill-fgcolor-theme EQ zcl_excel_style_color=>c_theme_not_set AND
          ls_fill-fgcolor-tint IS INITIAL AND ls_fill-bgcolor-indexed EQ zcl_excel_style_color=>c_indexed_sys_foreground.

          " bgcolor
          create_xl_styles_color_node(
              io_document        = io_ixml_document
              io_parent          = lo_sub_element_2
              is_color           = ls_fill-bgcolor
              iv_color_elem_name = lc_xml_node_bgcolor ).

        ENDIF.

        lo_element_fill->append_child( new_child = lo_sub_element_2 ). "pattern

        lo_sub_element->append_child( new_child = lo_element_fill ).
      ENDIF.
    ENDIF.

    io_dxf_element->append_child( new_child = lo_sub_element ).
  ENDMETHOD.