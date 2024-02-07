  METHOD load_style_font.

    DATA: lo_node_font TYPE REF TO if_ixml_element,
          lo_node2     TYPE REF TO if_ixml_element,
          lo_font      TYPE REF TO zcl_excel_style_font,
          ls_color     TYPE t_color.

    lo_node_font = io_xml_element.

    CREATE OBJECT lo_font.
*--------------------------------------------------------------------*
*   Bold
*--------------------------------------------------------------------*
    IF lo_node_font->find_from_name_ns( name = 'b' uri = namespace-main ) IS BOUND.
      lo_font->bold = abap_true.
    ENDIF.

*--------------------------------------------------------------------*
*   Italic
*--------------------------------------------------------------------*
    IF lo_node_font->find_from_name_ns( name = 'i' uri = namespace-main ) IS BOUND.
      lo_font->italic = abap_true.
    ENDIF.

*--------------------------------------------------------------------*
*   Underline
*--------------------------------------------------------------------*
    lo_node2 = lo_node_font->find_from_name_ns( name = 'u' uri = namespace-main ).
    IF lo_node2 IS BOUND.
      lo_font->underline      = abap_true.
      lo_font->underline_mode = lo_node2->get_attribute( 'val' ).
    ENDIF.

*--------------------------------------------------------------------*
*   StrikeThrough
*--------------------------------------------------------------------*
    IF lo_node_font->find_from_name_ns( name = 'strike' uri = namespace-main ) IS BOUND.
      lo_font->strikethrough = abap_true.
    ENDIF.

*--------------------------------------------------------------------*
*   Fontsize
*--------------------------------------------------------------------*
    lo_node2 = lo_node_font->find_from_name_ns( name = 'sz' uri = namespace-main ).
    IF lo_node2 IS BOUND.
      lo_font->size = lo_node2->get_attribute( 'val' ).
    ENDIF.

*--------------------------------------------------------------------*
*   Fontname
*--------------------------------------------------------------------*
    lo_node2 = lo_node_font->find_from_name_ns( name = 'name' uri = namespace-main ).
    IF lo_node2 IS BOUND.
      lo_font->name = lo_node2->get_attribute( 'val' ).
    ELSE.
      lo_node2 = lo_node_font->find_from_name_ns( name = 'rFont' uri = namespace-main ).
      IF lo_node2 IS BOUND.
        lo_font->name = lo_node2->get_attribute( 'val' ).
      ENDIF.
    ENDIF.

*--------------------------------------------------------------------*
*   Fontfamily
*--------------------------------------------------------------------*
    lo_node2 = lo_node_font->find_from_name_ns( name = 'family' uri = namespace-main ).
    IF lo_node2 IS BOUND.
      lo_font->family = lo_node2->get_attribute( 'val' ).
    ENDIF.

*--------------------------------------------------------------------*
*   Fontscheme
*--------------------------------------------------------------------*
    lo_node2 = lo_node_font->find_from_name_ns( name = 'scheme' uri = namespace-main ).
    IF lo_node2 IS BOUND.
      lo_font->scheme = lo_node2->get_attribute( 'val' ).
    ELSE.
      CLEAR lo_font->scheme.
    ENDIF.

*--------------------------------------------------------------------*
*   Fontcolor
*--------------------------------------------------------------------*
    lo_node2 = lo_node_font->find_from_name_ns( name = 'color' uri = namespace-main ).
    IF lo_node2 IS BOUND.
      fill_struct_from_attributes( EXPORTING
                                     ip_element   =  lo_node2
                                   CHANGING
                                     cp_structure = ls_color ).
      lo_font->color-rgb = ls_color-rgb.
      IF ls_color-indexed IS NOT INITIAL.
        lo_font->color-indexed = ls_color-indexed.
      ENDIF.

      IF ls_color-theme IS NOT INITIAL.
        lo_font->color-theme = ls_color-theme.
      ENDIF.
      lo_font->color-tint = ls_color-tint.
    ENDIF.

    ro_font = lo_font.

  ENDMETHOD.