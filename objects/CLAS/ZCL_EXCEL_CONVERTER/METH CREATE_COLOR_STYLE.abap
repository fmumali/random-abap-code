  METHOD create_color_style.
    DATA: ls_styles TYPE ts_styles.
    DATA: lo_style TYPE REF TO zcl_excel_style.

    READ TABLE wt_styles INTO ls_styles WITH KEY guid = i_style.
    IF sy-subrc = 0.
      lo_style                 = wo_excel->add_new_style( ).
      lo_style->font->bold                 = ls_styles-style->font->bold.
      lo_style->alignment->horizontal      = ls_styles-style->alignment->horizontal.
      lo_style->number_format->format_code = ls_styles-style->number_format->format_code.

      lo_style->font->color-rgb      = is_colors-fontcolor.
      lo_style->fill->filltype       = zcl_excel_style_fill=>c_fill_solid.
      lo_style->fill->fgcolor-rgb    = is_colors-fillcolor.

      ro_style = lo_style.
    ENDIF.
  ENDMETHOD.