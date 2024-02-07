  METHOD create_style_hdr.
    DATA: lo_style TYPE REF TO zcl_excel_style.

    lo_style                       = wo_excel->add_new_style( ).
    lo_style->font->bold           = abap_true.
    lo_style->font->color-rgb      = zcl_excel_style_color=>c_white.
    lo_style->fill->filltype       = zcl_excel_style_fill=>c_fill_solid.
    lo_style->fill->fgcolor-rgb    = 'FF4F81BD'.
    IF i_alignment IS SUPPLIED AND i_alignment IS NOT INITIAL.
      lo_style->alignment->horizontal = i_alignment.
    ENDIF.
    ro_style = lo_style .
  ENDMETHOD.