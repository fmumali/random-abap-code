  METHOD create_style_stripped.
    DATA:   lo_style    TYPE REF TO zcl_excel_style.
    DATA:  l_format    TYPE zexcel_number_format.

    lo_style                       = wo_excel->add_new_style( ).
    lo_style->fill->filltype       = zcl_excel_style_fill=>c_fill_solid.
    lo_style->fill->fgcolor-rgb    = 'FFDBE5F1'.
    IF i_alignment IS SUPPLIED AND i_alignment IS NOT INITIAL.
      lo_style->alignment->horizontal = i_alignment.
    ENDIF.
    IF i_inttype IS SUPPLIED AND i_inttype IS NOT INITIAL.
      l_format = set_cell_format(  i_inttype  = i_inttype
                                   i_decimals = i_decimals ) .
      IF l_format IS NOT INITIAL.
        lo_style->number_format->format_code = l_format.
      ENDIF.
    ENDIF.
    ro_style = lo_style.

  ENDMETHOD.