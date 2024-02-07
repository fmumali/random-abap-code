  METHOD create_style_normal.
    DATA: lo_style TYPE REF TO zcl_excel_style,
          l_format TYPE zexcel_number_format.

    IF i_inttype IS SUPPLIED AND i_inttype IS NOT INITIAL.
      l_format = set_cell_format(  i_inttype  = i_inttype
                                   i_decimals = i_decimals ) .
    ENDIF.
    IF l_format IS NOT INITIAL OR
       ( i_alignment IS SUPPLIED AND i_alignment IS NOT INITIAL ) .

      lo_style                       = wo_excel->add_new_style( ).

      IF i_alignment IS SUPPLIED AND i_alignment IS NOT INITIAL.
        lo_style->alignment->horizontal = i_alignment.
      ENDIF.

      IF l_format IS NOT INITIAL.
        lo_style->number_format->format_code = l_format.
      ENDIF.

      ro_style = lo_style .

    ENDIF.
  ENDMETHOD.