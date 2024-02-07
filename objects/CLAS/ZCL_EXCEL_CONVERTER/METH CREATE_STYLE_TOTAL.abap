  METHOD create_style_total.
    DATA:   lo_style    TYPE REF TO zcl_excel_style.
    DATA:  l_format    TYPE zexcel_number_format.

    lo_style                                   = wo_excel->add_new_style( ).
    lo_style->font->bold                       = abap_true.

    CREATE OBJECT lo_style->borders->top.
    lo_style->borders->top->border_style       = zcl_excel_style_border=>c_border_thin.
    lo_style->borders->top->border_color-rgb   = zcl_excel_style_color=>c_black.

    CREATE OBJECT lo_style->borders->right.
    lo_style->borders->right->border_style       = zcl_excel_style_border=>c_border_none.
    lo_style->borders->right->border_color-rgb   = zcl_excel_style_color=>c_black.

    CREATE OBJECT lo_style->borders->down.
    lo_style->borders->down->border_style      = zcl_excel_style_border=>c_border_double.
    lo_style->borders->down->border_color-rgb  = zcl_excel_style_color=>c_black.

    CREATE OBJECT lo_style->borders->left.
    lo_style->borders->left->border_style       = zcl_excel_style_border=>c_border_none.
    lo_style->borders->left->border_color-rgb   = zcl_excel_style_color=>c_black.

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

    ro_style = lo_style .

  ENDMETHOD.