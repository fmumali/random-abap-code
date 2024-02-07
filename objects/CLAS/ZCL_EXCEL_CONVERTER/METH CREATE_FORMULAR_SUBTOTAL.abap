  METHOD create_formular_subtotal.
    DATA: l_row_alpha_start TYPE string,
          l_row_alpha_end   TYPE string,
          l_func_num        TYPE string.

    l_row_alpha_start   = i_row_int_start.
    l_row_alpha_end     = i_row_int_end.

    l_func_num = get_function_number( i_totals_function = i_totals_function ).
    CONCATENATE 'SUBTOTAL(' l_func_num ',' i_column l_row_alpha_start ':' i_column l_row_alpha_end ')' INTO r_formula.
  ENDMETHOD.