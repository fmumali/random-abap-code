  METHOD create_formular_total.
    DATA: l_row_alpha   TYPE string,
          l_row_e_alpha TYPE string.

    l_row_alpha   = w_row_int + 1.
    l_row_e_alpha = i_row_int.

    CONCATENATE i_totals_function '(' i_column l_row_alpha ':' i_column l_row_e_alpha ')' INTO r_formula.
  ENDMETHOD.