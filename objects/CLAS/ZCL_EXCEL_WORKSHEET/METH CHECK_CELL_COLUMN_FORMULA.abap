  METHOD check_cell_column_formula.

    FIELD-SYMBOLS <fs_column_formula> TYPE zcl_excel_worksheet=>mty_s_column_formula.

    IF ip_value IS NOT INITIAL OR ip_formula IS NOT INITIAL.
      zcx_excel=>raise_text( c_messages-formula_id_only_is_possible ).
    ENDIF.
    READ TABLE it_column_formulas WITH TABLE KEY id = ip_column_formula_id ASSIGNING <fs_column_formula>.
    IF sy-subrc <> 0.
      zcx_excel=>raise_text( c_messages-column_formula_id_not_found ).
    ENDIF.
    IF ip_row < <fs_column_formula>-table_top_left_row + 1
          OR ip_row > <fs_column_formula>-table_bottom_right_row + 1
          OR ip_column < <fs_column_formula>-table_left_column_int
          OR ip_column > <fs_column_formula>-table_right_column_int.
      zcx_excel=>raise_text( c_messages-formula_not_in_this_table ).
    ENDIF.
    IF ip_column <> <fs_column_formula>-column.
      zcx_excel=>raise_text( c_messages-formula_in_other_column ).
    ENDIF.

  ENDMETHOD.