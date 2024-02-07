  METHOD delete_merge.

    DATA: lv_column TYPE i.
*--------------------------------------------------------------------*
* If cell information is passed delete merge including this cell,
* otherwise delete all merges
*--------------------------------------------------------------------*
    IF   ip_cell_column IS INITIAL
      OR ip_cell_row    IS INITIAL.
      CLEAR me->mt_merged_cells.
    ELSE.
      lv_column = zcl_excel_common=>convert_column2int( ip_cell_column ).

      LOOP AT me->mt_merged_cells TRANSPORTING NO FIELDS
      WHERE row_from <= ip_cell_row AND row_to >= ip_cell_row
        AND col_from <= lv_column AND col_to >= lv_column.
        DELETE me->mt_merged_cells.
        EXIT.
      ENDLOOP.
    ENDIF.

  ENDMETHOD.                    "DELETE_MERGE