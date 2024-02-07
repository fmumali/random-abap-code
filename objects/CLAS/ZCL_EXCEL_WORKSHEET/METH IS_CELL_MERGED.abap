  METHOD is_cell_merged.

    DATA: lv_column TYPE i.

    FIELD-SYMBOLS: <ls_merged_cell> LIKE LINE OF me->mt_merged_cells.

    lv_column = zcl_excel_common=>convert_column2int( ip_column ).

    rp_is_merged = abap_false.                                        " Assume not in merged area

    LOOP AT me->mt_merged_cells ASSIGNING <ls_merged_cell>.

      IF    <ls_merged_cell>-col_from <= lv_column
        AND <ls_merged_cell>-col_to   >= lv_column
        AND <ls_merged_cell>-row_from <= ip_row
        AND <ls_merged_cell>-row_to   >= ip_row.
        rp_is_merged = abap_true.                                     " until we are proven different
        RETURN.
      ENDIF.

    ENDLOOP.

  ENDMETHOD.                    "IS_CELL_MERGED