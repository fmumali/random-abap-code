  METHOD get_merge.

    FIELD-SYMBOLS: <ls_merged_cell> LIKE LINE OF me->mt_merged_cells.

    DATA: lv_col_from    TYPE string,
          lv_col_to      TYPE string,
          lv_row_from    TYPE string,
          lv_row_to      TYPE string,
          lv_merge_range TYPE string.

    LOOP AT me->mt_merged_cells ASSIGNING <ls_merged_cell>.

      lv_col_from = zcl_excel_common=>convert_column2alpha( <ls_merged_cell>-col_from ).
      lv_col_to   = zcl_excel_common=>convert_column2alpha( <ls_merged_cell>-col_to   ).
      lv_row_from = <ls_merged_cell>-row_from.
      lv_row_to   = <ls_merged_cell>-row_to  .
      CONCATENATE lv_col_from lv_row_from ':' lv_col_to lv_row_to
         INTO lv_merge_range.
      CONDENSE lv_merge_range NO-GAPS.
      APPEND lv_merge_range TO merge_range.

    ENDLOOP.

  ENDMETHOD.                    "GET_MERGE