  METHOD set_merge.

    DATA: ls_merge        TYPE mty_merge,
          lv_column_start TYPE zexcel_cell_column,
          lv_column_end   TYPE zexcel_cell_column,
          lv_row          TYPE zexcel_cell_row,
          lv_row_to       TYPE zexcel_cell_row,
          lv_errormessage TYPE string.

    normalize_range_parameter( EXPORTING ip_range        = ip_range
                                         ip_column_start = ip_column_start ip_column_end = ip_column_end
                                         ip_row          = ip_row          ip_row_to     = ip_row_to
                               IMPORTING ep_column_start = lv_column_start ep_column_end = lv_column_end
                                         ep_row          = lv_row          ep_row_to     = lv_row_to ).

    IF ip_value IS SUPPLIED OR ip_formula IS SUPPLIED.
      " if there is a value or formula set the value to the top-left cell
      "maybe it is necessary to support other paramters for set_cell
      IF ip_value IS SUPPLIED.
        me->set_cell( ip_row = lv_row ip_column = lv_column_start
                      ip_value = ip_value ).
      ENDIF.
      IF ip_formula IS SUPPLIED.
        me->set_cell( ip_row = lv_row ip_column = lv_column_start
                      ip_value = ip_formula ).
      ENDIF.
    ENDIF.
    "call to set_merge_style to apply the style to all cells at the matrix
    IF ip_style IS SUPPLIED.
      me->set_merge_style( ip_row = lv_row ip_column_start = lv_column_start
                           ip_row_to = lv_row_to ip_column_end = lv_column_end
                           ip_style = ip_style ).
    ENDIF.
    ...
*--------------------------------------------------------------------*
* Build new range area to insert into range table
*--------------------------------------------------------------------*
    ls_merge-row_from = lv_row.
    ls_merge-row_to   = lv_row_to.
    ls_merge-col_from = lv_column_start.
    ls_merge-col_to   = lv_column_end.

*--------------------------------------------------------------------*
* Check merge not overlapping with existing merges
*--------------------------------------------------------------------*
    LOOP AT me->mt_merged_cells TRANSPORTING NO FIELDS WHERE NOT (    row_from > ls_merge-row_to
                                                                   OR row_to   < ls_merge-row_from
                                                                   OR col_from > ls_merge-col_to
                                                                   OR col_to   < ls_merge-col_from ).
      lv_errormessage = 'Overlapping merges'(404).
      zcx_excel=>raise_text( lv_errormessage ).

    ENDLOOP.

*--------------------------------------------------------------------*
* Everything seems ok --> add to merge table
*--------------------------------------------------------------------*
    INSERT ls_merge INTO TABLE me->mt_merged_cells.

  ENDMETHOD.                    "SET_MERGE