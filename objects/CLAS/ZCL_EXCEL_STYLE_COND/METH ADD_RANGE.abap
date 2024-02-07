  METHOD add_range.
    DATA: lv_column    TYPE zexcel_cell_column,
          lv_row_alpha TYPE string,
          lv_col_alpha TYPE string,
          lv_coords1   TYPE string,
          lv_coords2   TYPE string.


    lv_column = zcl_excel_common=>convert_column2int( ip_start_column ).

    lv_col_alpha = ip_start_column.
    lv_row_alpha = ip_start_row.
    SHIFT lv_row_alpha RIGHT DELETING TRAILING space.
    SHIFT lv_row_alpha LEFT DELETING LEADING space.
    CONCATENATE lv_col_alpha lv_row_alpha INTO lv_coords1.

    IF ip_stop_column IS NOT INITIAL.
      lv_column = zcl_excel_common=>convert_column2int( ip_stop_column ).
    ELSE.
      lv_column = zcl_excel_common=>convert_column2int( ip_start_column ).
    ENDIF.

    IF ip_stop_row IS NOT INITIAL. " If we don't get explicitly a stop column use start column
      lv_row_alpha = ip_stop_row.
    ELSE.
      lv_row_alpha = ip_start_row.
    ENDIF.
    IF ip_stop_column IS NOT INITIAL. " If we don't get explicitly a stop column use start column
      lv_col_alpha = ip_stop_column.
    ELSE.
      lv_col_alpha = ip_start_column.
    ENDIF.
    SHIFT lv_row_alpha RIGHT DELETING TRAILING space.
    SHIFT lv_row_alpha LEFT DELETING LEADING space.
    CONCATENATE lv_col_alpha lv_row_alpha INTO lv_coords2.
    IF lv_coords2 IS NOT INITIAL AND lv_coords2 <> lv_coords1.
      CONCATENATE me->mv_rule_range ` ` lv_coords1 ':' lv_coords2 INTO me->mv_rule_range.
    ELSE.
      CONCATENATE me->mv_rule_range ` ` lv_coords1  INTO me->mv_rule_range.
    ENDIF.
    SHIFT me->mv_rule_range LEFT DELETING LEADING space.

  ENDMETHOD.