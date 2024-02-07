  METHOD get_active_cell.

    DATA: lv_active_column TYPE zexcel_cell_column_alpha,
          lv_active_row    TYPE string.

    lv_active_column = zcl_excel_common=>convert_column2alpha( active_cell-cell_column ).
    lv_active_row    = active_cell-cell_row.
    SHIFT lv_active_row RIGHT DELETING TRAILING space.
    SHIFT lv_active_row LEFT DELETING LEADING space.
    CONCATENATE lv_active_column lv_active_row INTO ep_active_cell.

  ENDMETHOD.                    "GET_ACTIVE_CELL