  METHOD set_cell_style.

    DATA: lv_column     TYPE zexcel_cell_column,
          lv_row        TYPE zexcel_cell_row,
          lv_style_guid TYPE zexcel_cell_style.

    FIELD-SYMBOLS: <fs_sheet_content> TYPE zexcel_s_cell_data.

    lv_style_guid = normalize_style_parameter( ip_style ).

    normalize_columnrow_parameter( EXPORTING ip_columnrow = ip_columnrow
                                             ip_column    = ip_column
                                             ip_row       = ip_row
                                   IMPORTING ep_column    = lv_column
                                             ep_row       = lv_row ).

    READ TABLE sheet_content ASSIGNING <fs_sheet_content> WITH KEY cell_row    = lv_row
                                                                   cell_column = lv_column.

    IF sy-subrc EQ 0.
      <fs_sheet_content>-cell_style   = lv_style_guid.
    ELSE.
      set_cell( ip_column = ip_column ip_row = ip_row ip_value = '' ip_style = ip_style ).
    ENDIF.

  ENDMETHOD.                    "SET_CELL_STYLE