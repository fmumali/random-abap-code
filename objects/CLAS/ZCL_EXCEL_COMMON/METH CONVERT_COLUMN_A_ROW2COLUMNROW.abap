  METHOD convert_column_a_row2columnrow.
    DATA: lv_row_alpha    TYPE string,
          lv_column_alpha TYPE zexcel_cell_column_alpha.

    lv_row_alpha = i_row.
    lv_column_alpha = zcl_excel_common=>convert_column2alpha( i_column ).
    SHIFT lv_row_alpha RIGHT DELETING TRAILING space.
    SHIFT lv_row_alpha LEFT DELETING LEADING space.
    CONCATENATE lv_column_alpha lv_row_alpha INTO e_columnrow.

  ENDMETHOD.