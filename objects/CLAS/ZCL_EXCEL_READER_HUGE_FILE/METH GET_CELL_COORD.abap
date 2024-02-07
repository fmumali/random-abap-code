  METHOD get_cell_coord.

    zcl_excel_common=>convert_columnrow2column_a_row(
      EXPORTING
        i_columnrow = iv_coord
      IMPORTING
        e_column    = es_coord-column
        e_row       = es_coord-row
      ).

  ENDMETHOD.