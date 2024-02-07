  METHOD get_rows.

    DATA: row TYPE i.
    FIELD-SYMBOLS: <sheet_cell> TYPE zexcel_s_cell_data.

    IF sheet_content IS NOT INITIAL.

      row = 0.
      DO.
        " Find the next row
        READ TABLE sheet_content ASSIGNING <sheet_cell> WITH KEY cell_row = row.
        CASE sy-subrc.
          WHEN 4.
            " row doesn't exist, but it exists another row, SY-TABIX points to the first cell in this row.
            READ TABLE sheet_content ASSIGNING <sheet_cell> INDEX sy-tabix.
            ASSERT sy-subrc = 0.
            row = <sheet_cell>-cell_row.
          WHEN 8.
            " it was the last available row
            EXIT.
        ENDCASE.
        " This will create the row instance if it doesn't exist
        get_row( row ).
        row = row + 1.
      ENDDO.

    ENDIF.

    eo_rows = me->rows.
  ENDMETHOD.                    "GET_ROWS