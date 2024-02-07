  METHOD convert_columnrow2column_o_row.

    DATA: row       TYPE string.
    DATA: columnrow TYPE string.

    CLEAR e_column.

    columnrow = i_columnrow.

    FIND REGEX '^(\D*)(\d*)$' IN columnrow SUBMATCHES e_column
                                                      row.

    e_row = row.

  ENDMETHOD.