  METHOD convert_columnrow2column_a_row.
*--------------------------------------------------------------------*
    "issue #256 - replacing char processing with regex
*--------------------------------------------------------------------*
* Stefan Schmoecker, 2013-08-11
*    Allow input to be CLIKE instead of STRING
*--------------------------------------------------------------------*

    DATA: pane_cell_row_a TYPE string,
          lv_columnrow    TYPE string.

    lv_columnrow = i_columnrow.    " Get rid of trailing blanks

    FIND REGEX '^(\D+)(\d+)$' IN lv_columnrow SUBMATCHES e_column
                                                         pane_cell_row_a.
    IF e_column_int IS SUPPLIED.
      e_column_int = convert_column2int( ip_column = e_column ).
    ENDIF.
    e_row = pane_cell_row_a.

  ENDMETHOD.