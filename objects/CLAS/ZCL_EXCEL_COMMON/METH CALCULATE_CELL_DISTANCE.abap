  METHOD calculate_cell_distance.

    DATA: lv_reference_row       TYPE i,
          lv_reference_col_alpha TYPE zexcel_cell_column_alpha,
          lv_reference_col       TYPE i,
          lv_current_row         TYPE i,
          lv_current_col_alpha   TYPE zexcel_cell_column_alpha,
          lv_current_col         TYPE i.

*--------------------------------------------------------------------*
* Split reference  cell into numerical row/column representation
*--------------------------------------------------------------------*
    convert_columnrow2column_a_row( EXPORTING
                                      i_columnrow = iv_reference_cell
                                    IMPORTING
                                      e_column    = lv_reference_col_alpha
                                      e_row       = lv_reference_row ).
    lv_reference_col = convert_column2int( lv_reference_col_alpha ).

*--------------------------------------------------------------------*
* Split current  cell into numerical row/column representation
*--------------------------------------------------------------------*
    convert_columnrow2column_a_row( EXPORTING
                                      i_columnrow = iv_current_cell
                                    IMPORTING
                                      e_column    = lv_current_col_alpha
                                      e_row       = lv_current_row ).
    lv_current_col = convert_column2int( lv_current_col_alpha ).

*--------------------------------------------------------------------*
* Calculate row and column difference
* Positive:   Current cell below    reference cell
*         or  Current cell right of reference cell
* Negative:   Current cell above    reference cell
*         or  Current cell left  of reference cell
*--------------------------------------------------------------------*
    ev_row_difference = lv_current_row - lv_reference_row.
    ev_col_difference = lv_current_col - lv_reference_col.

  ENDMETHOD.