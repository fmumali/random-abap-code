  METHOD get_highest_row.
    me->update_dimension_range( ).
    r_highest_row = me->lower_cell-cell_row.
  ENDMETHOD.                    "GET_HIGHEST_ROW