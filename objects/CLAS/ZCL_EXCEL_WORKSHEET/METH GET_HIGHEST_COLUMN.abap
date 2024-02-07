  METHOD get_highest_column.
    me->update_dimension_range( ).
    r_highest_column = me->lower_cell-cell_column.
  ENDMETHOD.                    "GET_HIGHEST_COLUMN