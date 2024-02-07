  METHOD get_dimension_range.

    me->update_dimension_range( ).
    IF upper_cell EQ lower_cell. "only one cell
      " Worksheet not filled
      IF upper_cell-cell_coords IS INITIAL.
        ep_dimension_range = 'A1'.
      ELSE.
        ep_dimension_range = upper_cell-cell_coords.
      ENDIF.
    ELSE.
      CONCATENATE upper_cell-cell_coords ':' lower_cell-cell_coords INTO ep_dimension_range.
    ENDIF.

  ENDMETHOD.                    "GET_DIMENSION_RANGE