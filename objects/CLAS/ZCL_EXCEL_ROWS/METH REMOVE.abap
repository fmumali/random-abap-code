  METHOD remove.
    DELETE TABLE rows_hashed WITH TABLE KEY row_index = io_row->get_row_index( ) .
    rows->remove( io_row ).
  ENDMETHOD.                    "REMOVE