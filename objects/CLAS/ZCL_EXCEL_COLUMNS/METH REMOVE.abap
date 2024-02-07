  METHOD remove.
    DELETE TABLE columns_hashed WITH TABLE KEY column_index = io_column->get_column_index( ) .
    columns->remove( io_column ).
  ENDMETHOD.