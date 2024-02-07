  METHOD add.
    DATA: ls_hashed_column TYPE mty_s_hashed_column.

    ls_hashed_column-column_index = io_column->get_column_index( ).
    ls_hashed_column-column = io_column.

    INSERT ls_hashed_column INTO TABLE columns_hashed .

    columns->add( io_column ).
  ENDMETHOD.