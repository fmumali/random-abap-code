  METHOD add.
    DATA: ls_hashed_row TYPE mty_s_hashed_row.

    ls_hashed_row-row_index = io_row->get_row_index( ).
    ls_hashed_row-row = io_row.

    INSERT ls_hashed_row INTO TABLE rows_hashed.

    rows->add( io_row ).
  ENDMETHOD.                    "ADD