  PRIVATE SECTION.
    TYPES:
      BEGIN OF mty_s_hashed_row,
        row_index TYPE int4,
        row       TYPE REF TO zcl_excel_row,
      END OF mty_s_hashed_row ,
      mty_ts_hashed_row TYPE HASHED TABLE OF mty_s_hashed_row WITH UNIQUE KEY row_index.

    DATA rows TYPE REF TO zcl_excel_collection .
    DATA rows_hashed TYPE mty_ts_hashed_row .
