  PRIVATE SECTION.
    TYPES:
      BEGIN OF mty_s_hashed_column,
        column_index TYPE int4,
        column       TYPE REF TO zcl_excel_column,
      END OF mty_s_hashed_column ,
      mty_ts_hashed_column TYPE HASHED TABLE OF mty_s_hashed_column WITH UNIQUE KEY column_index.

    DATA columns TYPE REF TO zcl_excel_collection .
    DATA columns_hashed TYPE mty_ts_hashed_column .