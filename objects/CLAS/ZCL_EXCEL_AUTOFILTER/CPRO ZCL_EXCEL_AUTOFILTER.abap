  PROTECTED SECTION.

    METHODS get_column_filter
      IMPORTING
        !i_column        TYPE zexcel_cell_column
      RETURNING
        VALUE(rr_filter) TYPE REF TO ts_filter .
    METHODS is_row_hidden_single_values
      IMPORTING
        !iv_row             TYPE zexcel_cell_row
        !iv_col             TYPE zexcel_cell_column
        !is_filter          TYPE ts_filter
      RETURNING
        VALUE(rv_is_hidden) TYPE abap_bool .
    METHODS is_row_hidden_text_pattern
      IMPORTING
        !iv_row             TYPE zexcel_cell_row
        !iv_col             TYPE zexcel_cell_column
        !is_filter          TYPE ts_filter
      RETURNING
        VALUE(rv_is_hidden) TYPE abap_bool .
*"* private components of class ZCL_EXCEL_AUTOFILTER
*"* do not include other source files here!!!