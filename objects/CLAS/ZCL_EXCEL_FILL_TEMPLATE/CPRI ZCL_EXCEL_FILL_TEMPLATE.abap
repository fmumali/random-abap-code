  PRIVATE SECTION.

    TYPES: tt_cell_data_no_key TYPE STANDARD TABLE OF zexcel_s_cell_data WITH DEFAULT KEY.

    METHODS fill_range
      IMPORTING
        !iv_sheet              TYPE zexcel_sheet_title
        !iv_parent             TYPE zexcel_cell_row
        !iv_data               TYPE data
        VALUE(iv_range_length) TYPE zexcel_cell_row
        !io_sheet              TYPE REF TO zcl_excel_worksheet
      CHANGING
        !ct_cells              TYPE tt_cell_data_no_key
        !cv_diff               TYPE zexcel_cell_row
        !ct_merged_cells       TYPE zcl_excel_worksheet=>mty_ts_merge
      RAISING
        zcx_excel .
    METHODS get_range .
    METHODS validate_range
      IMPORTING
        !io_range TYPE REF TO zcl_excel_range .
    METHODS discard_overlapped .
    METHODS sign_range .
    METHODS find_var
      RAISING
        zcx_excel .
