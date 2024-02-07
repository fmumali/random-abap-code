  PROTECTED SECTION.

*"* protected components of class ZCL_EXCEL_SHEET_SETUP
*"* do not include other source files here!!!
    METHODS process_header_footer
      IMPORTING
        !ip_header                 TYPE zexcel_s_worksheet_head_foot
        !ip_side                   TYPE string
      RETURNING
        VALUE(rv_processed_string) TYPE string .