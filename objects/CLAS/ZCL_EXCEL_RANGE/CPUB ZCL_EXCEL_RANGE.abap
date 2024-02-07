CLASS zcl_excel_range DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

*"* public components of class ZCL_EXCEL_RANGE
*"* do not include other source files here!!!
  PUBLIC SECTION.

    CONSTANTS gcv_print_title_name TYPE string VALUE '_xlnm.Print_Titles'. "#EC NOTEXT
    DATA name TYPE zexcel_range_name .
    DATA guid TYPE zexcel_range_guid .

    METHODS get_guid
      RETURNING
        VALUE(ep_guid) TYPE zexcel_range_guid .
    METHODS set_value
      IMPORTING
        !ip_sheet_name   TYPE zexcel_sheet_title
        !ip_start_row    TYPE zexcel_cell_row
        !ip_start_column TYPE zexcel_cell_column_alpha
        !ip_stop_row     TYPE zexcel_cell_row
        !ip_stop_column  TYPE zexcel_cell_column_alpha .
    METHODS get_value
      RETURNING
        VALUE(ep_value) TYPE zexcel_range_value .
    METHODS set_range_value
      IMPORTING
        !ip_value TYPE zexcel_range_value .
*"* protected components of class ZABAP_EXCEL_WORKSHEET
*"* do not include other source files here!!!