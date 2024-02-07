CLASS zcl_excel_writer_csv DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

*"* public components of class ZCL_EXCEL_WRITER_CSV
*"* do not include other source files here!!!
  PUBLIC SECTION.

    INTERFACES zif_excel_writer .

    CLASS-METHODS set_delimiter
      IMPORTING
        VALUE(ip_value) TYPE c DEFAULT ';' .
    CLASS-METHODS set_enclosure
      IMPORTING
        VALUE(ip_value) TYPE c DEFAULT '"' .
    CLASS-METHODS set_endofline
      IMPORTING
        VALUE(ip_value) TYPE any DEFAULT cl_abap_char_utilities=>cr_lf .
    CLASS-METHODS set_active_sheet_index
      IMPORTING
        !i_active_worksheet TYPE zexcel_active_worksheet .
    CLASS-METHODS set_active_sheet_index_by_name
      IMPORTING
        !i_worksheet_name TYPE zexcel_worksheets_name .
*"* protected components of class ZCL_EXCEL_WRITER_2007
*"* do not include other source files here!!!