CLASS zcl_excel_converter DEFINITION
  PUBLIC
  CREATE PUBLIC .

*"* public components of class ZCL_EXCEL_CONVERTER
*"* do not include other source files here!!!
  PUBLIC SECTION.

    CLASS-METHODS class_constructor .
    METHODS ask_option
      RETURNING
        VALUE(rs_option) TYPE zexcel_s_converter_option
      RAISING
        zcx_excel .
    METHODS convert
      IMPORTING
        !is_option     TYPE zexcel_s_converter_option OPTIONAL
        !io_alv        TYPE REF TO object OPTIONAL
        !it_table      TYPE STANDARD TABLE OPTIONAL
        !i_row_int     TYPE i DEFAULT 1
        !i_column_int  TYPE i DEFAULT 1
        !i_table       TYPE flag OPTIONAL
        !i_style_table TYPE zexcel_table_style OPTIONAL
        !io_worksheet  TYPE REF TO zcl_excel_worksheet OPTIONAL
      CHANGING
        !co_excel      TYPE REF TO zcl_excel OPTIONAL
      RAISING
        zcx_excel .
    METHODS create_path
      RETURNING
        VALUE(r_path) TYPE string .
    METHODS get_file
      EXPORTING
        !e_bytecount TYPE i
        !et_file     TYPE solix_tab
        !e_file      TYPE xstring
      RAISING
        zcx_excel .
    METHODS get_option
      RETURNING
        VALUE(rs_option) TYPE zexcel_s_converter_option .
    METHODS open_file
      RAISING
        zcx_excel .
    METHODS set_option
      IMPORTING
        !is_option TYPE zexcel_s_converter_option .
    METHODS write_file
      IMPORTING
        !i_path TYPE string OPTIONAL
      RAISING
        zcx_excel .
*"* protected components of class ZCL_EXCEL_CONVERTER
*"* do not include other source files here!!!