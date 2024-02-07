CLASS zcl_excel_autofilters DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.
*"* public components of class ZCL_EXCEL_AUTOFILTERS
*"* do not include other source files here!!!

    CONSTANTS c_autofilter TYPE string VALUE '_xlnm._FilterDatabase'. "#EC NOTEXT

    METHODS add
      IMPORTING
        !io_sheet            TYPE REF TO zcl_excel_worksheet
      RETURNING
        VALUE(ro_autofilter) TYPE REF TO zcl_excel_autofilter
      RAISING
        zcx_excel .
    METHODS clear .
    METHODS get
      IMPORTING
        !io_worksheet        TYPE REF TO zcl_excel_worksheet
      RETURNING
        VALUE(ro_autofilter) TYPE REF TO zcl_excel_autofilter .
    METHODS is_empty
      RETURNING
        VALUE(r_empty) TYPE flag .
    METHODS remove
      IMPORTING
        !io_sheet TYPE any .
    METHODS size
      RETURNING
        VALUE(r_size) TYPE i .
*"* protected components of class ZABAP_EXCEL_WORKSHEETS
*"* do not include other source files here!!!
*"* protected components of class ZABAP_EXCEL_WORKSHEETS
*"* do not include other source files here!!!