*----------------------------------------------------------------------*
*       CLASS ZCL_EXCEL_ROWS DEFINITION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
CLASS zcl_excel_rows DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

*"* public components of class ZCL_EXCEL_ROWS
*"* do not include other source files here!!!
*"* protected components of class ZABAP_EXCEL_WORKSHEETS
*"* do not include other source files here!!!
  PUBLIC SECTION.
    METHODS add
      IMPORTING
        !io_row TYPE REF TO zcl_excel_row .
    METHODS clear .
    METHODS constructor .
    METHODS get
      IMPORTING
        !ip_index     TYPE i
      RETURNING
        VALUE(eo_row) TYPE REF TO zcl_excel_row .
    METHODS get_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS is_empty
      RETURNING
        VALUE(is_empty) TYPE flag .
    METHODS remove
      IMPORTING
        !io_row TYPE REF TO zcl_excel_row .
    METHODS size
      RETURNING
        VALUE(ep_size) TYPE i .
    METHODS get_min_index
      RETURNING
        VALUE(ep_index) TYPE i .
    METHODS get_max_index
      RETURNING
        VALUE(ep_index) TYPE i .