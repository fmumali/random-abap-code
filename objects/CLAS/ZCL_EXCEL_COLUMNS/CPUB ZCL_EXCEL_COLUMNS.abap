CLASS zcl_excel_columns DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

*"* public components of class ZCL_EXCEL_COLUMNS
*"* do not include other source files here!!!
  PUBLIC SECTION.
    METHODS add
      IMPORTING
        !io_column TYPE REF TO zcl_excel_column .
    METHODS clear .
    METHODS constructor .
    METHODS get
      IMPORTING
        !ip_index        TYPE i
      RETURNING
        VALUE(eo_column) TYPE REF TO zcl_excel_column .
    METHODS get_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS is_empty
      RETURNING
        VALUE(is_empty) TYPE flag .
    METHODS remove
      IMPORTING
        !io_column TYPE REF TO zcl_excel_column .
    METHODS size
      RETURNING
        VALUE(ep_size) TYPE i .
*"* protected components of class ZABAP_EXCEL_WORKSHEETS
*"* do not include other source files here!!!