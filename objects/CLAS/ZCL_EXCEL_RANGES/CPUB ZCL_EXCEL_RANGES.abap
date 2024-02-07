CLASS zcl_excel_ranges DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

*"* public components of class ZCL_EXCEL_RANGES
*"* do not include other source files here!!!
  PUBLIC SECTION.

    METHODS add
      IMPORTING
        !ip_range TYPE REF TO zcl_excel_range .
    METHODS clear .
    METHODS constructor .
    METHODS get
      IMPORTING
        !ip_index       TYPE i
      RETURNING
        VALUE(eo_range) TYPE REF TO zcl_excel_range .
    METHODS get_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS is_empty
      RETURNING
        VALUE(is_empty) TYPE flag .
    METHODS remove
      IMPORTING
        !ip_range TYPE REF TO zcl_excel_range .
    METHODS size
      RETURNING
        VALUE(ep_size) TYPE i .
*"* protected components of class ZABAP_EXCEL_WORKSHEETS
*"* do not include other source files here!!!