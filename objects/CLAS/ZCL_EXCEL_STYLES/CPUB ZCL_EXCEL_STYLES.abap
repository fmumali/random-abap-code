CLASS zcl_excel_styles DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

*"* public components of class ZCL_EXCEL_STYLES
*"* do not include other source files here!!!
  PUBLIC SECTION.

    METHODS add
      IMPORTING
        !ip_style TYPE REF TO zcl_excel_style .
    METHODS clear .
    METHODS constructor .
    METHODS get
      IMPORTING
        !ip_index       TYPE i
      RETURNING
        VALUE(eo_style) TYPE REF TO zcl_excel_style .
    METHODS get_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS is_empty
      RETURNING
        VALUE(is_empty) TYPE flag .
    METHODS remove
      IMPORTING
        !ip_style TYPE REF TO zcl_excel_style .
    METHODS size
      RETURNING
        VALUE(ep_size) TYPE i .
    METHODS register_new_style
      IMPORTING
        !io_style            TYPE REF TO zcl_excel_style
      RETURNING
        VALUE(ep_style_code) TYPE i .
*"* protected components of class ZABAP_EXCEL_WORKSHEETS
*"* do not include other source files here!!!
*"* protected components of class ZABAP_EXCEL_WORKSHEETS
*"* do not include other source files here!!!
*"* protected components of class ZABAP_EXCEL_WORKSHEETS
*"* do not include other source files here!!!