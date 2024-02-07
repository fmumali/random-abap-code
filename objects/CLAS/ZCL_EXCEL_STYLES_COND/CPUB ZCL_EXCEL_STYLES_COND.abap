CLASS zcl_excel_styles_cond DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

*"* public components of class ZCL_EXCEL_STYLES_COND
*"* do not include other source files here!!!
  PUBLIC SECTION.

    METHODS add
      IMPORTING
        !ip_style_cond TYPE REF TO zcl_excel_style_cond .
    METHODS clear .
    METHODS constructor .
    METHODS get
      IMPORTING
        !ip_index            TYPE zexcel_active_worksheet
      RETURNING
        VALUE(eo_style_cond) TYPE REF TO zcl_excel_style_cond .
    METHODS get_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS is_empty
      RETURNING
        VALUE(is_empty) TYPE flag .
    METHODS remove
      IMPORTING
        !ip_style_cond TYPE REF TO zcl_excel_style_cond .
    METHODS size
      RETURNING
        VALUE(ep_size) TYPE i .
*"* protected components of class ZABAP_EXCEL_WORKSHEETS
*"* do not include other source files here!!!