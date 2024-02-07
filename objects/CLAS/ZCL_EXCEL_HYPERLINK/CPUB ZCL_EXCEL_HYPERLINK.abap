CLASS zcl_excel_hyperlink DEFINITION
  PUBLIC
  FINAL
  CREATE PRIVATE .

*"* public components of class ZCL_EXCEL_HYPERLINK
*"* do not include other source files here!!!
  PUBLIC SECTION.

    CLASS-METHODS create_external_link
      IMPORTING
        !iv_url        TYPE string
      RETURNING
        VALUE(ov_link) TYPE REF TO zcl_excel_hyperlink .
    CLASS-METHODS create_internal_link
      IMPORTING
        !iv_location   TYPE string
      RETURNING
        VALUE(ov_link) TYPE REF TO zcl_excel_hyperlink .
    METHODS is_internal
      RETURNING
        VALUE(ev_ret) TYPE abap_bool .
    METHODS set_cell_reference
      IMPORTING
        !ip_column TYPE simple
        !ip_row    TYPE zexcel_cell_row
      RAISING
        zcx_excel .
    METHODS get_ref
      RETURNING
        VALUE(ev_ref) TYPE string .
    METHODS get_url
      RETURNING
        VALUE(ev_url) TYPE string .
*"* protected components of class ZCL_EXCEL_HYPERLINK
*"* do not include other source files here!!!