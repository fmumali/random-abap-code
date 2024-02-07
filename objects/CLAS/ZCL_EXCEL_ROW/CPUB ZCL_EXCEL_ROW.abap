CLASS zcl_excel_row DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

*"* public components of class ZCL_EXCEL_ROW
*"* do not include other source files here!!!
  PUBLIC SECTION.

    METHODS constructor
      IMPORTING
        !ip_index TYPE int4 DEFAULT 0 .
    METHODS get_collapsed
      IMPORTING
        !io_worksheet      TYPE REF TO zcl_excel_worksheet OPTIONAL
      RETURNING
        VALUE(r_collapsed) TYPE abap_bool .
    METHODS get_outline_level
      IMPORTING
        !io_worksheet          TYPE REF TO zcl_excel_worksheet OPTIONAL
      RETURNING
        VALUE(r_outline_level) TYPE int4 .
    METHODS get_row_height
      RETURNING
        VALUE(r_row_height) TYPE f .
    METHODS get_custom_height
      RETURNING
        VALUE(r_custom_height) TYPE abap_bool .
    METHODS get_row_index
      RETURNING
        VALUE(r_row_index) TYPE int4 .
    METHODS get_visible
      IMPORTING
        !io_worksheet    TYPE REF TO zcl_excel_worksheet OPTIONAL
      RETURNING
        VALUE(r_visible) TYPE abap_bool .
    METHODS get_xf_index
      RETURNING
        VALUE(r_xf_index) TYPE int4 .
    METHODS set_collapsed
      IMPORTING
        !ip_collapsed TYPE abap_bool .
    METHODS set_outline_level
      IMPORTING
        !ip_outline_level TYPE int4
      RAISING
        zcx_excel .
    METHODS set_row_height
      IMPORTING
        !ip_row_height    TYPE simple
        !ip_custom_height TYPE abap_bool DEFAULT abap_true
      RAISING
        zcx_excel .
    METHODS set_row_index
      IMPORTING
        !ip_index TYPE int4 .
    METHODS set_visible
      IMPORTING
        !ip_visible TYPE abap_bool .
    METHODS set_xf_index
      IMPORTING
        !ip_xf_index TYPE int4 .
*"* protected components of class ZCL_EXCEL_ROW
*"* do not include other source files here!!!