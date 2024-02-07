CLASS zcl_excel_column DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

*"* public components of class ZCL_EXCEL_COLUMN
*"* do not include other source files here!!!
  PUBLIC SECTION.

    METHODS constructor
      IMPORTING
        !ip_index     TYPE zexcel_cell_column_alpha
        !ip_worksheet TYPE REF TO zcl_excel_worksheet
        !ip_excel     TYPE REF TO zcl_excel
      RAISING
        zcx_excel .
    METHODS get_auto_size
      RETURNING
        VALUE(r_auto_size) TYPE abap_bool .
    METHODS get_collapsed
      RETURNING
        VALUE(r_collapsed) TYPE abap_bool .
    METHODS get_column_index
      RETURNING
        VALUE(r_column_index) TYPE int4 .
    METHODS get_outline_level
      RETURNING
        VALUE(r_outline_level) TYPE int4 .
    METHODS get_visible
      RETURNING
        VALUE(r_visible) TYPE abap_bool .
    METHODS get_width
      RETURNING
        VALUE(r_width) TYPE f .
    METHODS get_xf_index
      RETURNING
        VALUE(r_xf_index) TYPE int4 .
    METHODS set_auto_size
      IMPORTING
        !ip_auto_size    TYPE abap_bool
      RETURNING
        VALUE(io_column) TYPE REF TO zcl_excel_column .
    METHODS set_collapsed
      IMPORTING
        !ip_collapsed    TYPE abap_bool
      RETURNING
        VALUE(io_column) TYPE REF TO zcl_excel_column .
    METHODS set_column_index
      IMPORTING
        !ip_index        TYPE zexcel_cell_column_alpha
      RETURNING
        VALUE(io_column) TYPE REF TO zcl_excel_column
      RAISING
        zcx_excel .
    METHODS set_outline_level
      IMPORTING
        !ip_outline_level TYPE int4 .
    METHODS set_visible
      IMPORTING
        !ip_visible      TYPE abap_bool
      RETURNING
        VALUE(io_column) TYPE REF TO zcl_excel_column .
    METHODS set_width
      IMPORTING
        !ip_width        TYPE simple
      RETURNING
        VALUE(io_column) TYPE REF TO zcl_excel_column
      RAISING
        zcx_excel .
    METHODS set_xf_index
      IMPORTING
        !ip_xf_index     TYPE int4
      RETURNING
        VALUE(io_column) TYPE REF TO zcl_excel_column .
    METHODS set_column_style_by_guid
      IMPORTING
        !ip_style_guid TYPE zexcel_cell_style
      RAISING
        zcx_excel .
    METHODS get_column_style_guid
      RETURNING
        VALUE(ep_style_guid) TYPE zexcel_cell_style
      RAISING
        zcx_excel .
*"* protected components of class ZCL_EXCEL_COLUMN
*"* do not include other source files here!!!