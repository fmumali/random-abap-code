CLASS zcl_excel_worksheets DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

*"* public components of class ZCL_EXCEL_WORKSHEETS
*"* do not include other source files here!!!
  PUBLIC SECTION.

    DATA active_worksheet TYPE zexcel_active_worksheet VALUE 1. "#EC NOTEXT .  .  .  .  .  .  .  .  . " .
    DATA name TYPE zexcel_worksheets_name VALUE 'Worksheets'. "#EC NOTEXT .  .  .  .  .  .  .  .  . " .

    METHODS add
      IMPORTING
        !ip_worksheet TYPE REF TO zcl_excel_worksheet .
    METHODS clear .
    METHODS constructor .
    METHODS get
      IMPORTING
        !ip_index           TYPE zexcel_active_worksheet
      RETURNING
        VALUE(eo_worksheet) TYPE REF TO zcl_excel_worksheet .
    METHODS get_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS is_empty
      RETURNING
        VALUE(is_empty) TYPE flag .
    METHODS remove
      IMPORTING
        !ip_worksheet TYPE REF TO zcl_excel_worksheet .
    METHODS size
      RETURNING
        VALUE(ep_size) TYPE i .
*"* protected components of class ZCL_EXCEL_WORKSHEETS
*"* do not include other source files here!!!