CLASS zcl_excel_drawings DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

*"* public components of class ZCL_EXCEL_DRAWINGS
*"* do not include other source files here!!!
    DATA type TYPE zexcel_drawing_type READ-ONLY VALUE 'IMAGE'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .

    METHODS add
      IMPORTING
        !ip_drawing TYPE REF TO zcl_excel_drawing .
    METHODS include
      IMPORTING
        !ip_drawing TYPE REF TO zcl_excel_drawing .
    METHODS clear .
    METHODS constructor
      IMPORTING
        !ip_type TYPE zexcel_drawing_type .
    METHODS get
      IMPORTING
        !ip_index         TYPE zexcel_active_worksheet
      RETURNING
        VALUE(eo_drawing) TYPE REF TO zcl_excel_drawing .
    METHODS get_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS is_empty
      RETURNING
        VALUE(is_empty) TYPE flag .
    METHODS remove
      IMPORTING
        !ip_drawing TYPE REF TO zcl_excel_drawing .
    METHODS size
      RETURNING
        VALUE(ep_size) TYPE i .
    METHODS get_type
      RETURNING
        VALUE(rp_type) TYPE zexcel_drawing_type .
*"* protected components of class ZCL_EXCEL_DRAWINGS
*"* do not include other source files here!!!
*"* protected components of class ZCL_EXCEL_DRAWINGS
*"* do not include other source files here!!!