CLASS zcl_excel_comments DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    METHODS add
      IMPORTING
        !ip_comment TYPE REF TO zcl_excel_comment .
    METHODS include
      IMPORTING
        !ip_comment TYPE REF TO zcl_excel_comment .
    METHODS clear .
    METHODS constructor .
    METHODS get
      IMPORTING
        !ip_index         TYPE zexcel_active_worksheet
      RETURNING
        VALUE(eo_comment) TYPE REF TO zcl_excel_comment .
    METHODS get_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS is_empty
      RETURNING
        VALUE(is_empty) TYPE flag .
    METHODS remove
      IMPORTING
        !ip_comment TYPE REF TO zcl_excel_comment .
    METHODS size
      RETURNING
        VALUE(ep_size) TYPE i .