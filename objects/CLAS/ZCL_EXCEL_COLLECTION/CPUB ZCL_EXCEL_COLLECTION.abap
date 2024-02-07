CLASS zcl_excel_collection DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    TYPES:
      ty_collection TYPE STANDARD TABLE OF REF TO object .

    DATA collection TYPE ty_collection READ-ONLY .

    METHODS size
      RETURNING
        VALUE(size) TYPE i .
    METHODS is_empty
      RETURNING
        VALUE(is_empty) TYPE abap_bool .
    METHODS get
      IMPORTING
        !index        TYPE i
      RETURNING
        VALUE(object) TYPE REF TO object .
    METHODS get_iterator
      RETURNING
        VALUE(iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS add
      IMPORTING
        !element TYPE REF TO object .
    METHODS remove
      IMPORTING
        !element TYPE REF TO object .
    METHODS clear .