CLASS zcl_excel_collection_iterator DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.
    METHODS get_index
      RETURNING
        VALUE(index) TYPE i.
    METHODS has_next
      RETURNING
        VALUE(has_next) TYPE abap_bool.
    METHODS get_next
      RETURNING
        VALUE(object) TYPE REF TO object.
    METHODS constructor
      IMPORTING
        collection TYPE REF TO zcl_excel_collection.