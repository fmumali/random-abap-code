CLASS zcl_excel_data_validations DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

*"* public components of class ZCL_EXCEL_DATA_VALIDATIONS
*"* do not include other source files here!!!
  PUBLIC SECTION.

    METHODS add
      IMPORTING
        !ip_data_validation TYPE REF TO zcl_excel_data_validation .
    METHODS clear .
    METHODS constructor .
    METHODS get_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS is_empty
      RETURNING
        VALUE(is_empty) TYPE flag .
    METHODS remove
      IMPORTING
        !ip_data_validation TYPE REF TO zcl_excel_data_validation .
    METHODS size
      RETURNING
        VALUE(ep_size) TYPE i .
*"* protected components of class ZCL_EXCEL_DATA_VALIDATIONS
*"* do not include other source files here!!!