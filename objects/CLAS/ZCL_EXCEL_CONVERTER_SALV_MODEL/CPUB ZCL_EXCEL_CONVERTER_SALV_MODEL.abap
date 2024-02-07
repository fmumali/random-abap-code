CLASS zcl_excel_converter_salv_model DEFINITION
  PUBLIC
  INHERITING FROM cl_salv_model
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.
    CLASS-METHODS is_get_metadata_callable
      IMPORTING
        io_salv       TYPE REF TO cl_salv_table
      RETURNING
        VALUE(result) TYPE abap_bool.