CLASS zcl_excel_reader_2007 DEFINITION
  PUBLIC
  CREATE PUBLIC .

  PUBLIC SECTION.
*"* public components of class ZCL_EXCEL_READER_2007
*"* do not include other source files here!!!

    INTERFACES zif_excel_reader .

    CLASS-METHODS fill_struct_from_attributes
      IMPORTING
        !ip_element   TYPE REF TO if_ixml_element
      CHANGING
        !cp_structure TYPE any .