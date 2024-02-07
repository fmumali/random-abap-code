CLASS zcl_excel_converter_alv DEFINITION
  PUBLIC
  ABSTRACT
  CREATE PUBLIC .

*"* public components of class ZCL_EXCEL_CONVERTER_ALV
*"* do not include other source files here!!!
  PUBLIC SECTION.

    INTERFACES zif_excel_converter
      ALL METHODS ABSTRACT .

    CLASS-METHODS class_constructor .
*"* protected components of class ZCL_EXCEL_CONVERTER_ALV
*"* do not include other source files here!!!