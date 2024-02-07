CLASS zcl_excel_style_borders DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

*"* public components of class ZCL_EXCEL_STYLE_BORDERS
*"* do not include other source files here!!!
  PUBLIC SECTION.

    DATA allborders TYPE REF TO zcl_excel_style_border .
    CONSTANTS c_diagonal_both TYPE zexcel_diagonal VALUE 3. "#EC NOTEXT
    CONSTANTS c_diagonal_down TYPE zexcel_diagonal VALUE 2. "#EC NOTEXT
    CONSTANTS c_diagonal_none TYPE zexcel_diagonal VALUE 0. "#EC NOTEXT
    CONSTANTS c_diagonal_up TYPE zexcel_diagonal VALUE 1.   "#EC NOTEXT
    DATA diagonal TYPE REF TO zcl_excel_style_border .
    DATA diagonal_mode TYPE zexcel_diagonal .
    DATA down TYPE REF TO zcl_excel_style_border .
    DATA left TYPE REF TO zcl_excel_style_border .
    DATA right TYPE REF TO zcl_excel_style_border .
    DATA top TYPE REF TO zcl_excel_style_border .

    METHODS get_structure
      RETURNING
        VALUE(es_fill) TYPE zexcel_s_style_border .
*"* protected components of class ZABAP_EXCEL_STYLE_FONT
*"* do not include other source files here!!!
*"* protected components of class ZABAP_EXCEL_STYLE_FONT
*"* do not include other source files here!!!