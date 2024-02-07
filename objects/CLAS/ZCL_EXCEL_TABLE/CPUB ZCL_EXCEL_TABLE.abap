CLASS zcl_excel_table DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

*"* public components of class ZCL_EXCEL_TABLE
*"* do not include other source files here!!!
  PUBLIC SECTION.

    CONSTANTS builtinstyle_dark1 TYPE zexcel_table_style VALUE 'TableStyleDark1'. "#EC NOTEXT
    CONSTANTS builtinstyle_dark2 TYPE zexcel_table_style VALUE 'TableStyleDark2'. "#EC NOTEXT
    CONSTANTS builtinstyle_dark3 TYPE zexcel_table_style VALUE 'TableStyleDark3'. "#EC NOTEXT
    CONSTANTS builtinstyle_dark4 TYPE zexcel_table_style VALUE 'TableStyleDark4'. "#EC NOTEXT
    CONSTANTS builtinstyle_dark5 TYPE zexcel_table_style VALUE 'TableStyleDark5'. "#EC NOTEXT
    CONSTANTS builtinstyle_dark6 TYPE zexcel_table_style VALUE 'TableStyleDark6'. "#EC NOTEXT
    CONSTANTS builtinstyle_dark7 TYPE zexcel_table_style VALUE 'TableStyleDark7'. "#EC NOTEXT
    CONSTANTS builtinstyle_dark8 TYPE zexcel_table_style VALUE 'TableStyleDark8'. "#EC NOTEXT
    CONSTANTS builtinstyle_dark9 TYPE zexcel_table_style VALUE 'TableStyleDark9'. "#EC NOTEXT
    CONSTANTS builtinstyle_dark10 TYPE zexcel_table_style VALUE 'TableStyleDark10'. "#EC NOTEXT
    CONSTANTS builtinstyle_dark11 TYPE zexcel_table_style VALUE 'TableStyleDark11'. "#EC NOTEXT
    CONSTANTS builtinstyle_light1 TYPE zexcel_table_style VALUE 'TableStyleLight1'. "#EC NOTEXT
    CONSTANTS builtinstyle_light2 TYPE zexcel_table_style VALUE 'TableStyleLight2'. "#EC NOTEXT
    CONSTANTS builtinstyle_light3 TYPE zexcel_table_style VALUE 'TableStyleLight3'. "#EC NOTEXT
    CONSTANTS builtinstyle_light4 TYPE zexcel_table_style VALUE 'TableStyleLight4'. "#EC NOTEXT
    CONSTANTS builtinstyle_light5 TYPE zexcel_table_style VALUE 'TableStyleLight5'. "#EC NOTEXT
    CONSTANTS builtinstyle_pivot_light16 TYPE zexcel_table_style VALUE 'PivotStyleLight16'. "#EC NOTEXT
    CONSTANTS builtinstyle_light6 TYPE zexcel_table_style VALUE 'TableStyleLight6'. "#EC NOTEXT
    CONSTANTS totals_function_average TYPE zexcel_table_totals_function VALUE 'average'. "#EC NOTEXT
    CONSTANTS builtinstyle_light7 TYPE zexcel_table_style VALUE 'TableStyleLight7'. "#EC NOTEXT
    CONSTANTS totals_function_count TYPE zexcel_table_totals_function VALUE 'count'. "#EC NOTEXT
    CONSTANTS builtinstyle_light8 TYPE zexcel_table_style VALUE 'TableStyleLight8'. "#EC NOTEXT
    CONSTANTS totals_function_custom TYPE zexcel_table_totals_function VALUE 'custom'. "#EC NOTEXT
    CONSTANTS builtinstyle_light9 TYPE zexcel_table_style VALUE 'TableStyleLight9'. "#EC NOTEXT
    CONSTANTS totals_function_max TYPE zexcel_table_totals_function VALUE 'max'. "#EC NOTEXT
    CONSTANTS builtinstyle_light10 TYPE zexcel_table_style VALUE 'TableStyleLight10'. "#EC NOTEXT
    CONSTANTS totals_function_min TYPE zexcel_table_totals_function VALUE 'min'. "#EC NOTEXT
    CONSTANTS builtinstyle_light11 TYPE zexcel_table_style VALUE 'TableStyleLight11'. "#EC NOTEXT
    CONSTANTS totals_function_sum TYPE zexcel_table_totals_function VALUE 'sum'. "#EC NOTEXT
    CONSTANTS builtinstyle_light12 TYPE zexcel_table_style VALUE 'TableStyleLight12'. "#EC NOTEXT
    DATA fieldcat TYPE zexcel_t_fieldcatalog .
    CONSTANTS builtinstyle_light13 TYPE zexcel_table_style VALUE 'TableStyleLight13'. "#EC NOTEXT
    CONSTANTS builtinstyle_light14 TYPE zexcel_table_style VALUE 'TableStyleLight14'. "#EC NOTEXT
    DATA settings TYPE zexcel_s_table_settings .
    CONSTANTS builtinstyle_light15 TYPE zexcel_table_style VALUE 'TableStyleLight15'. "#EC NOTEXT
    CONSTANTS builtinstyle_light16 TYPE zexcel_table_style VALUE 'TableStyleLight16'. "#EC NOTEXT
    CONSTANTS builtinstyle_light17 TYPE zexcel_table_style VALUE 'TableStyleLight17'. "#EC NOTEXT
    CONSTANTS builtinstyle_light18 TYPE zexcel_table_style VALUE 'TableStyleLight18'. "#EC NOTEXT
    CONSTANTS builtinstyle_light19 TYPE zexcel_table_style VALUE 'TableStyleLight19'. "#EC NOTEXT
    CONSTANTS builtinstyle_light20 TYPE zexcel_table_style VALUE 'TableStyleLight20'. "#EC NOTEXT
    CONSTANTS builtinstyle_light21 TYPE zexcel_table_style VALUE 'TableStyleLight21'. "#EC NOTEXT
    CONSTANTS builtinstyle_medium1 TYPE zexcel_table_style VALUE 'TableStyleMedium1'. "#EC NOTEXT
    CONSTANTS builtinstyle_medium2 TYPE zexcel_table_style VALUE 'TableStyleMedium2'. "#EC NOTEXT
    CONSTANTS builtinstyle_medium3 TYPE zexcel_table_style VALUE 'TableStyleMedium3'. "#EC NOTEXT
    CONSTANTS builtinstyle_medium4 TYPE zexcel_table_style VALUE 'TableStyleMedium4'. "#EC NOTEXT
    CONSTANTS builtinstyle_medium5 TYPE zexcel_table_style VALUE 'TableStyleMedium5'. "#EC NOTEXT
    CONSTANTS builtinstyle_medium6 TYPE zexcel_table_style VALUE 'TableStyleMedium6'. "#EC NOTEXT
    CONSTANTS builtinstyle_medium7 TYPE zexcel_table_style VALUE 'TableStyleMedium7'. "#EC NOTEXT
    CONSTANTS builtinstyle_medium8 TYPE zexcel_table_style VALUE 'TableStyleMedium8'. "#EC NOTEXT
    CONSTANTS builtinstyle_medium9 TYPE zexcel_table_style VALUE 'TableStyleMedium9'. "#EC NOTEXT
    CONSTANTS builtinstyle_medium10 TYPE zexcel_table_style VALUE 'TableStyleMedium10'. "#EC NOTEXT
    CONSTANTS builtinstyle_medium11 TYPE zexcel_table_style VALUE 'TableStyleMedium11'. "#EC NOTEXT
    CONSTANTS builtinstyle_medium12 TYPE zexcel_table_style VALUE 'TableStyleMedium12'. "#EC NOTEXT
    CONSTANTS builtinstyle_medium13 TYPE zexcel_table_style VALUE 'TableStyleMedium13'. "#EC NOTEXT
    CONSTANTS builtinstyle_medium14 TYPE zexcel_table_style VALUE 'TableStyleMedium14'. "#EC NOTEXT
    CONSTANTS builtinstyle_medium15 TYPE zexcel_table_style VALUE 'TableStyleMedium15'. "#EC NOTEXT
    CONSTANTS builtinstyle_medium16 TYPE zexcel_table_style VALUE 'TableStyleMedium16'. "#EC NOTEXT
    CONSTANTS builtinstyle_medium17 TYPE zexcel_table_style VALUE 'TableStyleMedium17'. "#EC NOTEXT
    CONSTANTS builtinstyle_medium18 TYPE zexcel_table_style VALUE 'TableStyleMedium18'. "#EC NOTEXT
    CONSTANTS builtinstyle_medium19 TYPE zexcel_table_style VALUE 'TableStyleMedium19'. "#EC NOTEXT
    CONSTANTS builtinstyle_medium20 TYPE zexcel_table_style VALUE 'TableStyleMedium20'. "#EC NOTEXT
    CONSTANTS builtinstyle_medium21 TYPE zexcel_table_style VALUE 'TableStyleMedium21'. "#EC NOTEXT
    CONSTANTS builtinstyle_medium22 TYPE zexcel_table_style VALUE 'TableStyleMedium22'. "#EC NOTEXT
    CONSTANTS builtinstyle_medium23 TYPE zexcel_table_style VALUE 'TableStyleMedium23'. "#EC NOTEXT
    CONSTANTS builtinstyle_medium24 TYPE zexcel_table_style VALUE 'TableStyleMedium24'. "#EC NOTEXT
    CONSTANTS builtinstyle_medium25 TYPE zexcel_table_style VALUE 'TableStyleMedium26'. "#EC NOTEXT
    CONSTANTS builtinstyle_medium27 TYPE zexcel_table_style VALUE 'TableStyleMedium27'. "#EC NOTEXT

    METHODS get_totals_formula
      IMPORTING
        !ip_column        TYPE clike
        !ip_function      TYPE zexcel_table_totals_function
      RETURNING
        VALUE(ep_formula) TYPE string
      RAISING
        zcx_excel .
    METHODS has_totals
      RETURNING
        VALUE(ep_result) TYPE abap_bool .
    METHODS set_data
      IMPORTING
        !ir_data TYPE STANDARD TABLE .
    METHODS get_id
      RETURNING
        VALUE(ov_id) TYPE i .
    METHODS set_id
      IMPORTING
        !iv_id TYPE i .
    METHODS get_name
      RETURNING
        VALUE(ov_name) TYPE string .
    METHODS get_reference
      IMPORTING
        !ip_include_totals_row TYPE abap_bool DEFAULT abap_true
      RETURNING
        VALUE(ov_reference)    TYPE string
      RAISING
        zcx_excel .
    METHODS get_bottom_row_integer
      RETURNING
        VALUE(ev_row) TYPE i .
    METHODS get_right_column_integer
      RETURNING
        VALUE(ev_column) TYPE i
      RAISING
        zcx_excel .
*"* protected components of class ZCL_EXCEL_TABLE
*"* do not include other source files here!!!
*"* protected components of class ZCL_EXCEL_TABLE
*"* do not include other source files here!!!
*"* protected components of class ZCL_EXCEL_TABLE
*"* do not include other source files here!!!