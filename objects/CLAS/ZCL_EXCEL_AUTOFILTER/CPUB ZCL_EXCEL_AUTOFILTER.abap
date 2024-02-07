CLASS zcl_excel_autofilter DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

*"* public components of class ZCL_EXCEL_AUTOFILTER
*"* do not include other source files here!!!
  PUBLIC SECTION.

    TYPES tv_filter_rule TYPE string .
    TYPES tv_logical_operator TYPE c LENGTH 3 .
    TYPES:
      BEGIN OF ts_filter,
        column           TYPE zexcel_cell_column,
        rule             TYPE tv_filter_rule,
        t_values         TYPE HASHED TABLE OF zexcel_cell_value WITH UNIQUE KEY table_line,
        tr_textfilter1   TYPE RANGE OF string,
        logical_operator TYPE tv_logical_operator,
        tr_textfilter2   TYPE RANGE OF string,
      END OF ts_filter .
    TYPES:
      tt_filters TYPE HASHED TABLE OF ts_filter WITH UNIQUE KEY column .

    DATA filter_area TYPE zexcel_s_autofilter_area .
    CONSTANTS mc_filter_rule_single_values TYPE tv_filter_rule VALUE 'single_values'. "#EC NOTEXT
    CONSTANTS mc_filter_rule_text_pattern TYPE tv_filter_rule VALUE 'text_pattern'. "#EC NOTEXT
    CONSTANTS mc_logical_operator_and TYPE tv_logical_operator VALUE 'and'. "#EC NOTEXT
    CONSTANTS mc_logical_operator_none TYPE tv_logical_operator VALUE space. "#EC NOTEXT
    CONSTANTS mc_logical_operator_or TYPE tv_logical_operator VALUE 'or'. "#EC NOTEXT

    METHODS constructor
      IMPORTING
        !io_sheet TYPE REF TO zcl_excel_worksheet .
    METHODS get_filter_area
      RETURNING
        VALUE(rs_area) TYPE zexcel_s_autofilter_area
      RAISING
        zcx_excel .
    METHODS get_filter_range
      RETURNING
        VALUE(r_range) TYPE zexcel_cell_value
      RAISING
        zcx_excel.
    METHODS get_filter_reference
      RETURNING
        VALUE(r_ref) TYPE zexcel_range_value
      RAISING
        zcx_excel .
    METHODS get_values
      RETURNING
        VALUE(rt_filter) TYPE zexcel_t_autofilter_values .
    METHODS is_row_hidden
      IMPORTING
        !iv_row             TYPE zexcel_cell_row
      RETURNING
        VALUE(rv_is_hidden) TYPE abap_bool .
    METHODS set_filter_area
      IMPORTING
        !is_area TYPE zexcel_s_autofilter_area .
    METHODS set_text_filter
      IMPORTING
        !i_column            TYPE zexcel_cell_column
        !iv_textfilter1      TYPE clike
        !iv_logical_operator TYPE tv_logical_operator DEFAULT mc_logical_operator_none
        !iv_textfilter2      TYPE clike OPTIONAL .
    METHODS set_value
      IMPORTING
        !i_column TYPE zexcel_cell_column
        !i_value  TYPE zexcel_cell_value .
    METHODS set_values
      IMPORTING
        !it_values TYPE zexcel_t_autofilter_values .
*"* protected components of class ZABAP_EXCEL_WORKSHEET
*"* do not include other source files here!!!