CLASS zcl_excel_fill_template DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    TYPES: tv_range_name TYPE c LENGTH 30,
           BEGIN OF ts_range,
             sheet  TYPE zexcel_sheet_title,
             name   TYPE tv_range_name,
             start  TYPE zexcel_cell_row,
             stop   TYPE zexcel_cell_row,
             id     TYPE i,
             parent TYPE i,
             length TYPE i,
           END OF ts_range,
           tt_ranges TYPE STANDARD TABLE OF ts_range WITH DEFAULT KEY,
           BEGIN OF ts_variable,
             sheet  TYPE zexcel_sheet_title,
             parent TYPE i,
             name   TYPE tv_range_name,
           END OF ts_variable,
           BEGIN OF ts_name_style,
             sheet           TYPE zexcel_sheet_title,
             name            TYPE tv_range_name,
             parent          TYPE i,
             numeric_counter TYPE i,
             date_counter    TYPE i,
             time_counter    TYPE i,
             text_counter    TYPE i,
           END OF ts_name_style,
           tt_variables    TYPE STANDARD TABLE OF ts_variable WITH DEFAULT KEY,
           tt_sheet_titles TYPE STANDARD TABLE OF zexcel_sheet_title WITH DEFAULT KEY,
           tt_name_styles  TYPE STANDARD TABLE OF ts_name_style WITH DEFAULT KEY.

    DATA mt_sheet TYPE tt_sheet_titles READ-ONLY.
    DATA mt_range TYPE tt_ranges READ-ONLY.
    DATA mt_var TYPE tt_variables READ-ONLY.
    DATA mo_excel TYPE REF TO zcl_excel READ-ONLY.
    DATA mt_name_styles TYPE tt_name_styles READ-ONLY.

    CLASS-METHODS create
      IMPORTING
        !io_excel                 TYPE REF TO zcl_excel
      RETURNING
        VALUE(eo_template_filler) TYPE REF TO zcl_excel_fill_template
      RAISING
        zcx_excel.

    METHODS fill_sheet
      IMPORTING
        !iv_data TYPE zcl_excel_template_data=>ts_template_data_sheet
      RAISING
        zcx_excel .
