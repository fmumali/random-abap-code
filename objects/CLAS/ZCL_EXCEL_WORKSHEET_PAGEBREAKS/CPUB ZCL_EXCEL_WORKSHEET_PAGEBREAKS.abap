CLASS zcl_excel_worksheet_pagebreaks DEFINITION
  PUBLIC
  CREATE PUBLIC .

  PUBLIC SECTION.

    TYPES:
      BEGIN OF ts_pagebreak_at ,
        cell_row    TYPE zexcel_cell_row,
        cell_column TYPE zexcel_cell_column,
      END OF ts_pagebreak_at .
    TYPES:
      tt_pagebreak_at TYPE HASHED TABLE OF  ts_pagebreak_at WITH UNIQUE KEY cell_row cell_column .

    METHODS add_pagebreak
      IMPORTING
        !ip_column TYPE simple
        !ip_row    TYPE zexcel_cell_row
      RAISING
        zcx_excel .
    METHODS get_all_pagebreaks
      RETURNING
        VALUE(rt_pagebreaks) TYPE tt_pagebreak_at .