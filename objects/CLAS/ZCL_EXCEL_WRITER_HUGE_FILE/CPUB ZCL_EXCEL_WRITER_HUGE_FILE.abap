CLASS zcl_excel_writer_huge_file DEFINITION
  PUBLIC
  INHERITING FROM zcl_excel_writer_2007
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    TYPES:
      BEGIN OF ty_cell,
        name    TYPE c LENGTH 10, "AAA1234567"
        style   TYPE i,
        type    TYPE c LENGTH 9,
        formula TYPE string,
        value   TYPE string,
      END OF ty_cell .

    DATA:
      cells TYPE STANDARD TABLE OF ty_cell .

    METHODS get_cells
      IMPORTING
        !i_row   TYPE i
        !i_index TYPE i .