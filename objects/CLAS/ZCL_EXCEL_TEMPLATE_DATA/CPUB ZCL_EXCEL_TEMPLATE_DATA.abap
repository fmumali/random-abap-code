CLASS zcl_excel_template_data DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    TYPES: tt_sheet_titles TYPE STANDARD TABLE OF zexcel_sheet_title WITH DEFAULT KEY,
           BEGIN OF ts_template_data_sheet,
             sheet TYPE zexcel_sheet_title,
             data  TYPE REF TO data,
           END OF ts_template_data_sheet,
           tt_template_data_sheets TYPE STANDARD TABLE OF ts_template_data_sheet WITH DEFAULT KEY.

    DATA mt_data TYPE tt_template_data_sheets READ-ONLY.

    METHODS add
      IMPORTING
        iv_sheet TYPE zexcel_sheet_title
        iv_data  TYPE data .