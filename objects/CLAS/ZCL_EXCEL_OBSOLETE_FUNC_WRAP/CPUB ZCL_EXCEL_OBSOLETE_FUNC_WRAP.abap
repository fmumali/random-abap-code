CLASS zcl_excel_obsolete_func_wrap DEFINITION
  PUBLIC
  CREATE PUBLIC .

  PUBLIC SECTION.

    CLASS-METHODS guid_create
      RETURNING
        VALUE(rv_guid_16) TYPE zexcel_guid .