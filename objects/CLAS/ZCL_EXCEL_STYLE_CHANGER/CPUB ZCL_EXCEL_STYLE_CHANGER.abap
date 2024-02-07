CLASS zcl_excel_style_changer DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    INTERFACES zif_excel_style_changer.

    CLASS-METHODS create
      IMPORTING
        excel         TYPE REF TO zcl_excel
      RETURNING
        VALUE(result) TYPE REF TO zif_excel_style_changer
      RAISING
        zcx_excel.
