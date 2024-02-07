  PRIVATE SECTION.

    DATA location TYPE string .
    DATA cell_reference TYPE string .
    DATA internal TYPE abap_bool .
    DATA column TYPE zexcel_cell_column_alpha .
    DATA row TYPE zexcel_cell_row .

    CLASS-METHODS create
      IMPORTING
        !iv_url        TYPE string
        !iv_internal   TYPE abap_bool
      RETURNING
        VALUE(ov_link) TYPE REF TO zcl_excel_hyperlink .