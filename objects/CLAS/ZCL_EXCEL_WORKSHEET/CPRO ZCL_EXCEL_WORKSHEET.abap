  PROTECTED SECTION.
    METHODS set_table_reference
      IMPORTING
        !ip_column    TYPE zexcel_cell_column
        !ip_row       TYPE zexcel_cell_row
        !ir_table     TYPE REF TO zcl_excel_table
        !ip_fieldname TYPE zexcel_fieldname
        !ip_header    TYPE abap_bool
      RAISING
        zcx_excel .