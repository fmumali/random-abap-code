  PRIVATE SECTION.

    DATA column_index TYPE int4 .
    DATA width TYPE f .
    DATA auto_size TYPE abap_bool .
    DATA visible TYPE abap_bool .
    DATA outline_level TYPE int4 .
    DATA collapsed TYPE abap_bool .
    DATA xf_index TYPE int4 .
    DATA style_guid TYPE zexcel_cell_style .
    DATA excel TYPE REF TO zcl_excel .
    DATA worksheet TYPE REF TO zcl_excel_worksheet .