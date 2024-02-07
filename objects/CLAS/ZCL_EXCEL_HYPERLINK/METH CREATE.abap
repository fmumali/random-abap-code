  METHOD create.
    DATA: lo_hyperlink TYPE REF TO zcl_excel_hyperlink.

    CREATE OBJECT lo_hyperlink.

    lo_hyperlink->location = iv_url.
    lo_hyperlink->internal = iv_internal.

    ov_link = lo_hyperlink.
  ENDMETHOD.