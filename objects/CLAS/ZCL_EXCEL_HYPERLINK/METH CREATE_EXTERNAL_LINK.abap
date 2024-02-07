  METHOD create_external_link.

    ov_link = zcl_excel_hyperlink=>create( iv_url = iv_url
                                           iv_internal = abap_false ).
  ENDMETHOD.