  METHOD create_internal_link.
    ov_link = zcl_excel_hyperlink=>create( iv_url = iv_location
                                           iv_internal = abap_true ).
  ENDMETHOD.