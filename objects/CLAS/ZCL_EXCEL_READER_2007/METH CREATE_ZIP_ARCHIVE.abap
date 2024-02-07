  METHOD create_zip_archive.
    CASE i_use_alternate_zip.
      WHEN space.
        e_zip = lcl_abap_zip_archive=>create( i_xlsx_binary ).
      WHEN OTHERS.
        e_zip = lcl_alternate_zip_archive=>create( i_data                = i_xlsx_binary
                                                   i_alternate_zip_class = i_use_alternate_zip ).
    ENDCASE.
  ENDMETHOD.