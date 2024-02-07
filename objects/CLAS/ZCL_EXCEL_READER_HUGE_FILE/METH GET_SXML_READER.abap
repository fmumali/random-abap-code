  METHOD get_sxml_reader.

    DATA: lv_xml TYPE xstring.

    lv_xml = get_from_zip_archive( iv_path ).
    eo_reader = cl_sxml_string_reader=>create( lv_xml ).

  ENDMETHOD.