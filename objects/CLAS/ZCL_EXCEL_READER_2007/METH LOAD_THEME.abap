  METHOD load_theme.
    DATA theme TYPE REF TO zcl_excel_theme.
    DATA: lo_theme_xml TYPE REF TO if_ixml_document.
    CREATE OBJECT theme.
    lo_theme_xml = me->get_ixml_from_zip_archive( iv_path ).
    theme->read_theme( io_theme_xml = lo_theme_xml  ).
    ip_excel->set_theme( io_theme = theme ).
  ENDMETHOD.