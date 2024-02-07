  METHOD create_xml_document.
    DATA lo_encoding TYPE REF TO if_ixml_encoding.
    lo_encoding = me->ixml->create_encoding( byte_order = if_ixml_encoding=>co_platform_endian
                                             character_set = 'utf-8' ).
    ro_document = me->ixml->create_document( ).
    ro_document->set_encoding( lo_encoding ).
    ro_document->set_standalone( abap_true ).
  ENDMETHOD.