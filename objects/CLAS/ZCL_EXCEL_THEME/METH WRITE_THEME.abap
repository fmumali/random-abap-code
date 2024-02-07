  METHOD write_theme.
    DATA: lo_ixml         TYPE REF TO if_ixml,
          lo_element_root TYPE REF TO if_ixml_element,
          lo_encoding     TYPE REF TO if_ixml_encoding.
    DATA: lo_streamfactory  TYPE REF TO if_ixml_stream_factory.
    DATA: lo_ostream TYPE REF TO if_ixml_ostream.
    DATA: lo_renderer TYPE REF TO if_ixml_renderer.
    DATA: lo_document TYPE REF TO if_ixml_document.
    lo_ixml = cl_ixml=>create( ).

    lo_encoding = lo_ixml->create_encoding( byte_order = if_ixml_encoding=>co_platform_endian
                                            character_set = 'UTF-8' ).
    lo_document = lo_ixml->create_document( ).
    lo_document->set_encoding( lo_encoding ).
    lo_document->set_standalone( abap_true ).

    lo_element_root = lo_document->create_simple_element_ns( prefix = c_theme_prefix
                                                             name   = c_theme
                                                            parent = lo_document
                                                            ).
    lo_element_root->set_attribute_ns( name  = c_theme_xmlns
                                       value = c_theme_xmlns_val ).
    lo_element_root->set_attribute_ns( name  = c_theme_name
                                       value = name ).

    elements->build_xml( io_document = lo_document ).
    objectdefaults->build_xml( io_document = lo_document ).
    extclrschemelst->build_xml( io_document = lo_document ).
    extlst->build_xml( io_document = lo_document ).

    lo_streamfactory = lo_ixml->create_stream_factory( ).
    lo_ostream = lo_streamfactory->create_ostream_xstring( string = rv_xstring ).
    lo_renderer = lo_ixml->create_renderer( ostream  = lo_ostream document = lo_document ).
    lo_renderer->render( ).

  ENDMETHOD.                    "write_theme