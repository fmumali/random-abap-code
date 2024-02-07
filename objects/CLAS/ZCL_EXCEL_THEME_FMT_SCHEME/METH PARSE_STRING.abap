  METHOD parse_string.
    DATA li_stream   TYPE REF TO if_ixml_istream.
    DATA li_ixml     TYPE REF TO if_ixml.
    DATA li_document TYPE REF TO if_ixml_document.
    DATA li_factory  TYPE REF TO if_ixml_stream_factory.
    DATA li_parser   TYPE REF TO if_ixml_parser.
    DATA li_istream  TYPE REF TO if_ixml_istream.

    li_ixml = cl_ixml=>create( ).
    li_document = li_ixml->create_document( ).
    li_factory = li_ixml->create_stream_factory( ).
    li_istream = li_factory->create_istream_string( iv_string ).
    li_parser = li_ixml->create_parser(
      stream_factory = li_factory
      istream        = li_istream
      document       = li_document ).
    li_parser->add_strip_space_element( ).
    li_parser->parse( ).
    li_istream->close( ).
    ri_node = li_document->get_first_child( ).

  ENDMETHOD.