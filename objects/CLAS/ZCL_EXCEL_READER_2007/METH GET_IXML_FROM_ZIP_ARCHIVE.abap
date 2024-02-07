  METHOD get_ixml_from_zip_archive.

    DATA: lv_content       TYPE xstring,
          lo_ixml          TYPE REF TO if_ixml,
          lo_streamfactory TYPE REF TO if_ixml_stream_factory,
          lo_istream       TYPE REF TO if_ixml_istream,
          lo_parser        TYPE REF TO if_ixml_parser.

*--------------------------------------------------------------------*
* Load XML file from archive into an input stream,
* and parse that stream into an ixml object
*--------------------------------------------------------------------*
    lv_content        = me->get_from_zip_archive( i_filename ).
    lo_ixml           = cl_ixml=>create( ).
    lo_streamfactory  = lo_ixml->create_stream_factory( ).
    lo_istream        = lo_streamfactory->create_istream_xstring( lv_content ).
    r_ixml            = lo_ixml->create_document( ).
    lo_parser         = lo_ixml->create_parser( stream_factory = lo_streamfactory
                                                istream        = lo_istream
                                                document       = r_ixml ).
    lo_parser->set_normalizing( is_normalizing ).
    lo_parser->set_validating( mode = if_ixml_parser=>co_no_validation ).
    lo_parser->parse( ).

  ENDMETHOD.