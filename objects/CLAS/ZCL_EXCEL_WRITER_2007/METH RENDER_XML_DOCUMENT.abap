  METHOD render_xml_document.
    DATA lo_streamfactory TYPE REF TO if_ixml_stream_factory.
    DATA lo_ostream       TYPE REF TO if_ixml_ostream.
    DATA lo_renderer      TYPE REF TO if_ixml_renderer.
    DATA lv_string        TYPE string.

    " So that the rendering of io_document to a XML text in UTF-8 XSTRING works for all Unicode characters (Chinese,
    " emoticons, etc.) the method CREATE_OSTREAM_CSTRING must be used instead of CREATE_OSTREAM_XSTRING as explained
    " in note 2922674 below (original there: https://launchpad.support.sap.com/#/notes/2922674), and then the STRING
    " variable can be converted into UTF-8.
    "
    " Excerpt from Note 2922674 - Support for Unicode Characters U+10000 to U+10FFFF in the iXML kernel library / ABAP package SIXML.
    "
    "   You are running a unicode system with SAP Netweaver / SAP_BASIS release equal or lower than 7.51.
    "
    "   Some functions in the iXML kernel library / ABAP package SIXML does not fully or incorrectly support unicode
    "   characters of the supplementary planes. This is caused by using UCS-2 in codepage conversion functions.
    "   Therefore, when reading from iXML input steams, the characters from the supplementary planes, that are not
    "   supported by UCS-2, might be replaced by the character #. When writing to iXML output streams, UTF-16 surrogate
    "   pairs, representing characters from the supplementary planes, might be incorrectly encoded in UTF-8.
    "
    "   The characters incorrectly encoded in UTF-8, might be accepted as input for the iXML parser or external parsers,
    "   but might also be rejected.
    "
    "   Support for unicode characters of the supplementary planes was introduced for SAP_BASIS 7.51 or lower with note
    "   2220720, but later withdrawn with note 2346627 for functional issues.
    "
    "   Characters of the supplementary planes are supported with ABAP Platform 1709 / SAP_BASIS 7.52 and higher.
    "
    "   Please note, that the iXML runtime behaves like the ABAP runtime concerning the handling of unicode characters of
    "   the supplementary planes. In iXML and ABAP, these characters have length 2 (as returned by ABAP build-in function
    "   STRLEN), and string processing functions like SUBSTRING might split these characters into 2 invalid characters
    "   with length 1. These invalid characters are commonly referred to as broken surrogate pairs.
    "
    "   A workaround for the incorrect UTF-8 encoding in SAP_BASIS 7.51 or lower is to render the document to an ABAP
    "   variable with type STRING using a output stream created with factory method IF_IXML_STREAM_FACTORY=>CREATE_OSTREAM_CSTRING
    "   and then to convert the STRING variable to UTF-8 using method CL_ABAP_CODEPAGE=>CONVERT_TO.

    " 1) RENDER TO XML STRING
    lo_streamfactory = me->ixml->create_stream_factory( ).
    lo_ostream = lo_streamfactory->create_ostream_cstring( string = lv_string ).
    lo_renderer = me->ixml->create_renderer( ostream  = lo_ostream document = io_document ).
    lo_renderer->render( ).

    " 2) CONVERT IT TO UTF-8
    "-----------------
    " The beginning of the XML string has these 57 characters:
    "   X<?xml version="1.0" encoding="utf-16" standalone="yes"?>
    "   (where "X" is the special character corresponding to the utf-16 BOM, hexadecimal FFFE or FEFF,
    "   but there's no "X" in non-Unicode SAP systems)
    " The encoding must be removed otherwise Excel would fail to decode correctly the UTF-8 XML.
    " For a better performance, it's assumed that "encoding" is in the first 100 characters.
    IF strlen( lv_string ) < 100.
      REPLACE REGEX 'encoding="[^"]+"' IN lv_string WITH ``.
    ELSE.
      REPLACE REGEX 'encoding="[^"]+"' IN SECTION LENGTH 100 OF lv_string WITH ``.
    ENDIF.
    " Convert XML text to UTF-8 (NB: if 2 first bytes are the UTF-16 BOM, they are converted into 3 bytes of UTF-8 BOM)
    ep_content = cl_abap_codepage=>convert_to( source = lv_string ).
    " Add the UTF-8 Byte Order Mark if missing (NB: that serves as substitute of "encoding")
    IF xstrlen( ep_content ) >= 3 AND ep_content(3) <> cl_abap_char_utilities=>byte_order_mark_utf8.
      CONCATENATE cl_abap_char_utilities=>byte_order_mark_utf8 ep_content INTO ep_content IN BYTE MODE.
    ENDIF.

  ENDMETHOD.