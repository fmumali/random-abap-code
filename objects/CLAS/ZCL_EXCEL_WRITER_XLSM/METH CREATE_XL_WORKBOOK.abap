  METHOD create_xl_workbook.

** Constant node name
    DATA: lc_xml_attr_codename      TYPE string VALUE 'codeName'.

    DATA: lo_ixml          TYPE REF TO if_ixml,
          lo_document      TYPE REF TO if_ixml_document,
          lo_document_xml  TYPE REF TO cl_xml_document,
          lo_element_root  TYPE REF TO if_ixml_node,
          lo_element       TYPE REF TO if_ixml_element,
          lo_collection    TYPE REF TO if_ixml_node_collection,
          lo_iterator      TYPE REF TO if_ixml_node_iterator,
          lo_streamfactory TYPE REF TO if_ixml_stream_factory,
          lo_ostream       TYPE REF TO if_ixml_ostream,
          lo_renderer      TYPE REF TO if_ixml_renderer.

**********************************************************************
* STEP 3: Create standard relationship
    ep_content = super->create_xl_workbook( ).

**********************************************************************
* STEP 2: modify XML adding the vbaProject relation

    CREATE OBJECT lo_document_xml.
    lo_document_xml->parse_xstring( ep_content ).

    lo_document ?= lo_document_xml->m_document.
    lo_element_root = lo_document->if_ixml_node~get_first_child( ).

    lo_collection = lo_document->get_elements_by_tag_name( 'fileVersion' ).
    lo_iterator = lo_collection->create_iterator( ).
    lo_element ?= lo_iterator->get_next( ).
    WHILE lo_element IS BOUND.
      lo_element->set_attribute_ns( name  = lc_xml_attr_codename
                                    value = me->excel->zif_excel_book_vba_project~codename ).
      lo_element ?= lo_iterator->get_next( ).
    ENDWHILE.

    lo_collection = lo_document->get_elements_by_tag_name( 'workbookPr' ).
    lo_iterator = lo_collection->create_iterator( ).
    lo_element ?= lo_iterator->get_next( ).
    WHILE lo_element IS BOUND.
      lo_element->set_attribute_ns( name  = lc_xml_attr_codename
                                    value = me->excel->zif_excel_book_vba_project~codename_pr ).
      lo_element ?= lo_iterator->get_next( ).
    ENDWHILE.

**********************************************************************
* STEP 3: Create xstring stream
    CLEAR ep_content.
    lo_ixml = cl_ixml=>create( ).
    lo_streamfactory = lo_ixml->create_stream_factory( ).
    lo_ostream = lo_streamfactory->create_ostream_xstring( string = ep_content ).
    lo_renderer = lo_ixml->create_renderer( ostream  = lo_ostream document = lo_document ).
    lo_renderer->render( ).
  ENDMETHOD.