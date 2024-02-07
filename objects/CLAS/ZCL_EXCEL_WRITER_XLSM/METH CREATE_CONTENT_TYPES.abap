  METHOD create_content_types.
** Constant node name
    DATA: lc_xml_node_workb_ct    TYPE string VALUE 'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
          lc_xml_node_default     TYPE string VALUE 'Default',
          " Node attributes
          lc_xml_attr_partname    TYPE string VALUE 'PartName',
          lc_xml_attr_extension   TYPE string VALUE 'Extension',
          lc_xml_attr_contenttype TYPE string VALUE 'ContentType',
          lc_xml_node_workb_pn    TYPE string VALUE '/xl/workbook.xml',
          lc_xml_node_bin_ext     TYPE string VALUE 'bin',
          lc_xml_node_bin_ct      TYPE string VALUE 'application/vnd.ms-office.vbaProject'.


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

    DATA: lv_contenttype TYPE string.

**********************************************************************
* STEP 3: Create standard contentType
    ep_content = super->create_content_types( ).

**********************************************************************
* STEP 2: modify XML adding the extension bin definition

    CREATE OBJECT lo_document_xml.
    lo_document_xml->parse_xstring( ep_content ).

    lo_document ?= lo_document_xml->m_document.
    lo_element_root = lo_document->if_ixml_node~get_first_child( ).

    " extension node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_default
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_extension
                                  value = lc_xml_node_bin_ext ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                  value = lc_xml_node_bin_ct ).
    lo_element_root->append_child( new_child = lo_element ).

**********************************************************************
* STEP 3: modify XML changing the contentType of node Override /xl/workbook.xml

    lo_collection = lo_document->get_elements_by_tag_name( 'Override' ).
    lo_iterator = lo_collection->create_iterator( ).
    lo_element ?= lo_iterator->get_next( ).
    WHILE lo_element IS BOUND.
      lv_contenttype = lo_element->get_attribute_ns( lc_xml_attr_partname ).
      IF lv_contenttype EQ lc_xml_node_workb_pn.
        lo_element->remove_attribute_ns( lc_xml_attr_contenttype ).
        lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                      value = lc_xml_node_workb_ct ).
        EXIT.
      ENDIF.
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