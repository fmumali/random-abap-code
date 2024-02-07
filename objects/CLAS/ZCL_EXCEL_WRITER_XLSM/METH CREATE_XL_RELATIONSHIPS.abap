  METHOD create_xl_relationships.

** Constant node name
    DATA: lc_xml_node_relationship TYPE string VALUE 'Relationship',
          " Node attributes
          lc_xml_attr_id           TYPE string VALUE 'Id',
          lc_xml_attr_type         TYPE string VALUE 'Type',
          lc_xml_attr_target       TYPE string VALUE 'Target',
          " Node id
          lc_xml_node_ridx_id      TYPE string VALUE 'rId#',
          " Node type
          lc_xml_node_rid_vba_tp   TYPE string VALUE 'http://schemas.microsoft.com/office/2006/relationships/vbaProject',
          " Node target
          lc_xml_node_rid_vba_tg   TYPE string VALUE 'vbaProject.bin'.

    DATA: lo_ixml          TYPE REF TO if_ixml,
          lo_document      TYPE REF TO if_ixml_document,
          lo_document_xml  TYPE REF TO cl_xml_document,
          lo_element_root  TYPE REF TO if_ixml_node,
          lo_element       TYPE REF TO if_ixml_element,
          lo_streamfactory TYPE REF TO if_ixml_stream_factory,
          lo_ostream       TYPE REF TO if_ixml_ostream,
          lo_renderer      TYPE REF TO if_ixml_renderer.

    DATA: lv_xml_node_ridx_id TYPE string,
          lv_size             TYPE i,
          lv_syindex(2)       TYPE c.

**********************************************************************
* STEP 3: Create standard relationship
    ep_content = super->create_xl_relationships( ).

**********************************************************************
* STEP 2: modify XML adding the vbaProject relation

    CREATE OBJECT lo_document_xml.
    lo_document_xml->parse_xstring( ep_content ).

    lo_document ?= lo_document_xml->m_document.
    lo_element_root = lo_document->if_ixml_node~get_first_child( ).


    lv_size = excel->get_worksheets_size( ).

    " Relationship node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                     parent = lo_document ).
    ADD 4 TO lv_size.
    lv_syindex = lv_size.
    SHIFT lv_syindex RIGHT DELETING TRAILING space.
    SHIFT lv_syindex LEFT DELETING LEADING space.
    lv_xml_node_ridx_id = lc_xml_node_ridx_id.
    REPLACE ALL OCCURRENCES OF '#' IN lv_xml_node_ridx_id WITH lv_syindex.
    lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                  value = lv_xml_node_ridx_id ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                  value = lc_xml_node_rid_vba_tp ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                  value = lc_xml_node_rid_vba_tg ).
    lo_element_root->append_child( new_child = lo_element ).

**********************************************************************
* STEP 3: Create xstring stream
    CLEAR ep_content.
    lo_ixml = cl_ixml=>create( ).
    lo_streamfactory = lo_ixml->create_stream_factory( ).
    lo_ostream = lo_streamfactory->create_ostream_xstring( string = ep_content ).
    lo_renderer = lo_ixml->create_renderer( ostream  = lo_ostream document = lo_document ).
    lo_renderer->render( ).

  ENDMETHOD.