  METHOD create_relationships.


** Constant node name
    DATA: lc_xml_node_relationships TYPE string VALUE 'Relationships',
          lc_xml_node_relationship  TYPE string VALUE 'Relationship',
          " Node attributes
          lc_xml_attr_id            TYPE string VALUE 'Id',
          lc_xml_attr_type          TYPE string VALUE 'Type',
          lc_xml_attr_target        TYPE string VALUE 'Target',
          " Node namespace
          lc_xml_node_rels_ns       TYPE string VALUE 'http://schemas.openxmlformats.org/package/2006/relationships',
          " Node id
          lc_xml_node_rid1_id       TYPE string VALUE 'rId1',
          lc_xml_node_rid2_id       TYPE string VALUE 'rId2',
          lc_xml_node_rid3_id       TYPE string VALUE 'rId3',
          " Node type
          lc_xml_node_rid1_tp       TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
          lc_xml_node_rid2_tp       TYPE string VALUE 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
          lc_xml_node_rid3_tp       TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
          " Node target
          lc_xml_node_rid1_tg       TYPE string VALUE 'xl/workbook.xml',
          lc_xml_node_rid2_tg       TYPE string VALUE 'docProps/core.xml',
          lc_xml_node_rid3_tg       TYPE string VALUE 'docProps/app.xml'.

    DATA: lo_document     TYPE REF TO if_ixml_document,
          lo_element_root TYPE REF TO if_ixml_element,
          lo_element      TYPE REF TO if_ixml_element.

**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

**********************************************************************
* STEP 3: Create main node relationships
    lo_element_root  = lo_document->create_simple_element( name   = lc_xml_node_relationships
                                                           parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns'
                                       value = lc_xml_node_rels_ns ).

**********************************************************************
* STEP 4: Create subnodes
    " Theme node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                  value = lc_xml_node_rid3_id ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                  value = lc_xml_node_rid3_tp ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                  value = lc_xml_node_rid3_tg ).
    lo_element_root->append_child( new_child = lo_element ).

    " Styles node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                  value = lc_xml_node_rid2_id ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                  value = lc_xml_node_rid2_tp ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                  value = lc_xml_node_rid2_tg ).
    lo_element_root->append_child( new_child = lo_element ).

    " rels node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                  value = lc_xml_node_rid1_id ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                  value = lc_xml_node_rid1_tp ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                  value = lc_xml_node_rid1_tg ).
    lo_element_root->append_child( new_child = lo_element ).

**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).
  ENDMETHOD.