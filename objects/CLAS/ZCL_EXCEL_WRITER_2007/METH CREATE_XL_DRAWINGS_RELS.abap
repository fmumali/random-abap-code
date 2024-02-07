  METHOD create_xl_drawings_rels.

** Constant node name
    DATA: lc_xml_node_relationships TYPE string VALUE 'Relationships',
          lc_xml_node_relationship  TYPE string VALUE 'Relationship',
          " Node attributes
          lc_xml_attr_id            TYPE string VALUE 'Id',
          lc_xml_attr_type          TYPE string VALUE 'Type',
          lc_xml_attr_target        TYPE string VALUE 'Target',
          " Node namespace
          lc_xml_node_rels_ns       TYPE string VALUE 'http://schemas.openxmlformats.org/package/2006/relationships',
          lc_xml_node_rid_image_tp  TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
          lc_xml_node_rid_chart_tp  TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'.

    DATA: lo_document     TYPE REF TO if_ixml_document,
          lo_element_root TYPE REF TO if_ixml_element,
          lo_element      TYPE REF TO if_ixml_element,
          lo_iterator     TYPE REF TO zcl_excel_collection_iterator,
          lo_drawings     TYPE REF TO zcl_excel_drawings,
          lo_drawing      TYPE REF TO zcl_excel_drawing.

    DATA: lv_value   TYPE string,
          lv_counter TYPE i.

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

    " Add sheet Relationship nodes here
    lv_counter = 0.
    lo_drawings = io_worksheet->get_drawings( ).
    lo_iterator = lo_drawings->get_iterator( ).
    WHILE lo_iterator->has_next( ) EQ abap_true.
      lo_drawing ?= lo_iterator->get_next( ).
      ADD 1 TO lv_counter.

      lv_value = lv_counter.
      CONDENSE lv_value.
      CONCATENATE 'rId' lv_value INTO lv_value.

      lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                   parent = lo_document ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                    value = lv_value ).

      lv_value = lo_drawing->get_media_name( ).
      CASE lo_drawing->get_type( ).
        WHEN zcl_excel_drawing=>type_image.
          CONCATENATE '../media/' lv_value INTO lv_value.
          lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                        value = lc_xml_node_rid_image_tp ).

        WHEN zcl_excel_drawing=>type_chart.
          CONCATENATE '../charts/' lv_value INTO lv_value.
          lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                        value = lc_xml_node_rid_chart_tp ).

      ENDCASE.
      lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                    value = lv_value ).
      lo_element_root->append_child( new_child = lo_element ).
    ENDWHILE.


**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).

  ENDMETHOD.