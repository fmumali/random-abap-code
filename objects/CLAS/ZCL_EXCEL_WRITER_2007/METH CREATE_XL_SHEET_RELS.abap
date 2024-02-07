  METHOD create_xl_sheet_rels.


** Constant node name
    DATA: lc_xml_node_relationships      TYPE string VALUE 'Relationships',
          lc_xml_node_relationship       TYPE string VALUE 'Relationship',
          " Node attributes
          lc_xml_attr_id                 TYPE string VALUE 'Id',
          lc_xml_attr_type               TYPE string VALUE 'Type',
          lc_xml_attr_target             TYPE string VALUE 'Target',
          lc_xml_attr_target_mode        TYPE string VALUE 'TargetMode',
          lc_xml_val_external            TYPE string VALUE 'External',
          " Node namespace
          lc_xml_node_rels_ns            TYPE string VALUE 'http://schemas.openxmlformats.org/package/2006/relationships',
          lc_xml_node_rid_table_tp       TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table',
          lc_xml_node_rid_printer_tp     TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings',
          lc_xml_node_rid_drawing_tp     TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
          lc_xml_node_rid_comment_tp     TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',        " (+) Issue #180
          lc_xml_node_rid_drawing_cmt_tp TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing',      " (+) Issue #180
          lc_xml_node_rid_link_tp        TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'.

    DATA: lo_document     TYPE REF TO if_ixml_document,
          lo_element_root TYPE REF TO if_ixml_element,
          lo_element      TYPE REF TO if_ixml_element,
          lo_iterator     TYPE REF TO zcl_excel_collection_iterator,
          lo_table        TYPE REF TO zcl_excel_table,
          lo_link         TYPE REF TO zcl_excel_hyperlink.

    DATA: lv_value         TYPE string,
          lv_relation_id   TYPE i,
          lv_index_str     TYPE string,
          lv_comment_index TYPE i.

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
    lv_relation_id = 0.
    lo_iterator = io_worksheet->get_hyperlinks_iterator( ).
    WHILE lo_iterator->has_next( ) EQ abap_true.
      lo_link ?= lo_iterator->get_next( ).
      CHECK lo_link->is_internal( ) = abap_false.  " issue #340 - don't put internal links here
      ADD 1 TO lv_relation_id.

      lv_value = lv_relation_id.
      CONDENSE lv_value.
      CONCATENATE 'rId' lv_value INTO lv_value.

      lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                       parent = lo_document ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                    value = lv_value ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                    value = lc_xml_node_rid_link_tp ).

      lv_value = lo_link->get_url( ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                    value = lv_value ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_target_mode
                                    value = lc_xml_val_external ).
      lo_element_root->append_child( new_child = lo_element ).
    ENDWHILE.

* drawing
    DATA: lo_drawings TYPE REF TO zcl_excel_drawings.

    lo_drawings = io_worksheet->get_drawings( ).
    IF lo_drawings->is_empty( ) = abap_false.
      lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                       parent = lo_document ).
      ADD 1 TO lv_relation_id.

      lv_value = lv_relation_id.
      CONDENSE lv_value.
      CONCATENATE 'rId' lv_value INTO lv_value.
      lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                    value = lv_value ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                    value = lc_xml_node_rid_drawing_tp ).

      lv_index_str = iv_drawing_index.
      CONDENSE lv_index_str NO-GAPS.
      lv_value = me->c_xl_drawings.
      REPLACE 'xl' WITH '..' INTO lv_value.
      REPLACE '#' WITH lv_index_str INTO lv_value.
      lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                value = lv_value ).
      lo_element_root->append_child( new_child = lo_element ).
    ENDIF.

* Begin - Add - Issue #180
    DATA: lo_comments  TYPE REF TO zcl_excel_comments.

    lv_comment_index = iv_comment_index.

    lo_comments = io_worksheet->get_comments( ).
    IF lo_comments->is_empty( ) = abap_false.
      " Drawing for comment
      lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                       parent = lo_document ).

      ADD 1 TO lv_relation_id.
      ADD 1 TO lv_comment_index.

      lv_value = lv_relation_id.
      CONDENSE lv_value.
      CONCATENATE 'rId' lv_value INTO lv_value.
      lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                    value = lv_value ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                    value = lc_xml_node_rid_drawing_cmt_tp ).

      lv_index_str = iv_comment_index.
      CONDENSE lv_index_str NO-GAPS.
      lv_value = me->cl_xl_drawing_for_comments.
      REPLACE 'xl' WITH '..' INTO lv_value.
      REPLACE '#' WITH lv_index_str INTO lv_value.
      lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                    value = lv_value ).
      lo_element_root->append_child( new_child = lo_element ).

      " Comment
      lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                       parent = lo_document ).
      ADD 1 TO lv_relation_id.

      lv_value = lv_relation_id.
      CONDENSE lv_value.
      CONCATENATE 'rId' lv_value INTO lv_value.
      lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                    value = lv_value ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                    value = lc_xml_node_rid_comment_tp ).

      lv_index_str = iv_comment_index.
      CONDENSE lv_index_str NO-GAPS.
      lv_value = me->c_xl_comments.
      REPLACE 'xl' WITH '..' INTO lv_value.
      REPLACE '#' WITH lv_index_str INTO lv_value.
      lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                value = lv_value ).
      lo_element_root->append_child( new_child = lo_element ).
    ENDIF.
* End   - Add - Issue #180

**********************************************************************
* header footer image
    DATA: lt_drawings TYPE zexcel_t_drawings.
    lt_drawings = io_worksheet->get_header_footer_drawings( ).
    IF lines( lt_drawings ) > 0. "Header or footer image exist
      ADD 1 TO lv_relation_id.
      " Drawing for comment/header/footer
      lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                       parent = lo_document ).
      lv_value = lv_relation_id.
      CONDENSE lv_value.
      CONCATENATE 'rId' lv_value INTO lv_value.
      lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                    value = lv_value ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                    value = lc_xml_node_rid_drawing_cmt_tp ).

      lv_index_str = lv_comment_index.
      CONDENSE lv_index_str NO-GAPS.
      lv_value = me->cl_xl_drawing_for_comments.
      REPLACE 'xl' WITH '..' INTO lv_value.
      REPLACE '#' WITH lv_index_str INTO lv_value.
      lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                    value = lv_value ).
      lo_element_root->append_child( new_child = lo_element ).
    ENDIF.
*** End Header Footer
**********************************************************************


    lo_iterator = io_worksheet->get_tables_iterator( ).
    WHILE lo_iterator->has_next( ) EQ abap_true.
      lo_table ?= lo_iterator->get_next( ).
      ADD 1 TO lv_relation_id.

      lv_value = lv_relation_id.
      CONDENSE lv_value.
      CONCATENATE 'rId' lv_value INTO lv_value.

      lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                       parent = lo_document ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                    value = lv_value ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                    value = lc_xml_node_rid_table_tp ).

      lv_value = lo_table->get_name( ).
      CONCATENATE '../tables/' lv_value '.xml' INTO lv_value.
      lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                value = lv_value ).
      lo_element_root->append_child( new_child = lo_element ).
    ENDWHILE.

**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).

  ENDMETHOD.