  METHOD create_xl_drawings.


** Constant node name
    CONSTANTS: lc_xml_node_wsdr   TYPE string VALUE 'xdr:wsDr',
               lc_xml_node_ns_xdr TYPE string VALUE 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
               lc_xml_node_ns_a   TYPE string VALUE 'http://schemas.openxmlformats.org/drawingml/2006/main'.

    DATA: lo_document           TYPE REF TO if_ixml_document,
          lo_element_root       TYPE REF TO if_ixml_element,
          lo_element_cellanchor TYPE REF TO if_ixml_element,
          lo_iterator           TYPE REF TO zcl_excel_collection_iterator,
          lo_drawings           TYPE REF TO zcl_excel_drawings,
          lo_drawing            TYPE REF TO zcl_excel_drawing.
    DATA: lv_rel_id            TYPE i.



**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

***********************************************************************
* STEP 3: Create main node relationships
    lo_element_root  = lo_document->create_simple_element( name   = lc_xml_node_wsdr
                                                           parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:xdr'
                                       value = lc_xml_node_ns_xdr ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:a'
                                       value = lc_xml_node_ns_a ).

**********************************************************************
* STEP 4: Create drawings

    CLEAR: lv_rel_id.

    lo_drawings = io_worksheet->get_drawings( ).

    lo_iterator = lo_drawings->get_iterator( ).
    WHILE lo_iterator->has_next( ) EQ abap_true.
      lo_drawing ?= lo_iterator->get_next( ).

      ADD 1 TO lv_rel_id.
      lo_element_cellanchor = me->create_xl_drawing_anchor(
              io_drawing    = lo_drawing
              io_document   = lo_document
              ip_index      = lv_rel_id ).

      lo_element_root->append_child( new_child = lo_element_cellanchor ).

    ENDWHILE.

**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).

  ENDMETHOD.