  METHOD load_drawing_anchor.

    TYPES: BEGIN OF t_c_nv_pr,
             name TYPE string,
             id   TYPE string,
           END OF t_c_nv_pr.

    TYPES: BEGIN OF t_blip,
             cstate TYPE string,
             embed  TYPE string,
           END OF t_blip.

    TYPES: BEGIN OF t_chart,
             id TYPE string,
           END OF t_chart.

    TYPES: BEGIN OF t_ext,
             cx TYPE string,
             cy TYPE string,
           END OF t_ext.

    CONSTANTS: lc_xml_attr_true     TYPE string VALUE 'true',
               lc_xml_attr_true_int TYPE string VALUE '1'.
    CONSTANTS: lc_rel_chart TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
               lc_rel_image TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'.

    DATA: lo_drawing     TYPE REF TO zcl_excel_drawing,
          node           TYPE REF TO if_ixml_element,
          node2          TYPE REF TO if_ixml_element,
          node3          TYPE REF TO if_ixml_element,
          node4          TYPE REF TO if_ixml_element,

          ls_upper       TYPE zexcel_drawing_location,
          ls_lower       TYPE zexcel_drawing_location,
          ls_size        TYPE zexcel_drawing_size,
          ext            TYPE t_ext,
          lv_content     TYPE xstring,
          lv_relation_id TYPE string,
          lv_title       TYPE string,

          cnvpr          TYPE t_c_nv_pr,
          blip           TYPE t_blip,
          chart          TYPE t_chart,
          drawing_type   TYPE zexcel_drawing_type,

          rel_drawing    TYPE t_rel_drawing.

    node ?= io_anchor_element->find_from_name_ns( name = 'from' uri = namespace-xdr ).
    CHECK node IS NOT INITIAL.
    node2 ?= node->find_from_name_ns( name = 'col' uri = namespace-xdr ).
    ls_upper-col = node2->get_value( ).
    node2 ?= node->find_from_name_ns( name = 'row' uri = namespace-xdr ).
    ls_upper-row = node2->get_value( ).
    node2 ?= node->find_from_name_ns( name = 'colOff' uri = namespace-xdr ).
    ls_upper-col_offset = node2->get_value( ).
    node2 ?= node->find_from_name_ns( name = 'rowOff' uri = namespace-xdr ).
    ls_upper-row_offset = node2->get_value( ).

    node ?= io_anchor_element->find_from_name_ns( name = 'ext' uri = namespace-xdr ).
    IF node IS INITIAL.
      CLEAR ls_size.
    ELSE.
      me->fill_struct_from_attributes( EXPORTING ip_element = node CHANGING cp_structure = ext ).
      ls_size-width = ext-cx.
      ls_size-height = ext-cy.
      TRY.
          ls_size-width  = zcl_excel_drawing=>emu2pixel( ls_size-width ).
        CATCH cx_root.
      ENDTRY.
      TRY.
          ls_size-height = zcl_excel_drawing=>emu2pixel( ls_size-height ).
        CATCH cx_root.
      ENDTRY.
    ENDIF.

    node ?= io_anchor_element->find_from_name_ns( name = 'to' uri = namespace-xdr ).
    IF node IS INITIAL.
      CLEAR ls_lower.
    ELSE.
      node2 ?= node->find_from_name_ns( name = 'col' uri = namespace-xdr ).
      ls_lower-col = node2->get_value( ).
      node2 ?= node->find_from_name_ns( name = 'row' uri = namespace-xdr ).
      ls_lower-row = node2->get_value( ).
      node2 ?= node->find_from_name_ns( name = 'colOff' uri = namespace-xdr ).
      ls_lower-col_offset = node2->get_value( ).
      node2 ?= node->find_from_name_ns( name = 'rowOff' uri = namespace-xdr ).
      ls_lower-row_offset = node2->get_value( ).
    ENDIF.

    node ?= io_anchor_element->find_from_name_ns( name = 'pic' uri = namespace-xdr ).
    IF node IS NOT INITIAL.
      node2 ?= node->find_from_name_ns( name = 'nvPicPr' uri = namespace-xdr ).
      CHECK node2 IS NOT INITIAL.
      node3 ?= node2->find_from_name_ns( name = 'cNvPr' uri = namespace-xdr ).
      CHECK node3 IS NOT INITIAL.
      me->fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = cnvpr ).
      lv_title = cnvpr-name.

      node2 ?= node->find_from_name_ns( name = 'blipFill' uri = namespace-xdr ).
      CHECK node2 IS NOT INITIAL.
      node3 ?= node2->find_from_name_ns( name = 'blip' uri = namespace-a ).
      CHECK node3 IS NOT INITIAL.
      me->fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = blip ).
      lv_relation_id = blip-embed.

      drawing_type = zcl_excel_drawing=>type_image.
    ENDIF.

    node ?= io_anchor_element->find_from_name_ns( name = 'graphicFrame' uri = namespace-xdr ).
    IF node IS NOT INITIAL.
      node2 ?= node->find_from_name_ns( name = 'nvGraphicFramePr' uri = namespace-xdr ).
      CHECK node2 IS NOT INITIAL.
      node3 ?= node2->find_from_name_ns( name = 'cNvPr' uri = namespace-xdr ).
      CHECK node3 IS NOT INITIAL.
      me->fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = cnvpr ).
      lv_title = cnvpr-name.

      node2 ?= node->find_from_name_ns( name = 'graphic' uri = namespace-a ).
      CHECK node2 IS NOT INITIAL.
      node3 ?= node2->find_from_name_ns( name = 'graphicData' uri = namespace-a ).
      CHECK node3 IS NOT INITIAL.
      node4 ?= node2->find_from_name_ns( name = 'chart' uri = namespace-c ).
      CHECK node4 IS NOT INITIAL.
      me->fill_struct_from_attributes( EXPORTING ip_element = node4 CHANGING cp_structure = chart ).
      lv_relation_id = chart-id.

      drawing_type = zcl_excel_drawing=>type_chart.
    ENDIF.

    lo_drawing = io_worksheet->excel->add_new_drawing(
                      ip_type  = drawing_type
                      ip_title = lv_title ).
    io_worksheet->add_drawing( lo_drawing ).

    lo_drawing->set_position2(
      EXPORTING
        ip_from   = ls_upper
        ip_to     = ls_lower ).

    READ TABLE it_related_drawings INTO rel_drawing
          WITH KEY id = lv_relation_id.

    lo_drawing->set_media(
      EXPORTING
        ip_media = rel_drawing-content
        ip_media_type = rel_drawing-file_ext
        ip_width = ls_size-width
        ip_height = ls_size-height ).

    IF drawing_type = zcl_excel_drawing=>type_chart.
*  Begin fix for Issue #551
      DATA: lo_tmp_node_2                TYPE REF TO if_ixml_element.
      lo_tmp_node_2 ?= rel_drawing-content_xml->find_from_name_ns( name = 'pieChart' uri = namespace-c ).
      IF lo_tmp_node_2 IS NOT INITIAL.
        lo_drawing->graph_type = zcl_excel_drawing=>c_graph_pie.
      ELSE.
        lo_tmp_node_2 ?= rel_drawing-content_xml->find_from_name_ns( name = 'barChart' uri = namespace-c ).
        IF lo_tmp_node_2 IS NOT INITIAL.
          lo_drawing->graph_type = zcl_excel_drawing=>c_graph_bars.
        ELSE.
          lo_tmp_node_2 ?= rel_drawing-content_xml->find_from_name_ns( name = 'lineChart' uri = namespace-c ).
          IF lo_tmp_node_2 IS NOT INITIAL.
            lo_drawing->graph_type = zcl_excel_drawing=>c_graph_line.
          ENDIF.
        ENDIF.
      ENDIF.
* End fix for issue #551
      "-------------Added by Alessandro Iannacci - Should load chart attributes
      lo_drawing->load_chart_attributes( rel_drawing-content_xml ).
    ENDIF.

  ENDMETHOD.