  METHOD create_xl_drawing_anchor.

** Constant node name
    CONSTANTS: lc_xml_node_onecellanchor     TYPE string VALUE 'xdr:oneCellAnchor',
               lc_xml_node_twocellanchor     TYPE string VALUE 'xdr:twoCellAnchor',
               lc_xml_node_from              TYPE string VALUE 'xdr:from',
               lc_xml_node_to                TYPE string VALUE 'xdr:to',
               lc_xml_node_pic               TYPE string VALUE 'xdr:pic',
               lc_xml_node_ext               TYPE string VALUE 'xdr:ext',
               lc_xml_node_clientdata        TYPE string VALUE 'xdr:clientData',

               lc_xml_node_col               TYPE string VALUE 'xdr:col',
               lc_xml_node_coloff            TYPE string VALUE 'xdr:colOff',
               lc_xml_node_row               TYPE string VALUE 'xdr:row',
               lc_xml_node_rowoff            TYPE string VALUE 'xdr:rowOff',

               lc_xml_node_nvpicpr           TYPE string VALUE 'xdr:nvPicPr',
               lc_xml_node_cnvpr             TYPE string VALUE 'xdr:cNvPr',
               lc_xml_node_cnvpicpr          TYPE string VALUE 'xdr:cNvPicPr',
               lc_xml_node_piclocks          TYPE string VALUE 'a:picLocks',

               lc_xml_node_sppr              TYPE string VALUE 'xdr:spPr',
               lc_xml_node_apgeom            TYPE string VALUE 'a:prstGeom',
               lc_xml_node_aavlst            TYPE string VALUE 'a:avLst',

               lc_xml_node_graphicframe      TYPE string VALUE 'xdr:graphicFrame',
               lc_xml_node_nvgraphicframepr  TYPE string VALUE 'xdr:nvGraphicFramePr',
               lc_xml_node_cnvgraphicframepr TYPE string VALUE 'xdr:cNvGraphicFramePr',
               lc_xml_node_graphicframelocks TYPE string VALUE 'a:graphicFrameLocks',
               lc_xml_node_xfrm              TYPE string VALUE 'xdr:xfrm',
               lc_xml_node_aoff              TYPE string VALUE 'a:off',
               lc_xml_node_aext              TYPE string VALUE 'a:ext',
               lc_xml_node_agraphic          TYPE string VALUE 'a:graphic',
               lc_xml_node_agraphicdata      TYPE string VALUE 'a:graphicData',

               lc_xml_node_ns_c              TYPE string VALUE 'http://schemas.openxmlformats.org/drawingml/2006/chart',
               lc_xml_node_cchart            TYPE string VALUE 'c:chart',

               lc_xml_node_blipfill          TYPE string VALUE 'xdr:blipFill',
               lc_xml_node_ablip             TYPE string VALUE 'a:blip',
               lc_xml_node_astretch          TYPE string VALUE 'a:stretch',
               lc_xml_node_ns_r              TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'.

    DATA: lo_element_graphicframe TYPE REF TO if_ixml_element,
          lo_element              TYPE REF TO if_ixml_element,
          lo_element2             TYPE REF TO if_ixml_element,
          lo_element3             TYPE REF TO if_ixml_element,
          lo_element_from         TYPE REF TO if_ixml_element,
          lo_element_to           TYPE REF TO if_ixml_element,
          lo_element_ext          TYPE REF TO if_ixml_element,
          lo_element_pic          TYPE REF TO if_ixml_element,
          lo_element_clientdata   TYPE REF TO if_ixml_element,
          ls_position             TYPE zexcel_drawing_position,
          lv_col                  TYPE string, " zexcel_cell_column,
          lv_row                  TYPE string, " zexcel_cell_row.
          lv_col_offset           TYPE string,
          lv_row_offset           TYPE string,
          lv_value                TYPE string.

    ls_position = io_drawing->get_position( ).

    IF ls_position-anchor = 'ONE'.
      ep_anchor = io_document->create_simple_element( name   = lc_xml_node_onecellanchor
                                                                  parent = io_document ).
    ELSE.
      ep_anchor = io_document->create_simple_element( name   = lc_xml_node_twocellanchor
                                                                  parent = io_document ).
    ENDIF.

*   from cell ******************************
    lo_element_from = io_document->create_simple_element( name   = lc_xml_node_from
                                                          parent = io_document ).

    lv_col = ls_position-from-col.
    lv_row = ls_position-from-row.
    lv_col_offset = ls_position-from-col_offset.
    lv_row_offset = ls_position-from-row_offset.
    CONDENSE lv_col NO-GAPS.
    CONDENSE lv_row NO-GAPS.
    CONDENSE lv_col_offset NO-GAPS.
    CONDENSE lv_row_offset NO-GAPS.

    lo_element = io_document->create_simple_element( name = lc_xml_node_col
                                                     parent = io_document ).
    lo_element->set_value( value = lv_col ).
    lo_element_from->append_child( new_child = lo_element ).

    lo_element = io_document->create_simple_element( name = lc_xml_node_coloff
                                                     parent = io_document ).
    lo_element->set_value( value = lv_col_offset ).
    lo_element_from->append_child( new_child = lo_element ).

    lo_element = io_document->create_simple_element( name = lc_xml_node_row
                                                     parent = io_document ).
    lo_element->set_value( value = lv_row ).
    lo_element_from->append_child( new_child = lo_element ).

    lo_element = io_document->create_simple_element( name = lc_xml_node_rowoff
                                                     parent = io_document ).
    lo_element->set_value( value = lv_row_offset ).
    lo_element_from->append_child( new_child = lo_element ).
    ep_anchor->append_child( new_child = lo_element_from ).

    IF ls_position-anchor = 'ONE'.

*   ext ******************************
      lo_element_ext = io_document->create_simple_element( name   = lc_xml_node_ext
                                                           parent = io_document ).

      lv_value = io_drawing->get_width_emu_str( ).
      lo_element_ext->set_attribute_ns( name  = 'cx'
                                     value = lv_value ).
      lv_value = io_drawing->get_height_emu_str( ).
      lo_element_ext->set_attribute_ns( name  = 'cy'
                                     value = lv_value ).
      ep_anchor->append_child( new_child = lo_element_ext ).

    ELSEIF ls_position-anchor = 'TWO'.

*   to cell ******************************
      lo_element_to = io_document->create_simple_element( name   = lc_xml_node_to
                                                          parent = io_document ).

      lv_col = ls_position-to-col.
      lv_row = ls_position-to-row.
      lv_col_offset = ls_position-to-col_offset.
      lv_row_offset = ls_position-to-row_offset.
      CONDENSE lv_col NO-GAPS.
      CONDENSE lv_row NO-GAPS.
      CONDENSE lv_col_offset NO-GAPS.
      CONDENSE lv_row_offset NO-GAPS.

      lo_element = io_document->create_simple_element( name = lc_xml_node_col
                                                       parent = io_document ).
      lo_element->set_value( value = lv_col ).
      lo_element_to->append_child( new_child = lo_element ).

      lo_element = io_document->create_simple_element( name = lc_xml_node_coloff
                                                       parent = io_document ).
      lo_element->set_value( value = lv_col_offset ).
      lo_element_to->append_child( new_child = lo_element ).

      lo_element = io_document->create_simple_element( name = lc_xml_node_row
                                                       parent = io_document ).
      lo_element->set_value( value = lv_row ).
      lo_element_to->append_child( new_child = lo_element ).

      lo_element = io_document->create_simple_element( name = lc_xml_node_rowoff
                                                       parent = io_document ).
      lo_element->set_value( value = lv_row_offset ).
      lo_element_to->append_child( new_child = lo_element ).
      ep_anchor->append_child( new_child = lo_element_to ).

    ENDIF.

    CASE io_drawing->get_type( ).
      WHEN zcl_excel_drawing=>type_image.
*     pic **********************************
        lo_element_pic = io_document->create_simple_element( name   = lc_xml_node_pic
                                                             parent = io_document ).
*     nvPicPr
        lo_element  = io_document->create_simple_element( name = lc_xml_node_nvpicpr
                                                          parent = io_document ).
*     cNvPr
        lo_element2 = io_document->create_simple_element( name = lc_xml_node_cnvpr
                                                          parent = io_document ).
        lv_value = sy-index.
        CONDENSE lv_value.
        lo_element2->set_attribute_ns( name  = 'id'
                                       value = lv_value ).
        lo_element2->set_attribute_ns( name  = 'name'
                                       value = io_drawing->title ).
        lo_element->append_child( new_child = lo_element2 ).

*     cNvPicPr
        lo_element2 = io_document->create_simple_element( name = lc_xml_node_cnvpicpr
                                                          parent = io_document ).

*     picLocks
        lo_element3 = io_document->create_simple_element( name = lc_xml_node_piclocks
                                                          parent = io_document ).
        lo_element3->set_attribute_ns( name  = 'noChangeAspect'
                                       value = '1' ).

        lo_element2->append_child( new_child = lo_element3 ).
        lo_element->append_child( new_child = lo_element2 ).
        lo_element_pic->append_child( new_child = lo_element ).

*     blipFill
        lv_value = ip_index.
        CONDENSE lv_value.
        CONCATENATE 'rId' lv_value INTO lv_value.

        lo_element  = io_document->create_simple_element( name = lc_xml_node_blipfill
                                                          parent = io_document ).
        lo_element2 = io_document->create_simple_element( name = lc_xml_node_ablip
                                                          parent = io_document ).
        lo_element2->set_attribute_ns( name  = 'xmlns:r'
                                       value = lc_xml_node_ns_r ).
        lo_element2->set_attribute_ns( name  = 'r:embed'
                                       value = lv_value ).
        lo_element->append_child( new_child = lo_element2 ).

        lo_element2  = io_document->create_simple_element( name = lc_xml_node_astretch
                                                          parent = io_document ).
        lo_element->append_child( new_child = lo_element2 ).

        lo_element_pic->append_child( new_child = lo_element ).

*     spPr
        lo_element  = io_document->create_simple_element( name = lc_xml_node_sppr
                                                          parent = io_document ).

        lo_element2 = io_document->create_simple_element( name = lc_xml_node_apgeom
                                                          parent = io_document ).
        lo_element2->set_attribute_ns( name  = 'prst'
                                       value = 'rect' ).
        lo_element3 = io_document->create_simple_element( name = lc_xml_node_aavlst
                                                          parent = io_document ).
        lo_element2->append_child( new_child = lo_element3 ).
        lo_element->append_child( new_child = lo_element2 ).

        lo_element_pic->append_child( new_child = lo_element ).
        ep_anchor->append_child( new_child = lo_element_pic ).
      WHEN zcl_excel_drawing=>type_chart.
*     graphicFrame **********************************
        lo_element_graphicframe = io_document->create_simple_element( name   = lc_xml_node_graphicframe
                                                             parent = io_document ).
*     nvGraphicFramePr
        lo_element  = io_document->create_simple_element( name = lc_xml_node_nvgraphicframepr
                                                          parent = io_document ).
*     cNvPr
        lo_element2 = io_document->create_simple_element( name = lc_xml_node_cnvpr
                                                          parent = io_document ).
        lv_value = sy-index.
        CONDENSE lv_value.
        lo_element2->set_attribute_ns( name  = 'id'
                                       value = lv_value ).
        lo_element2->set_attribute_ns( name  = 'name'
                                       value = io_drawing->title ).
        lo_element->append_child( new_child = lo_element2 ).
*     cNvGraphicFramePr
        lo_element2 = io_document->create_simple_element( name = lc_xml_node_cnvgraphicframepr
                                                          parent = io_document ).
        lo_element3 = io_document->create_simple_element( name = lc_xml_node_graphicframelocks
                                                          parent = io_document ).
        lo_element2->append_child( new_child = lo_element3 ).
        lo_element->append_child( new_child = lo_element2 ).
        lo_element_graphicframe->append_child( new_child = lo_element ).

*     xfrm
        lo_element  = io_document->create_simple_element( name = lc_xml_node_xfrm
                                                          parent = io_document ).
*     off
        lo_element2 = io_document->create_simple_element( name = lc_xml_node_aoff
                                                          parent = io_document ).
        lo_element2->set_attribute_ns( name  = 'y' value = '0' ).
        lo_element2->set_attribute_ns( name  = 'x' value = '0' ).
        lo_element->append_child( new_child = lo_element2 ).
*     ext
        lo_element2 = io_document->create_simple_element( name = lc_xml_node_aext
                                                          parent = io_document ).
        lo_element2->set_attribute_ns( name  = 'cy' value = '0' ).
        lo_element2->set_attribute_ns( name  = 'cx' value = '0' ).
        lo_element->append_child( new_child = lo_element2 ).
        lo_element_graphicframe->append_child( new_child = lo_element ).

*     graphic
        lo_element  = io_document->create_simple_element( name = lc_xml_node_agraphic
                                                          parent = io_document ).
*     graphicData
        lo_element2 = io_document->create_simple_element( name = lc_xml_node_agraphicdata
                                                          parent = io_document ).
        lo_element2->set_attribute_ns( name  = 'uri' value = lc_xml_node_ns_c ).

*     chart
        lo_element3 = io_document->create_simple_element( name = lc_xml_node_cchart
                                                          parent = io_document ).

        lo_element3->set_attribute_ns( name  = 'xmlns:r'
                                       value = lc_xml_node_ns_r ).
        lo_element3->set_attribute_ns( name  = 'xmlns:c'
                                       value = lc_xml_node_ns_c ).

        lv_value = ip_index.
        CONDENSE lv_value.
        CONCATENATE 'rId' lv_value INTO lv_value.
        lo_element3->set_attribute_ns( name  = 'r:id'
                                       value = lv_value ).
        lo_element2->append_child( new_child = lo_element3 ).
        lo_element->append_child( new_child = lo_element2 ).
        lo_element_graphicframe->append_child( new_child = lo_element ).
        ep_anchor->append_child( new_child = lo_element_graphicframe ).

    ENDCASE.

*   client data ***************************
    lo_element_clientdata = io_document->create_simple_element( name   = lc_xml_node_clientdata
                                                                parent = io_document ).
    ep_anchor->append_child( new_child = lo_element_clientdata ).

  ENDMETHOD.