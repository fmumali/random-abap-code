  METHOD load_chart_attributes.
    DATA: node                TYPE REF TO if_ixml_element.
    DATA: node2               TYPE REF TO if_ixml_element.
    DATA: node3               TYPE REF TO if_ixml_element.
    DATA: node4               TYPE REF TO if_ixml_element.

    DATA lo_barchart TYPE REF TO zcl_excel_graph_bars.
    DATA lo_piechart TYPE REF TO zcl_excel_graph_pie.
    DATA lo_linechart TYPE REF TO zcl_excel_graph_line.

    TYPES: BEGIN OF t_prop,
             val          TYPE string,
             rtl          TYPE string,
             lang         TYPE string,
             formatcode   TYPE string,
             sourcelinked TYPE string,
           END OF t_prop.

    TYPES: BEGIN OF t_pagemargins,
             b      TYPE string,
             l      TYPE string,
             r      TYPE string,
             t      TYPE string,
             header TYPE string,
             footer TYPE string,
           END OF t_pagemargins.

    DATA ls_prop TYPE t_prop.
    DATA ls_pagemargins TYPE t_pagemargins.

    DATA lo_collection TYPE REF TO if_ixml_node_collection.
    DATA lo_node       TYPE REF TO if_ixml_node.
    DATA lo_iterator   TYPE REF TO if_ixml_node_iterator.
    DATA lv_idx        TYPE i.
    DATA lv_order      TYPE i.
    DATA lv_invertifnegative      TYPE string.
    DATA lv_symbol      TYPE string.
    DATA lv_smooth      TYPE c.
    DATA lv_sername    TYPE string.
    DATA lv_label      TYPE string.
    DATA lv_value      TYPE string.
    DATA lv_axid       TYPE string.
    DATA lv_orientation TYPE string.
    DATA lv_delete TYPE string.
    DATA lv_axpos TYPE string.
    DATA lv_formatcode TYPE string.
    DATA lv_sourcelinked TYPE string.
    DATA lv_majortickmark TYPE string.
    DATA lv_minortickmark TYPE string.
    DATA lv_ticklblpos TYPE string.
    DATA lv_crossax TYPE string.
    DATA lv_crosses TYPE string.
    DATA lv_auto TYPE string.
    DATA lv_lblalgn TYPE string.
    DATA lv_lbloffset TYPE string.
    DATA lv_nomultilvllbl TYPE string.
    DATA lv_crossbetween TYPE string.

    node ?= ip_chart->if_ixml_node~get_first_child( ).
    CHECK node IS NOT INITIAL.

    CASE me->graph_type.
      WHEN c_graph_bars.
        CREATE OBJECT lo_barchart.
        me->graph = lo_barchart.
      WHEN c_graph_pie.
        CREATE OBJECT lo_piechart.
        me->graph = lo_piechart.
      WHEN c_graph_line.
        CREATE OBJECT lo_linechart.
        me->graph = lo_linechart.
      WHEN OTHERS.
    ENDCASE.

    "Fill properties
    node2 ?= node->find_from_name_ns( name = 'date1904' uri = namespace-c ).
    zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
    me->graph->ns_1904val = ls_prop-val.
    node2 ?= node->find_from_name_ns( name = 'lang' uri = namespace-c ).
    zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
    me->graph->ns_langval = ls_prop-val.
    node2 ?= node->find_from_name_ns( name = 'roundedCorners' uri = namespace-c ).
    zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
    me->graph->ns_roundedcornersval = ls_prop-val.

    "style
    node2 ?= node->find_from_name_ns( name = 'style' uri = namespace-c14 ).
    zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
    me->graph->ns_c14styleval = ls_prop-val.
    node2 ?= node->find_from_name_ns( name = 'style' uri = namespace-c ).
    zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
    me->graph->ns_styleval = ls_prop-val.
    "---------------------------Read graph properties
    "ADDED
    CLEAR node2.
    node2 ?= node->find_from_name_ns( name = 'title' uri = namespace-c ).
    IF node2 IS BOUND AND node2 IS NOT INITIAL.
      node3 ?= node2->find_from_name_ns( name = 't' uri = namespace-a ).
      me->graph->title = node3->get_value( ).
    ENDIF.
    "END

    node2 ?= node->find_from_name_ns( name = 'autoTitleDeleted' uri = namespace-c ).
    zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
    me->graph->ns_autotitledeletedval = ls_prop-val.

    "plotArea
    CASE me->graph_type.
      WHEN c_graph_bars.
        node2 ?= node->find_from_name_ns( name = 'barDir' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_barchart->ns_bardirval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'grouping' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_barchart->ns_groupingval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'varyColors' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_barchart->ns_varycolorsval = ls_prop-val.

        "Load series
        CALL METHOD node->get_elements_by_tag_name_ns
          EXPORTING
*           depth     = 0
            name = 'ser'
            uri  = namespace-c
          RECEIVING
            rval = lo_collection.
        CALL METHOD lo_collection->create_iterator
          RECEIVING
            rval = lo_iterator.
        lo_node = lo_iterator->get_next( ).
        IF lo_node IS BOUND.
          node2 ?= lo_node->query_interface( c_ixml_iid_element ).
        ENDIF.
        WHILE lo_node IS BOUND.
          node3 ?= node2->find_from_name_ns( name = 'idx' uri = namespace-c ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_idx = ls_prop-val.
          node3 ?= node2->find_from_name_ns( name = 'order' uri = namespace-c ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_order = ls_prop-val.
          node3 ?= node2->find_from_name_ns( name = 'invertIfNegative' uri = namespace-c ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_invertifnegative = ls_prop-val.
          node3 ?= node2->find_from_name_ns( name = 'v' uri = namespace-c ).
          IF node3 IS BOUND.
            lv_sername = node3->get_value( ).
          ENDIF.
          node3 ?= node2->find_from_name_ns( name = 'strRef' uri = namespace-c ).
          IF node3 IS BOUND.
            node4 ?= node3->find_from_name_ns( name = 'f' uri = namespace-c ).
            lv_label = node4->get_value( ).
          ENDIF.
          node3 ?= node2->find_from_name_ns( name = 'numRef' uri = namespace-c ).
          IF node3 IS BOUND.
            node4 ?= node3->find_from_name_ns( name = 'f' uri = namespace-c ).
            lv_value = node4->get_value( ).
          ENDIF.
          CALL METHOD lo_barchart->create_serie
            EXPORTING
              ip_idx              = lv_idx
              ip_order            = lv_order
              ip_invertifnegative = lv_invertifnegative
              ip_lbl              = lv_label
              ip_ref              = lv_value
              ip_sername          = lv_sername.
          lo_node = lo_iterator->get_next( ).
          IF lo_node IS BOUND.
            node2 ?= lo_node->query_interface( c_ixml_iid_element ).
          ENDIF.
        ENDWHILE.
        "note: numCache avoided
        node2 ?= node->find_from_name_ns( name = 'showLegendKey' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_barchart->ns_showlegendkeyval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'showVal' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_barchart->ns_showvalval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'showCatName' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_barchart->ns_showcatnameval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'showSerName' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_barchart->ns_showsernameval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'showPercent' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_barchart->ns_showpercentval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'showBubbleSize' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_barchart->ns_showbubblesizeval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'gapWidth' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_barchart->ns_gapwidthval = ls_prop-val.

        "Load axes
        node2 ?= node->find_from_name_ns( name = 'barChart' uri = namespace-c ).
        CALL METHOD node2->get_elements_by_tag_name_ns
          EXPORTING
*           depth     = 0
            name = 'axId'
            uri  = namespace-c
          RECEIVING
            rval = lo_collection.
        CALL METHOD lo_collection->create_iterator
          RECEIVING
            rval = lo_iterator.
        lo_node = lo_iterator->get_next( ).
        IF lo_node IS BOUND.
          node2 ?= lo_node->query_interface( c_ixml_iid_element ).
        ENDIF.
        WHILE lo_node IS BOUND.
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
          lv_axid = ls_prop-val.
          IF sy-index EQ 1. "catAx
            node2 ?= node->find_from_name_ns( name = 'catAx' uri = namespace-c ).
            node3 ?= node2->find_from_name_ns( name = 'orientation' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_orientation = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'delete' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_delete = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'axPos' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_axpos = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'numFmt' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_formatcode = ls_prop-formatcode.
            lv_sourcelinked = ls_prop-sourcelinked.
            node3 ?= node2->find_from_name_ns( name = 'majorTickMark' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_majortickmark = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'majorTickMark' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_minortickmark = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'tickLblPos' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_ticklblpos = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'crossAx' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_crossax = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'crosses' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_crosses = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'auto' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_auto = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'lblAlgn' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_lblalgn = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'lblOffset' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_lbloffset = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'noMultiLvlLbl' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_nomultilvllbl = ls_prop-val.
            CALL METHOD lo_barchart->create_ax
              EXPORTING
                ip_axid          = lv_axid
                ip_type          = zcl_excel_graph_bars=>c_catax
                ip_orientation   = lv_orientation
                ip_delete        = lv_delete
                ip_axpos         = lv_axpos
                ip_formatcode    = lv_formatcode
                ip_sourcelinked  = lv_sourcelinked
                ip_majortickmark = lv_majortickmark
                ip_minortickmark = lv_minortickmark
                ip_ticklblpos    = lv_ticklblpos
                ip_crossax       = lv_crossax
                ip_crosses       = lv_crosses
                ip_auto          = lv_auto
                ip_lblalgn       = lv_lblalgn
                ip_lbloffset     = lv_lbloffset
                ip_nomultilvllbl = lv_nomultilvllbl.
          ELSEIF sy-index EQ 2. "valAx
            node2 ?= node->find_from_name_ns( name = 'valAx' uri = namespace-c ).
            node3 ?= node2->find_from_name_ns( name = 'orientation' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_orientation = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'delete' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_delete = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'axPos' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_axpos = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'numFmt' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_formatcode = ls_prop-formatcode.
            lv_sourcelinked = ls_prop-sourcelinked.
            node3 ?= node2->find_from_name_ns( name = 'majorTickMark' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_majortickmark = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'majorTickMark' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_minortickmark = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'tickLblPos' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_ticklblpos = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'crossAx' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_crossax = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'crosses' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_crosses = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'crossBetween' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_crossbetween = ls_prop-val.
            CALL METHOD lo_barchart->create_ax
              EXPORTING
                ip_axid          = lv_axid
                ip_type          = zcl_excel_graph_bars=>c_valax
                ip_orientation   = lv_orientation
                ip_delete        = lv_delete
                ip_axpos         = lv_axpos
                ip_formatcode    = lv_formatcode
                ip_sourcelinked  = lv_sourcelinked
                ip_majortickmark = lv_majortickmark
                ip_minortickmark = lv_minortickmark
                ip_ticklblpos    = lv_ticklblpos
                ip_crossax       = lv_crossax
                ip_crosses       = lv_crosses
                ip_crossbetween  = lv_crossbetween.
          ENDIF.
          lo_node = lo_iterator->get_next( ).
          IF lo_node IS BOUND.
            node2 ?= lo_node->query_interface( c_ixml_iid_element ).
          ENDIF.
        ENDWHILE.

      WHEN c_graph_pie.
        node2 ?= node->find_from_name_ns( name = 'varyColors' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_piechart->ns_varycolorsval = ls_prop-val.

        "Load series
        CALL METHOD node->get_elements_by_tag_name_ns
          EXPORTING
*           depth     = 0
            name = 'ser'
            uri  = namespace-c
          RECEIVING
            rval = lo_collection.
        CALL METHOD lo_collection->create_iterator
          RECEIVING
            rval = lo_iterator.
        lo_node = lo_iterator->get_next( ).
        IF lo_node IS BOUND.
          node2 ?= lo_node->query_interface( c_ixml_iid_element ).
        ENDIF.
        WHILE lo_node IS BOUND.
          node3 ?= node2->find_from_name_ns( name = 'idx' uri = namespace-c ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_idx = ls_prop-val.
          node3 ?= node2->find_from_name_ns( name = 'order' uri = namespace-c ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_order = ls_prop-val.
          node3 ?= node2->find_from_name_ns( name = 'v' uri = namespace-c ).
          IF node3 IS BOUND.
            lv_sername = node3->get_value( ).
          ENDIF.
          node3 ?= node2->find_from_name_ns( name = 'strRef' uri = namespace-c ).
          IF node3 IS BOUND.
            node4 ?= node3->find_from_name_ns( name = 'f' uri = namespace-c ).
            lv_label = node4->get_value( ).
          ENDIF.
          node3 ?= node2->find_from_name_ns( name = 'numRef' uri = namespace-c ).
          IF node3 IS BOUND.
            node4 ?= node3->find_from_name_ns( name = 'f' uri = namespace-c ).
            lv_value = node4->get_value( ).
          ENDIF.
          CALL METHOD lo_piechart->create_serie
            EXPORTING
              ip_idx     = lv_idx
              ip_order   = lv_order
              ip_lbl     = lv_label
              ip_ref     = lv_value
              ip_sername = lv_sername.
          lo_node = lo_iterator->get_next( ).
          IF lo_node IS BOUND.
            node2 ?= lo_node->query_interface( c_ixml_iid_element ).
          ENDIF.
        ENDWHILE.

        "note: numCache avoided
        node2 ?= node->find_from_name_ns( name = 'showLegendKey' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_piechart->ns_showlegendkeyval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'showVal' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_piechart->ns_showvalval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'showCatName' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_piechart->ns_showcatnameval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'showSerName' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_piechart->ns_showsernameval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'showPercent' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_piechart->ns_showpercentval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'showBubbleSize' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_piechart->ns_showbubblesizeval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'showLeaderLines' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_piechart->ns_showleaderlinesval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'firstSliceAng' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_piechart->ns_firstsliceangval = ls_prop-val.
      WHEN c_graph_line.
        node2 ?= node->find_from_name_ns( name = 'grouping' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_linechart->ns_groupingval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'varyColors' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_linechart->ns_varycolorsval = ls_prop-val.

        "Load series
        CALL METHOD node->get_elements_by_tag_name_ns
          EXPORTING
*           depth     = 0
            name = 'ser'
            uri  = namespace-c
          RECEIVING
            rval = lo_collection.
        CALL METHOD lo_collection->create_iterator
          RECEIVING
            rval = lo_iterator.
        lo_node = lo_iterator->get_next( ).
        IF lo_node IS BOUND.
          node2 ?= lo_node->query_interface( c_ixml_iid_element ).
        ENDIF.
        WHILE lo_node IS BOUND.
          node3 ?= node2->find_from_name_ns( name = 'idx' uri = namespace-c ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_idx = ls_prop-val.
          node3 ?= node2->find_from_name_ns( name = 'order' uri = namespace-c ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_order = ls_prop-val.
          node3 ?= node2->find_from_name_ns( name = 'symbol' uri = namespace-c ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_symbol = ls_prop-val.
          node3 ?= node2->find_from_name_ns( name = 'smooth' uri = namespace-c ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_smooth = ls_prop-val.
          node3 ?= node2->find_from_name_ns( name = 'v' uri = namespace-c ).
          IF node3 IS BOUND.
            lv_sername = node3->get_value( ).
          ENDIF.
          node3 ?= node2->find_from_name_ns( name = 'strRef' uri = namespace-c ).
          IF node3 IS BOUND.
            node4 ?= node3->find_from_name_ns( name = 'f' uri = namespace-c ).
            lv_label = node4->get_value( ).
          ENDIF.
          node3 ?= node2->find_from_name_ns( name = 'numRef' uri = namespace-c ).
          IF node3 IS BOUND.
            node4 ?= node3->find_from_name_ns( name = 'f' uri = namespace-c ).
            lv_value = node4->get_value( ).
          ENDIF.
          CALL METHOD lo_linechart->create_serie
            EXPORTING
              ip_idx     = lv_idx
              ip_order   = lv_order
              ip_symbol  = lv_symbol
              ip_smooth  = lv_smooth
              ip_lbl     = lv_label
              ip_ref     = lv_value
              ip_sername = lv_sername.
          lo_node = lo_iterator->get_next( ).
          IF lo_node IS BOUND.
            node2 ?= lo_node->query_interface( c_ixml_iid_element ).
          ENDIF.
        ENDWHILE.
        "note: numCache avoided
        node2 ?= node->find_from_name_ns( name = 'showLegendKey' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_linechart->ns_showlegendkeyval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'showVal' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_linechart->ns_showvalval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'showCatName' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_linechart->ns_showcatnameval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'showSerName' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_linechart->ns_showsernameval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'showPercent' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_linechart->ns_showpercentval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'showBubbleSize' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_linechart->ns_showbubblesizeval = ls_prop-val.

        node ?= node->find_from_name_ns( name = 'lineChart' uri = namespace-c ).
        node2 ?= node->find_from_name_ns( name = 'marker' uri = namespace-c depth = '1' ).
        IF node2 IS BOUND.
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
          lo_linechart->ns_markerval = ls_prop-val.
        ENDIF.
        node2 ?= node->find_from_name_ns( name = 'smooth' uri = namespace-c depth = '1' ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_linechart->ns_smoothval = ls_prop-val.
        node ?= ip_chart->if_ixml_node~get_first_child( ).
        CHECK node IS NOT INITIAL.

        "Load axes
        node2 ?= node->find_from_name_ns( name = 'lineChart' uri = namespace-c ).
        CALL METHOD node2->get_elements_by_tag_name_ns
          EXPORTING
*           depth     = 0
            name = 'axId'
            uri  = namespace-c
          RECEIVING
            rval = lo_collection.
        CALL METHOD lo_collection->create_iterator
          RECEIVING
            rval = lo_iterator.
        lo_node = lo_iterator->get_next( ).
        IF lo_node IS BOUND.
          node2 ?= lo_node->query_interface( c_ixml_iid_element ).
        ENDIF.
        WHILE lo_node IS BOUND.
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
          lv_axid = ls_prop-val.
          IF sy-index EQ 1. "catAx
            node2 ?= node->find_from_name_ns( name = 'catAx' uri = namespace-c ).
            node3 ?= node2->find_from_name_ns( name = 'orientation' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_orientation = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'delete' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_delete = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'axPos' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_axpos = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'majorTickMark' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_majortickmark = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'majorTickMark' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_minortickmark = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'tickLblPos' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_ticklblpos = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'crossAx' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_crossax = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'crosses' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_crosses = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'auto' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_auto = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'lblAlgn' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_lblalgn = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'lblOffset' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_lbloffset = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'noMultiLvlLbl' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_nomultilvllbl = ls_prop-val.
            CALL METHOD lo_linechart->create_ax
              EXPORTING
                ip_axid          = lv_axid
                ip_type          = zcl_excel_graph_line=>c_catax
                ip_orientation   = lv_orientation
                ip_delete        = lv_delete
                ip_axpos         = lv_axpos
                ip_formatcode    = lv_formatcode
                ip_sourcelinked  = lv_sourcelinked
                ip_majortickmark = lv_majortickmark
                ip_minortickmark = lv_minortickmark
                ip_ticklblpos    = lv_ticklblpos
                ip_crossax       = lv_crossax
                ip_crosses       = lv_crosses
                ip_auto          = lv_auto
                ip_lblalgn       = lv_lblalgn
                ip_lbloffset     = lv_lbloffset
                ip_nomultilvllbl = lv_nomultilvllbl.
          ELSEIF sy-index EQ 2. "valAx
            node2 ?= node->find_from_name_ns( name = 'valAx' uri = namespace-c ).
            node3 ?= node2->find_from_name_ns( name = 'orientation' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_orientation = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'delete' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_delete = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'axPos' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_axpos = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'numFmt' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_formatcode = ls_prop-formatcode.
            lv_sourcelinked = ls_prop-sourcelinked.
            node3 ?= node2->find_from_name_ns( name = 'majorTickMark' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_majortickmark = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'majorTickMark' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_minortickmark = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'tickLblPos' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_ticklblpos = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'crossAx' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_crossax = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'crosses' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_crosses = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'crossBetween' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_crossbetween = ls_prop-val.
            CALL METHOD lo_linechart->create_ax
              EXPORTING
                ip_axid          = lv_axid
                ip_type          = zcl_excel_graph_line=>c_valax
                ip_orientation   = lv_orientation
                ip_delete        = lv_delete
                ip_axpos         = lv_axpos
                ip_formatcode    = lv_formatcode
                ip_sourcelinked  = lv_sourcelinked
                ip_majortickmark = lv_majortickmark
                ip_minortickmark = lv_minortickmark
                ip_ticklblpos    = lv_ticklblpos
                ip_crossax       = lv_crossax
                ip_crosses       = lv_crosses
                ip_crossbetween  = lv_crossbetween.
          ENDIF.
          lo_node = lo_iterator->get_next( ).
          IF lo_node IS BOUND.
            node2 ?= lo_node->query_interface( c_ixml_iid_element ).
          ENDIF.
        ENDWHILE.
      WHEN OTHERS.
    ENDCASE.

    "legend
    CASE me->graph_type.
      WHEN c_graph_bars.
        node2 ?= node->find_from_name_ns( name = 'legendPos' uri = namespace-c ).
        IF node2 IS BOUND.
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
          lo_barchart->ns_legendposval = ls_prop-val.
        ENDIF.
        node2 ?= node->find_from_name_ns( name = 'overlay' uri = namespace-c ).
        IF node2 IS BOUND.
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
          lo_barchart->ns_overlayval = ls_prop-val.
        ENDIF.
      WHEN c_graph_line.
        node2 ?= node->find_from_name_ns( name = 'legendPos' uri = namespace-c ).
        IF node2 IS BOUND.
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
          lo_linechart->ns_legendposval = ls_prop-val.
        ENDIF.
        node2 ?= node->find_from_name_ns( name = 'overlay' uri = namespace-c ).
        IF node2 IS BOUND.
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
          lo_linechart->ns_overlayval = ls_prop-val.
        ENDIF.
      WHEN c_graph_pie.
        node2 ?= node->find_from_name_ns( name = 'legendPos' uri = namespace-c ).
        IF node2 IS BOUND.
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
          lo_piechart->ns_legendposval = ls_prop-val.
        ENDIF.
        node2 ?= node->find_from_name_ns( name = 'overlay' uri = namespace-c ).
        IF node2 IS BOUND.
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
          lo_piechart->ns_overlayval = ls_prop-val.
        ENDIF.
        node2 ?= node->find_from_name_ns( name = 'pPr' uri = namespace-a ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_piechart->ns_pprrtl = ls_prop-rtl.
        node2 ?= node->find_from_name_ns( name = 'endParaRPr' uri = namespace-a ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_piechart->ns_endpararprlang = ls_prop-lang.

      WHEN OTHERS.
    ENDCASE.

    node2 ?= node->find_from_name_ns( name = 'plotVisOnly' uri = namespace-c ).
    zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
    me->graph->ns_plotvisonlyval = ls_prop-val.
    node2 ?= node->find_from_name_ns( name = 'dispBlanksAs' uri = namespace-c ).
    zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
    me->graph->ns_dispblanksasval = ls_prop-val.
    node2 ?= node->find_from_name_ns( name = 'showDLblsOverMax' uri = namespace-c ).
    zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
    me->graph->ns_showdlblsovermaxval = ls_prop-val.
    "---------------------

    node2 ?= node->find_from_name_ns( name = 'pageMargins' uri = namespace-c ).
    zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_pagemargins ).
    me->graph->pagemargins = ls_pagemargins.


  ENDMETHOD.