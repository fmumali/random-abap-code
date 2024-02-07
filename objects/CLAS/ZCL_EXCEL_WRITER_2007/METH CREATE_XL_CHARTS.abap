  METHOD create_xl_charts.


** Constant node name
    CONSTANTS: lc_xml_node_chartspace         TYPE string VALUE 'c:chartSpace',
               lc_xml_node_ns_c               TYPE string VALUE 'http://schemas.openxmlformats.org/drawingml/2006/chart',
               lc_xml_node_ns_a               TYPE string VALUE 'http://schemas.openxmlformats.org/drawingml/2006/main',
               lc_xml_node_ns_r               TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
               lc_xml_node_date1904           TYPE string VALUE 'c:date1904',
               lc_xml_node_lang               TYPE string VALUE 'c:lang',
               lc_xml_node_roundedcorners     TYPE string VALUE 'c:roundedCorners',
               lc_xml_node_altcont            TYPE string VALUE 'mc:AlternateContent',
               lc_xml_node_altcont_ns_mc      TYPE string VALUE 'http://schemas.openxmlformats.org/markup-compatibility/2006',
               lc_xml_node_choice             TYPE string VALUE 'mc:Choice',
               lc_xml_node_choice_ns_requires TYPE string VALUE 'c14',
               lc_xml_node_choice_ns_c14      TYPE string VALUE 'http://schemas.microsoft.com/office/drawing/2007/8/2/chart',
               lc_xml_node_style              TYPE string VALUE 'c14:style',
               lc_xml_node_fallback           TYPE string VALUE 'mc:Fallback',
               lc_xml_node_style2             TYPE string VALUE 'c:style',

               "---------------------------CHART
               lc_xml_node_chart              TYPE string VALUE 'c:chart',
               lc_xml_node_autotitledeleted   TYPE string VALUE 'c:autoTitleDeleted',
               "plotArea
               lc_xml_node_plotarea           TYPE string VALUE 'c:plotArea',
               lc_xml_node_layout             TYPE string VALUE 'c:layout',
               lc_xml_node_varycolors         TYPE string VALUE 'c:varyColors',
               lc_xml_node_ser                TYPE string VALUE 'c:ser',
               lc_xml_node_idx                TYPE string VALUE 'c:idx',
               lc_xml_node_order              TYPE string VALUE 'c:order',
               lc_xml_node_tx                 TYPE string VALUE 'c:tx',
               lc_xml_node_v                  TYPE string VALUE 'c:v',
               lc_xml_node_val                TYPE string VALUE 'c:val',
               lc_xml_node_cat                TYPE string VALUE 'c:cat',
               lc_xml_node_numref             TYPE string VALUE 'c:numRef',
               lc_xml_node_strref             TYPE string VALUE 'c:strRef',
               lc_xml_node_f                  TYPE string VALUE 'c:f', "this is the range
               lc_xml_node_overlap            TYPE string VALUE 'c:overlap',
               "note: numcache avoided
               lc_xml_node_dlbls              TYPE string VALUE 'c:dLbls',
               lc_xml_node_showlegendkey      TYPE string VALUE 'c:showLegendKey',
               lc_xml_node_showval            TYPE string VALUE 'c:showVal',
               lc_xml_node_showcatname        TYPE string VALUE 'c:showCatName',
               lc_xml_node_showsername        TYPE string VALUE 'c:showSerName',
               lc_xml_node_showpercent        TYPE string VALUE 'c:showPercent',
               lc_xml_node_showbubblesize     TYPE string VALUE 'c:showBubbleSize',
               "plotArea->pie
               lc_xml_node_piechart           TYPE string VALUE 'c:pieChart',
               lc_xml_node_showleaderlines    TYPE string VALUE 'c:showLeaderLines',
               lc_xml_node_firstsliceang      TYPE string VALUE 'c:firstSliceAng',
               "plotArea->line
               lc_xml_node_linechart          TYPE string VALUE 'c:lineChart',
               lc_xml_node_symbol             TYPE string VALUE 'c:symbol',
               lc_xml_node_marker             TYPE string VALUE 'c:marker',
               lc_xml_node_smooth             TYPE string VALUE 'c:smooth',
               "plotArea->bar
               lc_xml_node_invertifnegative   TYPE string VALUE 'c:invertIfNegative',
               lc_xml_node_barchart           TYPE string VALUE 'c:barChart',
               lc_xml_node_bardir             TYPE string VALUE 'c:barDir',
               lc_xml_node_gapwidth           TYPE string VALUE 'c:gapWidth',
               "plotArea->line + plotArea->bar
               lc_xml_node_grouping           TYPE string VALUE 'c:grouping',
               lc_xml_node_axid               TYPE string VALUE 'c:axId',
               lc_xml_node_catax              TYPE string VALUE 'c:catAx',
               lc_xml_node_valax              TYPE string VALUE 'c:valAx',
               lc_xml_node_scaling            TYPE string VALUE 'c:scaling',
               lc_xml_node_orientation        TYPE string VALUE 'c:orientation',
               lc_xml_node_delete             TYPE string VALUE 'c:delete',
               lc_xml_node_axpos              TYPE string VALUE 'c:axPos',
               lc_xml_node_numfmt             TYPE string VALUE 'c:numFmt',
               lc_xml_node_majorgridlines     TYPE string VALUE 'c:majorGridlines',
               lc_xml_node_majortickmark      TYPE string VALUE 'c:majorTickMark',
               lc_xml_node_minortickmark      TYPE string VALUE 'c:minorTickMark',
               lc_xml_node_ticklblpos         TYPE string VALUE 'c:tickLblPos',
               lc_xml_node_crossax            TYPE string VALUE 'c:crossAx',
               lc_xml_node_crosses            TYPE string VALUE 'c:crosses',
               lc_xml_node_auto               TYPE string VALUE 'c:auto',
               lc_xml_node_lblalgn            TYPE string VALUE 'c:lblAlgn',
               lc_xml_node_lbloffset          TYPE string VALUE 'c:lblOffset',
               lc_xml_node_nomultilvllbl      TYPE string VALUE 'c:noMultiLvlLbl',
               lc_xml_node_crossbetween       TYPE string VALUE 'c:crossBetween',
               "legend
               lc_xml_node_legend             TYPE string VALUE 'c:legend',
               "legend->pie
               lc_xml_node_legendpos          TYPE string VALUE 'c:legendPos',
*                  lc_xml_node_layout            TYPE string VALUE 'c:layout', "already exist
               lc_xml_node_overlay            TYPE string VALUE 'c:overlay',
               lc_xml_node_txpr               TYPE string VALUE 'c:txPr',
               lc_xml_node_bodypr             TYPE string VALUE 'a:bodyPr',
               lc_xml_node_lststyle           TYPE string VALUE 'a:lstStyle',
               lc_xml_node_p                  TYPE string VALUE 'a:p',
               lc_xml_node_ppr                TYPE string VALUE 'a:pPr',
               lc_xml_node_defrpr             TYPE string VALUE 'a:defRPr',
               lc_xml_node_endpararpr         TYPE string VALUE 'a:endParaRPr',
               "legend->bar + legend->line
               lc_xml_node_plotvisonly        TYPE string VALUE 'c:plotVisOnly',
               lc_xml_node_dispblanksas       TYPE string VALUE 'c:dispBlanksAs',
               lc_xml_node_showdlblsovermax   TYPE string VALUE 'c:showDLblsOverMax',
               "---------------------------END OF CHART

               lc_xml_node_printsettings      TYPE string VALUE 'c:printSettings',
               lc_xml_node_headerfooter       TYPE string VALUE 'c:headerFooter',
               lc_xml_node_pagemargins        TYPE string VALUE 'c:pageMargins',
               lc_xml_node_pagesetup          TYPE string VALUE 'c:pageSetup'.


    DATA: lo_document     TYPE REF TO if_ixml_document,
          lo_element_root TYPE REF TO if_ixml_element.


    DATA lo_element                               TYPE REF TO if_ixml_element.
    DATA lo_element2                              TYPE REF TO if_ixml_element.
    DATA lo_element3                              TYPE REF TO if_ixml_element.
    DATA lo_el_rootchart                           TYPE REF TO if_ixml_element.
    DATA lo_element4                              TYPE REF TO if_ixml_element.
    DATA lo_element5                              TYPE REF TO if_ixml_element.
    DATA lo_element6                              TYPE REF TO if_ixml_element.
    DATA lo_element7                              TYPE REF TO if_ixml_element.

**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

***********************************************************************
* STEP 3: Create main node relationships
    lo_element_root  = lo_document->create_simple_element( name   = lc_xml_node_chartspace
                                                           parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:c'
                                       value = lc_xml_node_ns_c ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:a'
                                       value = lc_xml_node_ns_a ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:r'
                                       value = lc_xml_node_ns_r ).

**********************************************************************
* STEP 4: Create chart

    DATA lo_chartb TYPE REF TO zcl_excel_graph_bars.
    DATA lo_chartp TYPE REF TO zcl_excel_graph_pie.
    DATA lo_chartl TYPE REF TO zcl_excel_graph_line.
    DATA lo_chart TYPE REF TO zcl_excel_graph.

    DATA ls_serie TYPE zcl_excel_graph=>s_series.
    DATA ls_ax TYPE zcl_excel_graph_bars=>s_ax.
    DATA lv_str TYPE string.

    "Identify chart type
    CASE io_drawing->graph_type.
      WHEN zcl_excel_drawing=>c_graph_bars.
        lo_chartb ?= io_drawing->graph.
      WHEN zcl_excel_drawing=>c_graph_pie.
        lo_chartp ?= io_drawing->graph.
      WHEN zcl_excel_drawing=>c_graph_line.
        lo_chartl ?= io_drawing->graph.
      WHEN OTHERS.
    ENDCASE.


    lo_chart = io_drawing->graph.

    lo_element = lo_document->create_simple_element( name = lc_xml_node_date1904
                                                         parent = lo_element_root ).
    lo_element->set_attribute_ns( name  = 'val'
                                      value = lo_chart->ns_1904val ).

    lo_element = lo_document->create_simple_element( name = lc_xml_node_lang
                                                         parent = lo_element_root ).
    lo_element->set_attribute_ns( name  = 'val'
                                      value = lo_chart->ns_langval ).

    lo_element = lo_document->create_simple_element( name = lc_xml_node_roundedcorners
                                                         parent = lo_element_root ).
    lo_element->set_attribute_ns( name  = 'val'
                                      value = lo_chart->ns_roundedcornersval ).

    lo_element = lo_document->create_simple_element( name = lc_xml_node_altcont
                                                         parent = lo_element_root ).
    lo_element->set_attribute_ns( name  = 'xmlns:mc'
                                      value = lc_xml_node_altcont_ns_mc ).

    "Choice
    lo_element2 = lo_document->create_simple_element( name = lc_xml_node_choice
                                                         parent = lo_element ).
    lo_element2->set_attribute_ns( name  = 'Requires'
                                      value = lc_xml_node_choice_ns_requires ).
    lo_element2->set_attribute_ns( name  = 'xmlns:c14'
                                      value = lc_xml_node_choice_ns_c14 ).

    "C14:style
    lo_element3 = lo_document->create_simple_element( name = lc_xml_node_style
                                                         parent = lo_element2 ).
    lo_element3->set_attribute_ns( name  = 'val'
                                      value = lo_chart->ns_c14styleval ).

    "Fallback
    lo_element2 = lo_document->create_simple_element( name = lc_xml_node_fallback
                                                         parent = lo_element ).

    "C:style
    lo_element3 = lo_document->create_simple_element( name = lc_xml_node_style2
                                                         parent = lo_element2 ).
    lo_element3->set_attribute_ns( name  = 'val'
                                      value = lo_chart->ns_styleval ).

    "---------------------------CHART
    lo_element = lo_document->create_simple_element( name = lc_xml_node_chart
                                                         parent = lo_element_root ).
    "Added
    IF lo_chart->title IS NOT INITIAL.
      lo_element2 = lo_document->create_simple_element( name = 'c:title'
                                                           parent = lo_element ).
      lo_element3 = lo_document->create_simple_element( name = 'c:tx'
                                                           parent = lo_element2 ).
      lo_element4 = lo_document->create_simple_element( name = 'c:rich'
                                                           parent = lo_element3 ).
      lo_element5 = lo_document->create_simple_element( name = 'a:bodyPr'
                                                           parent = lo_element4 ).
      lo_element5 = lo_document->create_simple_element( name = 'a:lstStyle'
                                                           parent = lo_element4 ).
      lo_element5 = lo_document->create_simple_element( name = 'a:p'
                                                           parent = lo_element4 ).
      lo_element6 = lo_document->create_simple_element( name = 'a:pPr'
                                                           parent = lo_element5 ).
      lo_element7 = lo_document->create_simple_element( name = 'a:defRPr'
                                                           parent = lo_element6 ).
      lo_element6 = lo_document->create_simple_element( name = 'a:r'
                                                           parent = lo_element5 ).
      lo_element7 = lo_document->create_simple_element( name = 'a:rPr'
                                                           parent = lo_element6 ).
      lo_element7->set_attribute_ns( name  = 'lang'
                                        value = 'en-US' ).
      lo_element7 = lo_document->create_simple_element( name = 'a:t'
                                                           parent = lo_element6 ).
      lo_element7->set_value( value = lo_chart->title ).
    ENDIF.
    "End
    lo_element2 = lo_document->create_simple_element( name = lc_xml_node_autotitledeleted
                                                         parent = lo_element ).
    lo_element2->set_attribute_ns( name  = 'val'
                                      value = lo_chart->ns_autotitledeletedval ).

    "plotArea
    lo_element2 = lo_document->create_simple_element( name = lc_xml_node_plotarea
                                                       parent = lo_element ).
    lo_element3 = lo_document->create_simple_element( name = lc_xml_node_layout
                                                       parent = lo_element2 ).
    CASE io_drawing->graph_type.
      WHEN zcl_excel_drawing=>c_graph_bars.
        "----bar
        lo_element3 = lo_document->create_simple_element( name = lc_xml_node_barchart
                                                     parent = lo_element2 ).
        lo_element4 = lo_document->create_simple_element( name = lc_xml_node_bardir
                                                     parent = lo_element3 ).
        lo_element4->set_attribute_ns( name  = 'val'
                                  value = lo_chartb->ns_bardirval ).
        lo_element4 = lo_document->create_simple_element( name = lc_xml_node_grouping
                                                     parent = lo_element3 ).
        lo_element4->set_attribute_ns( name  = 'val'
                                  value = lo_chartb->ns_groupingval ).
        lo_element4 = lo_document->create_simple_element( name = lc_xml_node_varycolors
                                                     parent = lo_element3 ).
        lo_element4->set_attribute_ns( name  = 'val'
                                  value = lo_chartb->ns_varycolorsval ).

        "series
        LOOP AT lo_chartb->series INTO ls_serie.
          lo_element4 = lo_document->create_simple_element( name = lc_xml_node_ser
                                                     parent = lo_element3 ).
          lo_element5 = lo_document->create_simple_element( name = lc_xml_node_idx
                                                     parent = lo_element4 ).
          IF ls_serie-idx IS NOT INITIAL.
            lv_str = ls_serie-idx.
          ELSE.
            lv_str = sy-tabix - 1.
          ENDIF.
          CONDENSE lv_str.
          lo_element5->set_attribute_ns( name  = 'val'
                                  value = lv_str ).
          lo_element5 = lo_document->create_simple_element( name = lc_xml_node_order
                                                     parent = lo_element4 ).
          lv_str = ls_serie-order.
          CONDENSE lv_str.
          lo_element5->set_attribute_ns( name  = 'val'
                                  value = lv_str ).
          IF ls_serie-sername IS NOT INITIAL.
            lo_element5 = lo_document->create_simple_element( name = lc_xml_node_tx
                                                      parent = lo_element4 ).
            lo_element6 = lo_document->create_simple_element( name = lc_xml_node_v
                                                      parent = lo_element5 ).
            lo_element6->set_value( value = ls_serie-sername ).
          ENDIF.
          lo_element5 = lo_document->create_simple_element( name = lc_xml_node_invertifnegative
                                                     parent = lo_element4 ).
          lo_element5->set_attribute_ns( name  = 'val'
                                  value = ls_serie-invertifnegative ).
          IF ls_serie-lbl IS NOT INITIAL.
            lo_element5 = lo_document->create_simple_element( name = lc_xml_node_cat
                                                       parent = lo_element4 ).
            lo_element6 = lo_document->create_simple_element( name = lc_xml_node_strref
                                                       parent = lo_element5 ).
            lo_element7 = lo_document->create_simple_element( name = lc_xml_node_f
                                                       parent = lo_element6 ).
            lo_element7->set_value( value = ls_serie-lbl ).
          ENDIF.
          IF ls_serie-ref IS NOT INITIAL.
            lo_element5 = lo_document->create_simple_element( name = lc_xml_node_val
                                                       parent = lo_element4 ).
            lo_element6 = lo_document->create_simple_element( name = lc_xml_node_numref
                                                       parent = lo_element5 ).
            lo_element7 = lo_document->create_simple_element( name = lc_xml_node_f
                                                       parent = lo_element6 ).
            lo_element7->set_value( value = ls_serie-ref ).
          ENDIF.
        ENDLOOP.
        "endseries
        IF lo_chartb->ns_groupingval = zcl_excel_graph_bars=>c_groupingval_stacked.
          lo_element4 = lo_document->create_simple_element( name = lc_xml_node_overlap
                                                            parent = lo_element3 ).
          lo_element4->set_attribute_ns( name  = 'val'
                                         value = '100' ).
        ENDIF.

        lo_element4 = lo_document->create_simple_element( name = lc_xml_node_dlbls
                                                     parent = lo_element3 ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showlegendkey
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartb->ns_showlegendkeyval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showval
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartb->ns_showvalval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showcatname
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartb->ns_showcatnameval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showsername
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartb->ns_showsernameval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showpercent
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartb->ns_showpercentval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showbubblesize
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartb->ns_showbubblesizeval ).

        lo_element4 = lo_document->create_simple_element( name = lc_xml_node_gapwidth
                                                     parent = lo_element3 ).
        lo_element4->set_attribute_ns( name  = 'val'
                                  value = lo_chartb->ns_gapwidthval ).

        "axes
        lo_el_rootchart = lo_element3.
        LOOP AT lo_chartb->axes INTO ls_ax.
          lo_element4 = lo_document->create_simple_element( name = lc_xml_node_axid
                                                     parent = lo_el_rootchart ).
          lo_element4->set_attribute_ns( name  = 'val'
                                  value = ls_ax-axid ).
          CASE ls_ax-type.
            WHEN zcl_excel_graph_bars=>c_catax.
              lo_element3 = lo_document->create_simple_element( name = lc_xml_node_catax
                                                     parent = lo_element2 ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_axid
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-axid ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_scaling
                                                     parent = lo_element3 ).
              lo_element5 = lo_document->create_simple_element( name = lc_xml_node_orientation
                                                     parent = lo_element4 ).
              lo_element5->set_attribute_ns( name  = 'val'
                                             value = ls_ax-orientation ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_delete
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-delete ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_axpos
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-axpos ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_numfmt
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'formatCode'
                                             value = ls_ax-formatcode ).
              lo_element4->set_attribute_ns( name  = 'sourceLinked'
                                             value = ls_ax-sourcelinked ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_majortickmark
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-majortickmark ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_minortickmark
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-minortickmark ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_ticklblpos
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-ticklblpos ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_crossax
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-crossax ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_crosses
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-crosses ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_auto
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-auto ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_lblalgn
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-lblalgn ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_lbloffset
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-lbloffset ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_nomultilvllbl
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-nomultilvllbl ).
            WHEN zcl_excel_graph_bars=>c_valax.
              lo_element3 = lo_document->create_simple_element( name = lc_xml_node_valax
                                                     parent = lo_element2 ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_axid
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-axid ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_scaling
                                                     parent = lo_element3 ).
              lo_element5 = lo_document->create_simple_element( name = lc_xml_node_orientation
                                                     parent = lo_element4 ).
              lo_element5->set_attribute_ns( name  = 'val'
                                             value = ls_ax-orientation ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_delete
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-delete ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_axpos
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-axpos ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_majorgridlines
                                                     parent = lo_element3 ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_numfmt
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'formatCode'
                                             value = ls_ax-formatcode ).
              lo_element4->set_attribute_ns( name  = 'sourceLinked'
                                             value = ls_ax-sourcelinked ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_majortickmark
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-majortickmark ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_minortickmark
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-minortickmark ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_ticklblpos
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-ticklblpos ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_crossax
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-crossax ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_crosses
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-crosses ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_crossbetween
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-crossbetween ).
            WHEN OTHERS.
          ENDCASE.
        ENDLOOP.
        "endaxes

      WHEN zcl_excel_drawing=>c_graph_pie.
        "----pie
        lo_element3 = lo_document->create_simple_element( name = lc_xml_node_piechart
                                                     parent = lo_element2 ).
        lo_element4 = lo_document->create_simple_element( name = lc_xml_node_varycolors
                                                     parent = lo_element3 ).
        lo_element4->set_attribute_ns( name  = 'val'
                                  value = lo_chartp->ns_varycolorsval ).

        "series
        LOOP AT lo_chartp->series INTO ls_serie.
          lo_element4 = lo_document->create_simple_element( name = lc_xml_node_ser
                                                     parent = lo_element3 ).
          lo_element5 = lo_document->create_simple_element( name = lc_xml_node_idx
                                                     parent = lo_element4 ).
          IF ls_serie-idx IS NOT INITIAL.
            lv_str = ls_serie-idx.
          ELSE.
            lv_str = sy-tabix - 1.
          ENDIF.
          CONDENSE lv_str.
          lo_element5->set_attribute_ns( name  = 'val'
                                  value = lv_str ).
          lo_element5 = lo_document->create_simple_element( name = lc_xml_node_order
                                                     parent = lo_element4 ).
          lv_str = ls_serie-order.
          CONDENSE lv_str.
          lo_element5->set_attribute_ns( name  = 'val'
                                  value = lv_str ).
          IF ls_serie-sername IS NOT INITIAL.
            lo_element5 = lo_document->create_simple_element( name = lc_xml_node_tx
                                                      parent = lo_element4 ).
            lo_element6 = lo_document->create_simple_element( name = lc_xml_node_v
                                                      parent = lo_element5 ).
            lo_element6->set_value( value = ls_serie-sername ).
          ENDIF.
          IF ls_serie-lbl IS NOT INITIAL.
            lo_element5 = lo_document->create_simple_element( name = lc_xml_node_cat
                                                       parent = lo_element4 ).
            lo_element6 = lo_document->create_simple_element( name = lc_xml_node_strref
                                                       parent = lo_element5 ).
            lo_element7 = lo_document->create_simple_element( name = lc_xml_node_f
                                                       parent = lo_element6 ).
            lo_element7->set_value( value = ls_serie-lbl ).
          ENDIF.
          IF ls_serie-ref IS NOT INITIAL.
            lo_element5 = lo_document->create_simple_element( name = lc_xml_node_val
                                                       parent = lo_element4 ).
            lo_element6 = lo_document->create_simple_element( name = lc_xml_node_numref
                                                       parent = lo_element5 ).
            lo_element7 = lo_document->create_simple_element( name = lc_xml_node_f
                                                       parent = lo_element6 ).
            lo_element7->set_value( value = ls_serie-ref ).
          ENDIF.
        ENDLOOP.
        "endseries

        lo_element4 = lo_document->create_simple_element( name = lc_xml_node_dlbls
                                                     parent = lo_element3 ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showlegendkey
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartp->ns_showlegendkeyval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showval
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartp->ns_showvalval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showcatname
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartp->ns_showcatnameval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showsername
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartp->ns_showsernameval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showpercent
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartp->ns_showpercentval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showbubblesize
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartp->ns_showbubblesizeval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showleaderlines
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartp->ns_showleaderlinesval ).
        lo_element4 = lo_document->create_simple_element( name = lc_xml_node_firstsliceang
                                                     parent = lo_element3 ).
        lo_element4->set_attribute_ns( name  = 'val'
                                  value = lo_chartp->ns_firstsliceangval ).
      WHEN zcl_excel_drawing=>c_graph_line.
        "----line
        lo_element3 = lo_document->create_simple_element( name = lc_xml_node_linechart
                                                     parent = lo_element2 ).
        lo_element4 = lo_document->create_simple_element( name = lc_xml_node_grouping
                                                     parent = lo_element3 ).
        lo_element4->set_attribute_ns( name  = 'val'
                                  value = lo_chartl->ns_groupingval ).
        lo_element4 = lo_document->create_simple_element( name = lc_xml_node_varycolors
                                                     parent = lo_element3 ).
        lo_element4->set_attribute_ns( name  = 'val'
                                  value = lo_chartl->ns_varycolorsval ).

        "series
        LOOP AT lo_chartl->series INTO ls_serie.
          lo_element4 = lo_document->create_simple_element( name = lc_xml_node_ser
                                                     parent = lo_element3 ).
          lo_element5 = lo_document->create_simple_element( name = lc_xml_node_idx
                                                     parent = lo_element4 ).
          IF ls_serie-idx IS NOT INITIAL.
            lv_str = ls_serie-idx.
          ELSE.
            lv_str = sy-tabix - 1.
          ENDIF.
          CONDENSE lv_str.
          lo_element5->set_attribute_ns( name  = 'val'
                                  value = lv_str ).
          lo_element5 = lo_document->create_simple_element( name = lc_xml_node_order
                                                     parent = lo_element4 ).
          lv_str = ls_serie-order.
          CONDENSE lv_str.
          lo_element5->set_attribute_ns( name  = 'val'
                                  value = lv_str ).
          IF ls_serie-sername IS NOT INITIAL.
            lo_element5 = lo_document->create_simple_element( name = lc_xml_node_tx
                                                      parent = lo_element4 ).
            lo_element6 = lo_document->create_simple_element( name = lc_xml_node_v
                                                      parent = lo_element5 ).
            lo_element6->set_value( value = ls_serie-sername ).
          ENDIF.
          lo_element5 = lo_document->create_simple_element( name = lc_xml_node_marker
                                                     parent = lo_element4 ).
          lo_element6 = lo_document->create_simple_element( name = lc_xml_node_symbol
                                                     parent = lo_element5 ).
          lo_element6->set_attribute_ns( name  = 'val'
                                  value = ls_serie-symbol ).
          IF ls_serie-lbl IS NOT INITIAL.
            lo_element5 = lo_document->create_simple_element( name = lc_xml_node_cat
                                                       parent = lo_element4 ).
            lo_element6 = lo_document->create_simple_element( name = lc_xml_node_strref
                                                       parent = lo_element5 ).
            lo_element7 = lo_document->create_simple_element( name = lc_xml_node_f
                                                       parent = lo_element6 ).
            lo_element7->set_value( value = ls_serie-lbl ).
          ENDIF.
          IF ls_serie-ref IS NOT INITIAL.
            lo_element5 = lo_document->create_simple_element( name = lc_xml_node_val
                                                       parent = lo_element4 ).
            lo_element6 = lo_document->create_simple_element( name = lc_xml_node_numref
                                                       parent = lo_element5 ).
            lo_element7 = lo_document->create_simple_element( name = lc_xml_node_f
                                                       parent = lo_element6 ).
            lo_element7->set_value( value = ls_serie-ref ).
          ENDIF.
          lo_element5 = lo_document->create_simple_element( name = lc_xml_node_smooth
                                                       parent = lo_element4 ).
          lo_element5->set_attribute_ns( name  = 'val'
                                  value = ls_serie-smooth ).
        ENDLOOP.
        "endseries

        lo_element4 = lo_document->create_simple_element( name = lc_xml_node_dlbls
                                                     parent = lo_element3 ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showlegendkey
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartl->ns_showlegendkeyval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showval
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartl->ns_showvalval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showcatname
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartl->ns_showcatnameval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showsername
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartl->ns_showsernameval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showpercent
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartl->ns_showpercentval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showbubblesize
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartl->ns_showbubblesizeval ).

        lo_element4 = lo_document->create_simple_element( name = lc_xml_node_marker
                                                     parent = lo_element3 ).
        lo_element4->set_attribute_ns( name  = 'val'
                                  value = lo_chartl->ns_markerval ).
        lo_element4 = lo_document->create_simple_element( name = lc_xml_node_smooth
                                                     parent = lo_element3 ).
        lo_element4->set_attribute_ns( name  = 'val'
                                  value = lo_chartl->ns_smoothval ).

        "axes
        lo_el_rootchart = lo_element3.
        LOOP AT lo_chartl->axes INTO ls_ax.
          lo_element4 = lo_document->create_simple_element( name = lc_xml_node_axid
                                                     parent = lo_el_rootchart ).
          lo_element4->set_attribute_ns( name  = 'val'
                                  value = ls_ax-axid ).
          CASE ls_ax-type.
            WHEN zcl_excel_graph_line=>c_catax.
              lo_element3 = lo_document->create_simple_element( name = lc_xml_node_catax
                                                     parent = lo_element2 ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_axid
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-axid ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_scaling
                                                     parent = lo_element3 ).
              lo_element5 = lo_document->create_simple_element( name = lc_xml_node_orientation
                                                     parent = lo_element4 ).
              lo_element5->set_attribute_ns( name  = 'val'
                                             value = ls_ax-orientation ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_delete
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-delete ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_axpos
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-axpos ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_majortickmark
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-majortickmark ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_minortickmark
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-minortickmark ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_ticklblpos
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-ticklblpos ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_crossax
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-crossax ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_crosses
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-crosses ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_auto
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-auto ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_lblalgn
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-lblalgn ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_lbloffset
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-lbloffset ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_nomultilvllbl
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-nomultilvllbl ).
            WHEN zcl_excel_graph_line=>c_valax.
              lo_element3 = lo_document->create_simple_element( name = lc_xml_node_valax
                                                     parent = lo_element2 ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_axid
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-axid ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_scaling
                                                     parent = lo_element3 ).
              lo_element5 = lo_document->create_simple_element( name = lc_xml_node_orientation
                                                     parent = lo_element4 ).
              lo_element5->set_attribute_ns( name  = 'val'
                                             value = ls_ax-orientation ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_delete
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-delete ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_axpos
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-axpos ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_majorgridlines
                                                     parent = lo_element3 ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_numfmt
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'formatCode'
                                             value = ls_ax-formatcode ).
              lo_element4->set_attribute_ns( name  = 'sourceLinked'
                                             value = ls_ax-sourcelinked ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_majortickmark
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-majortickmark ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_minortickmark
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-minortickmark ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_ticklblpos
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-ticklblpos ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_crossax
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-crossax ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_crosses
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-crosses ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_crossbetween
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-crossbetween ).
            WHEN OTHERS.
          ENDCASE.
        ENDLOOP.
        "endaxes

      WHEN OTHERS.
    ENDCASE.

    "legend
    IF lo_chart->print_label EQ abap_true.
      lo_element2 = lo_document->create_simple_element( name = lc_xml_node_legend
                                                         parent = lo_element ).
      CASE io_drawing->graph_type.
        WHEN zcl_excel_drawing=>c_graph_bars.
          "----bar
          lo_element3 = lo_document->create_simple_element( name = lc_xml_node_legendpos
                                                       parent = lo_element2 ).
          lo_element3->set_attribute_ns( name  = 'val'
                                    value = lo_chartb->ns_legendposval ).
          lo_element3 = lo_document->create_simple_element( name = lc_xml_node_layout
                                                       parent = lo_element2 ).
          lo_element3 = lo_document->create_simple_element( name = lc_xml_node_overlay
                                                       parent = lo_element2 ).
          lo_element3->set_attribute_ns( name  = 'val'
                                    value = lo_chartb->ns_overlayval ).
        WHEN zcl_excel_drawing=>c_graph_line.
          "----line
          lo_element3 = lo_document->create_simple_element( name = lc_xml_node_legendpos
                                                       parent = lo_element2 ).
          lo_element3->set_attribute_ns( name  = 'val'
                                    value = lo_chartl->ns_legendposval ).
          lo_element3 = lo_document->create_simple_element( name = lc_xml_node_layout
                                                       parent = lo_element2 ).
          lo_element3 = lo_document->create_simple_element( name = lc_xml_node_overlay
                                                       parent = lo_element2 ).
          lo_element3->set_attribute_ns( name  = 'val'
                                    value = lo_chartl->ns_overlayval ).
        WHEN zcl_excel_drawing=>c_graph_pie.
          "----pie
          lo_element3 = lo_document->create_simple_element( name = lc_xml_node_legendpos
                                                       parent = lo_element2 ).
          lo_element3->set_attribute_ns( name  = 'val'
                                    value = lo_chartp->ns_legendposval ).
          lo_element3 = lo_document->create_simple_element( name = lc_xml_node_layout
                                                       parent = lo_element2 ).
          lo_element3 = lo_document->create_simple_element( name = lc_xml_node_overlay
                                                       parent = lo_element2 ).
          lo_element3->set_attribute_ns( name  = 'val'
                                    value = lo_chartp->ns_overlayval ).
          lo_element3 = lo_document->create_simple_element( name = lc_xml_node_txpr
                                                       parent = lo_element2 ).
          lo_element4 = lo_document->create_simple_element( name = lc_xml_node_bodypr
                                                       parent = lo_element3 ).
          lo_element4 = lo_document->create_simple_element( name = lc_xml_node_lststyle
                                                       parent = lo_element3 ).
          lo_element4 = lo_document->create_simple_element( name = lc_xml_node_p
                                                       parent = lo_element3 ).
          lo_element5 = lo_document->create_simple_element( name = lc_xml_node_ppr
                                                       parent = lo_element4 ).
          lo_element5->set_attribute_ns( name  = 'rtl'
                                    value = lo_chartp->ns_pprrtl ).
          lo_element6 = lo_document->create_simple_element( name = lc_xml_node_defrpr
                                                       parent = lo_element5 ).
          lo_element5 = lo_document->create_simple_element( name = lc_xml_node_endpararpr
                                                       parent = lo_element4 ).
          lo_element5->set_attribute_ns( name  = 'lang'
                                    value = lo_chartp->ns_endpararprlang ).
        WHEN OTHERS.
      ENDCASE.
    ENDIF.

    lo_element2 = lo_document->create_simple_element( name = lc_xml_node_plotvisonly
                                                         parent = lo_element ).
    lo_element2->set_attribute_ns( name  = 'val'
                                      value = lo_chart->ns_plotvisonlyval ).
    lo_element2 = lo_document->create_simple_element( name = lc_xml_node_dispblanksas
                                                         parent = lo_element ).
    lo_element2->set_attribute_ns( name  = 'val'
                                      value = lo_chart->ns_dispblanksasval ).
    lo_element2 = lo_document->create_simple_element( name = lc_xml_node_showdlblsovermax
                                                         parent = lo_element ).
    lo_element2->set_attribute_ns( name  = 'val'
                                      value = lo_chart->ns_showdlblsovermaxval ).
    "---------------------------END OF CHART

    "printSettings
    lo_element = lo_document->create_simple_element( name = lc_xml_node_printsettings
                                                         parent = lo_element_root ).
    "headerFooter
    lo_element2 = lo_document->create_simple_element( name = lc_xml_node_headerfooter
                                                         parent = lo_element ).
    "pageMargins
    lo_element2 = lo_document->create_simple_element( name = lc_xml_node_pagemargins
                                                         parent = lo_element ).
    lo_element2->set_attribute_ns( name  = 'b'
                                      value = lo_chart->pagemargins-b ).
    lo_element2->set_attribute_ns( name  = 'l'
                                      value = lo_chart->pagemargins-l ).
    lo_element2->set_attribute_ns( name  = 'r'
                                      value = lo_chart->pagemargins-r ).
    lo_element2->set_attribute_ns( name  = 't'
                                      value = lo_chart->pagemargins-t ).
    lo_element2->set_attribute_ns( name  = 'header'
                                      value = lo_chart->pagemargins-header ).
    lo_element2->set_attribute_ns( name  = 'footer'
                                      value = lo_chart->pagemargins-footer ).
    "pageSetup
    lo_element2 = lo_document->create_simple_element( name = lc_xml_node_pagesetup
                                                         parent = lo_element ).

**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).
  ENDMETHOD.