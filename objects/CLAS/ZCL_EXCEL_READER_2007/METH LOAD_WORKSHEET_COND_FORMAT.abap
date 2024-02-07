  METHOD load_worksheet_cond_format.

    DATA: lo_ixml_cond_formats TYPE REF TO if_ixml_node_collection,
          lo_ixml_cond_format  TYPE REF TO if_ixml_element,
          lo_ixml_iterator     TYPE REF TO if_ixml_node_iterator,
          lo_ixml_rules        TYPE REF TO if_ixml_node_collection,
          lo_ixml_rule         TYPE REF TO if_ixml_element,
          lo_ixml_iterator2    TYPE REF TO if_ixml_node_iterator,
          lo_style_cond        TYPE REF TO zcl_excel_style_cond,
          lo_style_cond2       TYPE REF TO zcl_excel_style_cond.


    DATA: lv_area           TYPE string,
          lt_areas          TYPE STANDARD TABLE OF string WITH NON-UNIQUE DEFAULT KEY,
          lv_area_start_row TYPE zexcel_cell_row,
          lv_area_end_row   TYPE zexcel_cell_row,
          lv_area_start_col TYPE zexcel_cell_column_alpha,
          lv_area_end_col   TYPE zexcel_cell_column_alpha,
          lv_rule           TYPE zexcel_condition_rule.


    lo_ixml_cond_formats =  io_ixml_worksheet->get_elements_by_tag_name_ns( name = 'conditionalFormatting' uri = namespace-main ).
    lo_ixml_iterator     =  lo_ixml_cond_formats->create_iterator( ).
    lo_ixml_cond_format  ?= lo_ixml_iterator->get_next( ).

    WHILE lo_ixml_cond_format IS BOUND.

      CLEAR: lv_area,
             lo_ixml_rule,
             lo_style_cond.

*--------------------------------------------------------------------*
* Get type of rule
*--------------------------------------------------------------------*
      lo_ixml_rules       =  lo_ixml_cond_format->get_elements_by_tag_name_ns( name = 'cfRule' uri = namespace-main ).
      lo_ixml_iterator2   =  lo_ixml_rules->create_iterator( ).
      lo_ixml_rule        ?= lo_ixml_iterator2->get_next( ).

      WHILE lo_ixml_rule IS BOUND.
        lv_rule = lo_ixml_rule->get_attribute_ns( 'type' ).
        CLEAR lo_style_cond.

*--------------------------------------------------------------------*
* Depending on ruletype get additional information
*--------------------------------------------------------------------*
        CASE lv_rule.

          WHEN zcl_excel_style_cond=>c_rule_cellis.
            lo_style_cond = io_worksheet->add_new_style_cond( '' ).
            load_worksheet_cond_format_ci( io_ixml_rule  = lo_ixml_rule
                                           io_style_cond = lo_style_cond ).

          WHEN zcl_excel_style_cond=>c_rule_databar.
            lo_style_cond = io_worksheet->add_new_style_cond( '' ).
            load_worksheet_cond_format_db( io_ixml_rule  = lo_ixml_rule
                                           io_style_cond = lo_style_cond ).

          WHEN zcl_excel_style_cond=>c_rule_expression.
            lo_style_cond = io_worksheet->add_new_style_cond( '' ).
            load_worksheet_cond_format_ex( io_ixml_rule  = lo_ixml_rule
                                           io_style_cond = lo_style_cond ).

          WHEN zcl_excel_style_cond=>c_rule_iconset.
            lo_style_cond = io_worksheet->add_new_style_cond( '' ).
            load_worksheet_cond_format_is( io_ixml_rule  = lo_ixml_rule
                                           io_style_cond = lo_style_cond ).

          WHEN zcl_excel_style_cond=>c_rule_colorscale.
            lo_style_cond = io_worksheet->add_new_style_cond( '' ).
            load_worksheet_cond_format_cs( io_ixml_rule  = lo_ixml_rule
                                           io_style_cond = lo_style_cond ).

          WHEN zcl_excel_style_cond=>c_rule_top10.
            lo_style_cond = io_worksheet->add_new_style_cond( '' ).
            load_worksheet_cond_format_t10( io_ixml_rule  = lo_ixml_rule
                                           io_style_cond = lo_style_cond ).

          WHEN zcl_excel_style_cond=>c_rule_above_average.
            lo_style_cond = io_worksheet->add_new_style_cond( '' ).
            load_worksheet_cond_format_aa(  io_ixml_rule  = lo_ixml_rule
                                           io_style_cond = lo_style_cond ).
          WHEN OTHERS.
        ENDCASE.

        IF lo_style_cond IS BOUND.
          lo_style_cond->rule      = lv_rule.
          lo_style_cond->priority  = lo_ixml_rule->get_attribute_ns( 'priority' ).
*--------------------------------------------------------------------*
* Set area to which conditional formatting belongs
*--------------------------------------------------------------------*
          lv_area =  lo_ixml_cond_format->get_attribute_ns( 'sqref' ).
          SPLIT lv_area AT space INTO TABLE lt_areas.
          DELETE lt_areas WHERE table_line IS INITIAL.
          LOOP AT lt_areas INTO lv_area.

            zcl_excel_common=>convert_range2column_a_row( EXPORTING i_range        = lv_area
                                                          IMPORTING e_column_start = lv_area_start_col
                                                                    e_column_end   = lv_area_end_col
                                                                    e_row_start    = lv_area_start_row
                                                                    e_row_end      = lv_area_end_row   ).
            lo_style_cond->add_range( ip_start_column = lv_area_start_col
                                      ip_stop_column  = lv_area_end_col
                                      ip_start_row    = lv_area_start_row
                                      ip_stop_row     = lv_area_end_row   ).
          ENDLOOP.

        ENDIF.
        lo_ixml_rule        ?= lo_ixml_iterator2->get_next( ).
      ENDWHILE.


      lo_ixml_cond_format ?= lo_ixml_iterator->get_next( ).

    ENDWHILE.

  ENDMETHOD.