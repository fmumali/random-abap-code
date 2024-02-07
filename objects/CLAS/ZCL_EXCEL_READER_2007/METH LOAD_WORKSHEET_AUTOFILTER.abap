  METHOD load_worksheet_autofilter.

    TYPES: BEGIN OF lty_autofilter,
             ref TYPE string,
           END OF lty_autofilter.

    DATA: lo_ixml_autofilter_elem    TYPE REF TO if_ixml_element,
          lv_ref                     TYPE string,
          lo_ixml_filter_column_coll TYPE REF TO if_ixml_node_collection,
          lo_ixml_filter_column_iter TYPE REF TO if_ixml_node_iterator,
          lo_ixml_filter_column      TYPE REF TO if_ixml_element,
          lv_col_id                  TYPE i,
          lv_column                  TYPE zexcel_cell_column,
          lo_ixml_filters_coll       TYPE REF TO if_ixml_node_collection,
          lo_ixml_filters_iter       TYPE REF TO if_ixml_node_iterator,
          lo_ixml_filters            TYPE REF TO if_ixml_element,
          lo_ixml_filter_coll        TYPE REF TO if_ixml_node_collection,
          lo_ixml_filter_iter        TYPE REF TO if_ixml_node_iterator,
          lo_ixml_filter             TYPE REF TO if_ixml_element,
          lv_val                     TYPE string,
          lo_autofilters             TYPE REF TO zcl_excel_autofilters,
          lo_autofilter              TYPE REF TO zcl_excel_autofilter.

    lo_autofilters = io_worksheet->excel->get_autofilters_reference( ).

    lo_ixml_autofilter_elem = io_ixml_worksheet->find_from_name_ns( name = 'autoFilter' uri = namespace-main ).
    IF lo_ixml_autofilter_elem IS BOUND.
      lv_ref = lo_ixml_autofilter_elem->get_attribute_ns( 'ref' ).

      lo_ixml_filter_column_coll = lo_ixml_autofilter_elem->get_elements_by_tag_name_ns( name = 'filterColumn' uri = namespace-main ).
      lo_ixml_filter_column_iter = lo_ixml_filter_column_coll->create_iterator( ).
      lo_ixml_filter_column ?= lo_ixml_filter_column_iter->get_next( ).
      WHILE lo_ixml_filter_column IS BOUND.
        lv_col_id = lo_ixml_filter_column->get_attribute_ns( 'colId' ).
        lv_column = lv_col_id + 1.

        lo_ixml_filters_coll = lo_ixml_filter_column->get_elements_by_tag_name_ns( name = 'filters' uri = namespace-main ).
        lo_ixml_filters_iter = lo_ixml_filters_coll->create_iterator( ).
        lo_ixml_filters ?= lo_ixml_filters_iter->get_next( ).
        WHILE lo_ixml_filters IS BOUND.

          lo_ixml_filter_coll = lo_ixml_filter_column->get_elements_by_tag_name_ns( name = 'filter' uri = namespace-main ).
          lo_ixml_filter_iter = lo_ixml_filter_coll->create_iterator( ).
          lo_ixml_filter ?= lo_ixml_filter_iter->get_next( ).
          WHILE lo_ixml_filter IS BOUND.
            lv_val = lo_ixml_filter->get_attribute_ns( 'val' ).

            lo_autofilter = lo_autofilters->get( io_worksheet = io_worksheet ).
            IF lo_autofilter IS NOT BOUND.
              lo_autofilter = lo_autofilters->add( io_sheet = io_worksheet ).
            ENDIF.
            lo_autofilter->set_value(
                    i_column = lv_column
                    i_value  = lv_val ).

            lo_ixml_filter ?= lo_ixml_filter_iter->get_next( ).
          ENDWHILE.

          lo_ixml_filters ?= lo_ixml_filters_iter->get_next( ).
        ENDWHILE.

        lo_ixml_filter_column ?= lo_ixml_filter_column_iter->get_next( ).
      ENDWHILE.
    ENDIF.

  ENDMETHOD.