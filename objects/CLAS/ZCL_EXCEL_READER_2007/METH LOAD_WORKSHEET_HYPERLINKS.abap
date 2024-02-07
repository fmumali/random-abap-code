  METHOD load_worksheet_hyperlinks.

    DATA: lo_ixml_hyperlinks TYPE REF TO if_ixml_node_collection,
          lo_ixml_hyperlink  TYPE REF TO if_ixml_element,
          lo_ixml_iterator   TYPE REF TO if_ixml_node_iterator,
          lv_row_start       TYPE zexcel_cell_row,
          lv_row_end         TYPE zexcel_cell_row,
          lv_column_start    TYPE zexcel_cell_column_alpha,
          lv_column_end      TYPE zexcel_cell_column_alpha,
          lv_is_internal     TYPE abap_bool,
          lv_url             TYPE string,
          lv_value           TYPE zexcel_cell_value.

    DATA: BEGIN OF ls_hyperlink,
            ref      TYPE string,
            display  TYPE string,
            location TYPE string,
            tooltip  TYPE string,
            r_id     TYPE string,
          END OF ls_hyperlink.

    FIELD-SYMBOLS: <ls_external_hyperlink> LIKE LINE OF it_external_hyperlinks.

    lo_ixml_hyperlinks =  io_ixml_worksheet->get_elements_by_tag_name_ns( name = 'hyperlink' uri = namespace-main ).
    lo_ixml_iterator   =  lo_ixml_hyperlinks->create_iterator( ).
    lo_ixml_hyperlink  ?= lo_ixml_iterator->get_next( ).
    WHILE lo_ixml_hyperlink IS BOUND.

      CLEAR ls_hyperlink.
      CLEAR lv_url.

      ls_hyperlink-ref      = lo_ixml_hyperlink->get_attribute_ns( 'ref' ).
      ls_hyperlink-display  = lo_ixml_hyperlink->get_attribute_ns( 'display' ).
      ls_hyperlink-location = lo_ixml_hyperlink->get_attribute_ns( 'location' ).
      ls_hyperlink-tooltip  = lo_ixml_hyperlink->get_attribute_ns( 'tooltip' ).
      ls_hyperlink-r_id     = lo_ixml_hyperlink->get_attribute_ns( name = 'id' uri = namespace-r ).
      IF ls_hyperlink-r_id IS INITIAL.  " Internal link
        lv_is_internal = abap_true.
        lv_url = ls_hyperlink-location.
      ELSE.                             " External link
        READ TABLE it_external_hyperlinks ASSIGNING <ls_external_hyperlink> WITH TABLE KEY id = ls_hyperlink-r_id.
        IF sy-subrc = 0.
          lv_is_internal = abap_false.
          lv_url = <ls_external_hyperlink>-target.
        ENDIF.
      ENDIF.

      IF lv_url IS NOT INITIAL.  " because of unsupported external links

        zcl_excel_common=>convert_range2column_a_row(
          EXPORTING
            i_range        = ls_hyperlink-ref
          IMPORTING
            e_column_start = lv_column_start
            e_column_end   = lv_column_end
            e_row_start    = lv_row_start
            e_row_end      = lv_row_end ).

        io_worksheet->set_area_hyperlink(
          EXPORTING
            ip_column_start = lv_column_start
            ip_column_end   = lv_column_end
            ip_row          = lv_row_start
            ip_row_to       = lv_row_end
            ip_url          = lv_url
            ip_is_internal  = lv_is_internal ).

      ENDIF.

      lo_ixml_hyperlink ?= lo_ixml_iterator->get_next( ).

    ENDWHILE.


  ENDMETHOD.