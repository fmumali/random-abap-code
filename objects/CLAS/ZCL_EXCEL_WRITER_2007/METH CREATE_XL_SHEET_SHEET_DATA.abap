  METHOD create_xl_sheet_sheet_data.

    TYPES: BEGIN OF lty_table_area,
             left   TYPE i,
             right  TYPE i,
             top    TYPE i,
             bottom TYPE i,
           END OF lty_table_area.

    CONSTANTS: lc_dummy_cell_content       TYPE zexcel_s_cell_data-cell_value VALUE '})~~~ This is a dummy value for ABAP2XLSX and you should never find this in a real excelsheet Ihope'.

    CONSTANTS: lc_xml_node_sheetdata TYPE string VALUE 'sheetData',   " SheetData tag
               lc_xml_node_row       TYPE string VALUE 'row',         " Row tag
               lc_xml_attr_r         TYPE string VALUE 'r',           " Cell:  row-attribute
               lc_xml_attr_spans     TYPE string VALUE 'spans',       " Cell: spans-attribute
               lc_xml_node_c         TYPE string VALUE 'c',           " Cell tag
               lc_xml_node_v         TYPE string VALUE 'v',           " Cell: value
               lc_xml_node_f         TYPE string VALUE 'f',           " Cell: formula
               lc_xml_attr_s         TYPE string VALUE 's',           " Cell: style
               lc_xml_attr_t         TYPE string VALUE 't'.           " Cell: type

    DATA: col_count              TYPE int4,
          lo_autofilters         TYPE REF TO zcl_excel_autofilters,
          lo_autofilter          TYPE REF TO zcl_excel_autofilter,
          l_autofilter_hidden    TYPE flag,
          lt_values              TYPE zexcel_t_autofilter_values,
          ls_values              TYPE zexcel_s_autofilter_values,
          ls_area                TYPE zexcel_s_autofilter_area,

          lo_iterator            TYPE REF TO zcl_excel_collection_iterator,
          lo_table               TYPE REF TO zcl_excel_table,
          lt_table_areas         TYPE SORTED TABLE OF lty_table_area WITH NON-UNIQUE KEY left right top bottom,
          ls_table_area          LIKE LINE OF lt_table_areas,
          lo_column              TYPE REF TO zcl_excel_column,

          ls_sheet_content       LIKE LINE OF io_worksheet->sheet_content,
          ls_sheet_content_empty LIKE LINE OF io_worksheet->sheet_content,
          lv_current_row         TYPE i,
          lv_next_row            TYPE i,
          lv_last_row            TYPE i,

*        lts_row_dimensions     TYPE zexcel_t_worksheet_rowdimensio,
          lo_row_iterator        TYPE REF TO zcl_excel_collection_iterator,
          lo_row                 TYPE REF TO zcl_excel_row,
          lo_row_empty           TYPE REF TO zcl_excel_row,
          lts_row_outlines       TYPE zcl_excel_worksheet=>mty_ts_outlines_row,

          ls_last_row            TYPE zexcel_s_cell_data,
          ls_style_mapping       TYPE zexcel_s_styles_mapping,

          lo_element_2           TYPE REF TO if_ixml_element,
          lo_element_3           TYPE REF TO if_ixml_element,
          lo_element_4           TYPE REF TO if_ixml_element,

          lv_value               TYPE string,
          lv_style_guid          TYPE zexcel_cell_style.
    DATA: lt_column_formulas_used TYPE mty_column_formulas_used,
          lv_si                   TYPE i.

    FIELD-SYMBOLS: <ls_sheet_content> TYPE zexcel_s_cell_data,
                   <ls_row_outline>   LIKE LINE OF lts_row_outlines.


    " sheetData node
    rv_ixml_sheet_data_root = io_document->create_simple_element( name   = lc_xml_node_sheetdata
                                                                  parent = io_document ).

    " Get column count
    col_count      = io_worksheet->get_highest_column( ).
    " Get autofilter
    lo_autofilters = excel->get_autofilters_reference( ).
    lo_autofilter  = lo_autofilters->get( io_worksheet = io_worksheet ) .
    IF lo_autofilter IS BOUND.
      lt_values           = lo_autofilter->get_values( ) .
      ls_area             = lo_autofilter->get_filter_area( ) .
      l_autofilter_hidden = abap_true. " First defautl is not showing
    ENDIF.
*--------------------------------------------------------------------*
*issue #220 - If cell in tables-area don't use default from row or column or sheet - Coding 1 - start
*--------------------------------------------------------------------*
*Build table to hold all table-areas attached to this sheet
    lo_iterator = io_worksheet->get_tables_iterator( ).
    WHILE lo_iterator->has_next( ) EQ abap_true.
      lo_table ?= lo_iterator->get_next( ).
      ls_table_area-left   = zcl_excel_common=>convert_column2int( lo_table->settings-top_left_column ).
      ls_table_area-right  = lo_table->get_right_column_integer( ).
      ls_table_area-top    = lo_table->settings-top_left_row.
      ls_table_area-bottom = lo_table->get_bottom_row_integer( ).
      INSERT ls_table_area INTO TABLE lt_table_areas.
    ENDWHILE.
*--------------------------------------------------------------------*
*issue #220 - If cell in tables-area don't use default from row or column or sheet - Coding 1 - end
*--------------------------------------------------------------------*
*We have problems when the first rows or trailing rows are not set but we have rowinformation
*to solve this we add dummycontent into first and last line that will not be set
*Set first line if necessary
    READ TABLE io_worksheet->sheet_content TRANSPORTING NO FIELDS WITH KEY cell_row = 1.
    IF sy-subrc <> 0.
      ls_sheet_content_empty-cell_row      = 1.
      ls_sheet_content_empty-cell_column   = 1.
      ls_sheet_content_empty-cell_value    = lc_dummy_cell_content.
      INSERT ls_sheet_content_empty INTO TABLE io_worksheet->sheet_content.
    ENDIF.
*Set last line if necessary
*Last row with cell content
    lv_last_row = io_worksheet->get_highest_row( ).
*Last line with row-information set directly ( like line height, hidden-status ... )

    lo_row_iterator = io_worksheet->get_rows_iterator( ).
    WHILE lo_row_iterator->has_next( ) = abap_true.
      lo_row ?= lo_row_iterator->get_next( ).
      IF lo_row->get_row_index( ) > lv_last_row.
        lv_last_row = lo_row->get_row_index( ).
      ENDIF.
    ENDWHILE.

*Last line with row-information set indirectly by row outline
    lts_row_outlines = io_worksheet->get_row_outlines( ).
    LOOP AT lts_row_outlines ASSIGNING <ls_row_outline>.
      IF <ls_row_outline>-collapsed = 'X'.
        lv_current_row = <ls_row_outline>-row_to + 1.  " collapsed-status may be set on following row
      ELSE.
        lv_current_row = <ls_row_outline>-row_to.  " collapsed-status may be set on following row
      ENDIF.
      IF lv_current_row > lv_last_row.
        lv_last_row = lv_current_row.
      ENDIF.
    ENDLOOP.
    READ TABLE io_worksheet->sheet_content TRANSPORTING NO FIELDS WITH KEY cell_row = lv_last_row.
    IF sy-subrc <> 0.
      ls_sheet_content_empty-cell_row      = lv_last_row.
      ls_sheet_content_empty-cell_column   = 1.
      ls_sheet_content_empty-cell_value    = lc_dummy_cell_content.
      INSERT ls_sheet_content_empty INTO TABLE io_worksheet->sheet_content.
    ENDIF.

    CLEAR ls_sheet_content.
    LOOP AT io_worksheet->sheet_content INTO ls_sheet_content.
      IF lt_values IS INITIAL. " no values attached to autofilter  " issue #368 autofilter filtering too much
        CLEAR l_autofilter_hidden.
      ELSE.
        READ TABLE lt_values INTO ls_values WITH KEY column = ls_last_row-cell_column.
        IF sy-subrc = 0 AND ls_values-value = ls_last_row-cell_value.
          CLEAR l_autofilter_hidden.
        ENDIF.
      ENDIF.
      CLEAR ls_style_mapping.
*Create row element
*issues #346,#154, #195  - problems when we have information in row_dimension but no cell content in that row
*Get next line that may have to be added.  If we have empty lines this is the next line after previous cell content
*Otherwise it is the line of the current cell content
      lv_current_row = ls_last_row-cell_row + 1.
      IF lv_current_row > ls_sheet_content-cell_row.
        lv_current_row = ls_sheet_content-cell_row.
      ENDIF.
*Fill in empty lines if necessary - assign an emtpy sheet content
      lv_next_row = lv_current_row.
      WHILE lv_next_row <= ls_sheet_content-cell_row.
        lv_current_row = lv_next_row.
        lv_next_row = lv_current_row + 1.
        IF lv_current_row = ls_sheet_content-cell_row. " cell value found in this row
          ASSIGN ls_sheet_content TO <ls_sheet_content>.
        ELSE.
*Check if empty row is really necessary - this is basically the case when we have information in row_dimension
          lo_row_empty = io_worksheet->get_row( lv_current_row ).
          CHECK lo_row_empty->get_row_height( )                 >= 0          OR
                lo_row_empty->get_collapsed( io_worksheet )      = abap_true  OR
                lo_row_empty->get_outline_level( io_worksheet )  > 0          OR
                lo_row_empty->get_xf_index( )                   <> 0.
          " Dummyentry A1
          ls_sheet_content_empty-cell_row      = lv_current_row.
          ls_sheet_content_empty-cell_column   = 1.
          ASSIGN ls_sheet_content_empty TO <ls_sheet_content>.
        ENDIF.

        IF ls_last_row-cell_row NE <ls_sheet_content>-cell_row.
          IF lo_autofilter IS BOUND.
            IF ls_area-row_start >=  ls_last_row-cell_row OR " One less for header
              ls_area-row_end   < ls_last_row-cell_row .
              CLEAR l_autofilter_hidden.
            ENDIF.
          ELSE.
            CLEAR l_autofilter_hidden.
          ENDIF.
          IF ls_last_row-cell_row IS NOT INITIAL.
            " Row visibility of previos row.
            IF lo_row->get_visible( io_worksheet ) = abap_false OR
               l_autofilter_hidden = abap_true.
              lo_element_2->set_attribute_ns( name  = 'hidden' value = 'true' ).
            ENDIF.
            rv_ixml_sheet_data_root->append_child( new_child = lo_element_2 ). " row node
          ENDIF.
          " Add new row
          lo_element_2 = io_document->create_simple_element( name   = lc_xml_node_row
                                                             parent = io_document ).
          " r
          lv_value = <ls_sheet_content>-cell_row.
          SHIFT lv_value RIGHT DELETING TRAILING space.
          SHIFT lv_value LEFT DELETING LEADING space.

          lo_element_2->set_attribute_ns( name  = lc_xml_attr_r
                                          value = lv_value ).
          " Spans
          lv_value = col_count.
          CONCATENATE '1:' lv_value INTO lv_value.
          SHIFT lv_value RIGHT DELETING TRAILING space.
          SHIFT lv_value LEFT DELETING LEADING space.
          lo_element_2->set_attribute_ns( name  = lc_xml_attr_spans
                                          value = lv_value ).
          lo_row = io_worksheet->get_row( <ls_sheet_content>-cell_row ).
          " Row dimensions
          IF lo_row->get_custom_height( ) = abap_true.
            lo_element_2->set_attribute_ns( name  = 'customHeight' value = '1' ).
          ENDIF.
          IF lo_row->get_row_height( ) > 0.
            lv_value = lo_row->get_row_height( ).
            lo_element_2->set_attribute_ns( name  = 'ht' value = lv_value ).
          ENDIF.
          " Collapsed
          IF lo_row->get_collapsed( io_worksheet ) = abap_true.
            lo_element_2->set_attribute_ns( name  = 'collapsed' value = 'true' ).
          ENDIF.
          " Outline level
          IF lo_row->get_outline_level( io_worksheet ) > 0.
            lv_value = lo_row->get_outline_level( io_worksheet ).
            SHIFT lv_value RIGHT DELETING TRAILING space.
            SHIFT lv_value LEFT DELETING LEADING space.
            lo_element_2->set_attribute_ns( name  = 'outlineLevel' value = lv_value ).
          ENDIF.
          " Style
          IF lo_row->get_xf_index( ) <> 0.
            lv_value = lo_row->get_xf_index( ).
            lo_element_2->set_attribute_ns( name  = 's' value = lv_value ).
            lo_element_2->set_attribute_ns( name  = 'customFormat'  value = '1' ).
          ENDIF.
          IF lt_values IS INITIAL. " no values attached to autofilter  " issue #368 autofilter filtering too much
            CLEAR l_autofilter_hidden.
          ELSE.
            l_autofilter_hidden = abap_true. " First default is not showing
          ENDIF.
        ELSE.

        ENDIF.
      ENDWHILE.

      lo_element_3 = io_document->create_simple_element( name   = lc_xml_node_c
                                                         parent = io_document ).

      lo_element_3->set_attribute_ns( name  = lc_xml_attr_r
                                      value = <ls_sheet_content>-cell_coords ).

*begin of change issue #157 - allow column cellstyle
*if no cellstyle is set, look into column, then into sheet
      IF <ls_sheet_content>-cell_style IS NOT INITIAL.
        lv_style_guid = <ls_sheet_content>-cell_style.
      ELSE.
*--------------------------------------------------------------------*
*issue #220 - If cell in tables-area don't use default from row or column or sheet - Coding 2 - start
*--------------------------------------------------------------------*
*Check if cell in any of the table areas
        LOOP AT lt_table_areas TRANSPORTING NO FIELDS WHERE top    <= <ls_sheet_content>-cell_row
                                                        AND bottom >= <ls_sheet_content>-cell_row
                                                        AND left   <= <ls_sheet_content>-cell_column
                                                        AND right  >= <ls_sheet_content>-cell_column.
          EXIT.
        ENDLOOP.
        IF sy-subrc = 0.
          CLEAR lv_style_guid.     " No style --> EXCEL will use built-in-styles as declared in the tables-section
        ELSE.
*--------------------------------------------------------------------*
*issue #220 - If cell in tables-area don't use default from row or column or sheet - Coding 2 - end
*--------------------------------------------------------------------*
          lv_style_guid = io_worksheet->zif_excel_sheet_properties~get_style( ).
          lo_column ?= io_worksheet->get_column( <ls_sheet_content>-cell_column ).
          IF lo_column->get_column_index( ) = <ls_sheet_content>-cell_column.
            lv_style_guid = lo_column->get_column_style_guid( ).
            IF lv_style_guid IS INITIAL.
              lv_style_guid = io_worksheet->zif_excel_sheet_properties~get_style( ).
            ENDIF.
          ENDIF.

*--------------------------------------------------------------------*
*issue #220 - If cell in tables-area don't use default from row or column or sheet - Coding 3 - start
*--------------------------------------------------------------------*
        ENDIF.
*--------------------------------------------------------------------*
*issue #220 - If cell in tables-area don't use default from row or column or sheet - Coding 3 - end
*--------------------------------------------------------------------*
      ENDIF.
      IF lv_style_guid IS NOT INITIAL.
        READ TABLE styles_mapping INTO ls_style_mapping WITH KEY guid = lv_style_guid.
*end of change issue #157 - allow column cellstyles
        lv_value = ls_style_mapping-style.
        SHIFT lv_value RIGHT DELETING TRAILING space.
        SHIFT lv_value LEFT DELETING LEADING space.
        lo_element_3->set_attribute_ns( name  = lc_xml_attr_s
                                        value = lv_value ).
      ENDIF.

      " For cells with formula ignore the value - Excel will calculate it
      IF <ls_sheet_content>-cell_formula IS NOT INITIAL.
        " fomula node
        lo_element_4 = io_document->create_simple_element( name   = lc_xml_node_f
                                                           parent = io_document ).
        lo_element_4->set_value( value = <ls_sheet_content>-cell_formula ).
        lo_element_3->append_child( new_child = lo_element_4 ). " formula node
      ELSEIF <ls_sheet_content>-column_formula_id <> 0.
        create_xl_sheet_column_formula(
          EXPORTING
            io_document             = io_document
            it_column_formulas      = io_worksheet->column_formulas
            is_sheet_content        = <ls_sheet_content>
          IMPORTING
            eo_element              = lo_element_4
          CHANGING
            ct_column_formulas_used = lt_column_formulas_used
            cv_si                   = lv_si ).
        lo_element_3->append_child( new_child = lo_element_4 ).
      ELSEIF <ls_sheet_content>-cell_value IS NOT INITIAL           "cell can have just style or formula
         AND <ls_sheet_content>-cell_value <> lc_dummy_cell_content.
        IF <ls_sheet_content>-data_type IS NOT INITIAL.
          IF <ls_sheet_content>-data_type EQ 's_leading_blanks'.
            lo_element_3->set_attribute_ns( name  = lc_xml_attr_t
                                            value = 's' ).
          ELSE.
            lo_element_3->set_attribute_ns( name  = lc_xml_attr_t
                                            value = <ls_sheet_content>-data_type ).
          ENDIF.
        ENDIF.

        " value node
        lo_element_4 = io_document->create_simple_element( name   = lc_xml_node_v
                                                           parent = io_document ).

        IF <ls_sheet_content>-data_type EQ 's' OR <ls_sheet_content>-data_type EQ 's_leading_blanks'.
          lv_value = me->get_shared_string_index( ip_cell_value = <ls_sheet_content>-cell_value
                                                  it_rtf        = <ls_sheet_content>-rtf_tab ).
          CONDENSE lv_value.
          lo_element_4->set_value( value = lv_value ).
        ELSE.
          lv_value = <ls_sheet_content>-cell_value.
          CONDENSE lv_value.
          lo_element_4->set_value( value = lv_value ).
        ENDIF.

        lo_element_3->append_child( new_child = lo_element_4 ). " value node
      ENDIF.

      lo_element_2->append_child( new_child = lo_element_3 ). " column node
      ls_last_row = <ls_sheet_content>.
    ENDLOOP.
    IF sy-subrc = 0.
      READ TABLE lt_values INTO ls_values WITH KEY column = ls_last_row-cell_column.
      IF sy-subrc = 0 AND ls_values-value = ls_last_row-cell_value.
        CLEAR l_autofilter_hidden.
      ENDIF.
      IF lo_autofilter IS BOUND.
        IF ls_area-row_start >=  ls_last_row-cell_row OR " One less for header
          ls_area-row_end   < ls_last_row-cell_row .
          CLEAR l_autofilter_hidden.
        ENDIF.
      ELSE.
        CLEAR l_autofilter_hidden.
      ENDIF.
      " Row visibility of previos row.
      IF lo_row->get_visible( ) = abap_false OR
         l_autofilter_hidden = abap_true.
        lo_element_2->set_attribute_ns( name  = 'hidden' value = 'true' ).
      ENDIF.
      rv_ixml_sheet_data_root->append_child( new_child = lo_element_2 ). " row node
    ENDIF.
    DELETE io_worksheet->sheet_content WHERE cell_value = lc_dummy_cell_content.  " Get rid of dummyentries

  ENDMETHOD.