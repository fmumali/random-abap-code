  METHOD create_xl_sheet_pagebreaks.
    DATA: lo_pagebreaks     TYPE REF TO zcl_excel_worksheet_pagebreaks,
          lt_pagebreaks     TYPE zcl_excel_worksheet_pagebreaks=>tt_pagebreak_at,
          lt_rows           TYPE HASHED TABLE OF int4 WITH UNIQUE KEY table_line,
          lt_columns        TYPE HASHED TABLE OF int4 WITH UNIQUE KEY table_line,

          lo_node_rowbreaks TYPE REF TO if_ixml_element,
          lo_node_colbreaks TYPE REF TO if_ixml_element,
          lo_node_break     TYPE REF TO if_ixml_element,

          lv_value          TYPE string.


    FIELD-SYMBOLS: <ls_pagebreak> LIKE LINE OF lt_pagebreaks.

    lo_pagebreaks = io_worksheet->get_pagebreaks( ).
    CHECK lo_pagebreaks IS BOUND.

    lt_pagebreaks = lo_pagebreaks->get_all_pagebreaks( ).
    CHECK lt_pagebreaks IS NOT INITIAL.  " No need to proceed if don't have any pagebreaks.

    lo_node_rowbreaks = io_document->create_simple_element( name   = 'rowBreaks'
                                                            parent = io_document ).

    lo_node_colbreaks = io_document->create_simple_element( name   = 'colBreaks'
                                                            parent = io_document ).


    LOOP AT lt_pagebreaks ASSIGNING <ls_pagebreak>.

* Count how many rows and columns need to be broken
      INSERT <ls_pagebreak>-cell_row    INTO TABLE lt_rows.
      IF sy-subrc = 0. " New
        lv_value = <ls_pagebreak>-cell_row.
        CONDENSE lv_value.

        lo_node_break = io_document->create_simple_element( name   = 'brk'
                                                            parent = io_document ).
        lo_node_break->set_attribute( name = 'id'  value = lv_value ).
        lo_node_break->set_attribute( name = 'man' value = '1' ).      " Manual break
        lo_node_break->set_attribute( name = 'max' value = '16383' ).  " Max columns

        lo_node_rowbreaks->append_child( new_child = lo_node_break ).
      ENDIF.

      INSERT <ls_pagebreak>-cell_column INTO TABLE lt_columns.
      IF sy-subrc = 0. " New
        lv_value = <ls_pagebreak>-cell_column.
        CONDENSE lv_value.

        lo_node_break = io_document->create_simple_element( name   = 'brk'
                                                            parent = io_document ).
        lo_node_break->set_attribute( name = 'id'  value = lv_value ).
        lo_node_break->set_attribute( name = 'man' value = '1' ).        " Manual break
        lo_node_break->set_attribute( name = 'max' value = '1048575' ).  " Max rows

        lo_node_colbreaks->append_child( new_child = lo_node_break ).
      ENDIF.


    ENDLOOP.

    lv_value = lines( lt_rows ).
    CONDENSE lv_value.
    lo_node_rowbreaks->set_attribute( name = 'count'             value = lv_value ).
    lo_node_rowbreaks->set_attribute( name = 'manualBreakCount'  value = lv_value ).

    lv_value = lines( lt_rows ).
    CONDENSE lv_value.
    lo_node_colbreaks->set_attribute( name = 'count'             value = lv_value ).
    lo_node_colbreaks->set_attribute( name = 'manualBreakCount'  value = lv_value ).

    io_parent->append_child( new_child = lo_node_rowbreaks ).
    io_parent->append_child( new_child = lo_node_colbreaks ).

  ENDMETHOD.