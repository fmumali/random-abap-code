  METHOD load_worksheet_pagebreaks.

    DATA: lo_node           TYPE REF TO if_ixml_element,
          lo_ixml_rowbreaks TYPE REF TO if_ixml_node_collection,
          lo_ixml_colbreaks TYPE REF TO if_ixml_node_collection,
          lo_ixml_iterator  TYPE REF TO if_ixml_node_iterator,
          lo_ixml_rowbreak  TYPE REF TO if_ixml_element,
          lo_ixml_colbreak  TYPE REF TO if_ixml_element,
          lo_style_cond     TYPE REF TO zcl_excel_style_cond,
          lv_count          TYPE i.


    DATA: lt_pagebreaks TYPE STANDARD TABLE OF zcl_excel_worksheet_pagebreaks=>ts_pagebreak_at,
          lo_pagebreaks TYPE REF TO zcl_excel_worksheet_pagebreaks.

    FIELD-SYMBOLS: <ls_pagebreak_row> LIKE LINE OF lt_pagebreaks.
    FIELD-SYMBOLS: <ls_pagebreak_col> LIKE LINE OF lt_pagebreaks.

*--------------------------------------------------------------------*
* Get minimal number of cells where to add pagebreaks
* Since rows and columns are handled in separate nodes
* Build table to identify these cells
*--------------------------------------------------------------------*
    lo_node ?= io_ixml_worksheet->find_from_name_ns( name = 'rowBreaks' uri = namespace-main ).
    CHECK lo_node IS BOUND.
    lo_ixml_rowbreaks =  lo_node->get_elements_by_tag_name_ns( name = 'brk' uri = namespace-main ).
    lo_ixml_iterator  =  lo_ixml_rowbreaks->create_iterator( ).
    lo_ixml_rowbreak  ?= lo_ixml_iterator->get_next( ).
    WHILE lo_ixml_rowbreak IS BOUND.
      APPEND INITIAL LINE TO lt_pagebreaks ASSIGNING <ls_pagebreak_row>.
      <ls_pagebreak_row>-cell_row = lo_ixml_rowbreak->get_attribute_ns( 'id' ).

      lo_ixml_rowbreak  ?= lo_ixml_iterator->get_next( ).
    ENDWHILE.
    CHECK <ls_pagebreak_row> IS ASSIGNED.

    lo_node ?= io_ixml_worksheet->find_from_name_ns( name = 'colBreaks' uri = namespace-main ).
    CHECK lo_node IS BOUND.
    lo_ixml_colbreaks =  lo_node->get_elements_by_tag_name_ns( name = 'brk' uri = namespace-main ).
    lo_ixml_iterator  =  lo_ixml_colbreaks->create_iterator( ).
    lo_ixml_colbreak  ?= lo_ixml_iterator->get_next( ).
    CLEAR lv_count.
    WHILE lo_ixml_colbreak IS BOUND.
      ADD 1 TO lv_count.
      READ TABLE lt_pagebreaks INDEX lv_count ASSIGNING <ls_pagebreak_col>.
      IF sy-subrc <> 0.
        APPEND INITIAL LINE TO lt_pagebreaks ASSIGNING <ls_pagebreak_col>.
        <ls_pagebreak_col>-cell_row = <ls_pagebreak_row>-cell_row.
      ENDIF.
      <ls_pagebreak_col>-cell_column = lo_ixml_colbreak->get_attribute_ns( 'id' ).

      lo_ixml_colbreak  ?= lo_ixml_iterator->get_next( ).
    ENDWHILE.
*--------------------------------------------------------------------*
* Finally add each pagebreak
*--------------------------------------------------------------------*
    lo_pagebreaks = io_worksheet->get_pagebreaks( ).
    LOOP AT lt_pagebreaks ASSIGNING <ls_pagebreak_row>.
      lo_pagebreaks->add_pagebreak( ip_column = <ls_pagebreak_row>-cell_column
                                    ip_row    = <ls_pagebreak_row>-cell_row ).
    ENDLOOP.


  ENDMETHOD.