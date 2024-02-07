  METHOD load_worksheet_cond_format_ex.
    DATA: lo_ixml_nodes      TYPE REF TO if_ixml_node_collection,
          lo_ixml_iterator   TYPE REF TO if_ixml_node_iterator,
          lo_ixml            TYPE REF TO if_ixml_element,
          lv_dxf_style_index TYPE i,
          lo_excel_style     LIKE LINE OF me->styles.

    FIELD-SYMBOLS: <ls_dxf_style> LIKE LINE OF me->mt_dxf_styles.

    lv_dxf_style_index  = io_ixml_rule->get_attribute_ns( 'dxfId' ).
    READ TABLE me->mt_dxf_styles ASSIGNING <ls_dxf_style> WITH KEY dxf = lv_dxf_style_index.
    IF sy-subrc = 0.
      io_style_cond->mode_expression-cell_style = <ls_dxf_style>-guid.
    ENDIF.

    lo_ixml_nodes ?= io_ixml_rule->get_elements_by_tag_name_ns( name = 'formula' uri = namespace-main ).
    lo_ixml_iterator = lo_ixml_nodes->create_iterator( ).
    lo_ixml ?= lo_ixml_iterator->get_next( ).
    WHILE lo_ixml IS BOUND.

      CASE sy-index.
        WHEN 1.
          io_style_cond->mode_expression-formula  = lo_ixml->get_value( ).


        WHEN OTHERS.
          EXIT.
      ENDCASE.

      lo_ixml ?= lo_ixml_iterator->get_next( ).
    ENDWHILE.


  ENDMETHOD.