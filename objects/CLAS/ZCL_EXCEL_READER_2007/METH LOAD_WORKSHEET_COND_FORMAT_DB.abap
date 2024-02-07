  METHOD load_worksheet_cond_format_db.
    DATA: lo_ixml_nodes    TYPE REF TO if_ixml_node_collection,
          lo_ixml_iterator TYPE REF TO if_ixml_node_iterator,
          lo_ixml          TYPE REF TO if_ixml_element.

    lo_ixml ?= io_ixml_rule->find_from_name_ns( name = 'color' uri = namespace-main ).
    IF lo_ixml IS BOUND.
      io_style_cond->mode_databar-colorrgb = lo_ixml->get_attribute_ns( 'rgb' ).
    ENDIF.

    lo_ixml_nodes ?= io_ixml_rule->get_elements_by_tag_name_ns( name = 'cfvo' uri = namespace-main ).
    lo_ixml_iterator = lo_ixml_nodes->create_iterator( ).
    lo_ixml ?= lo_ixml_iterator->get_next( ).
    WHILE lo_ixml IS BOUND.

      CASE sy-index.
        WHEN 1.
          io_style_cond->mode_databar-cfvo1_type  = lo_ixml->get_attribute_ns( 'type' ).
          io_style_cond->mode_databar-cfvo1_value = lo_ixml->get_attribute_ns( 'val' ).

        WHEN 2.
          io_style_cond->mode_databar-cfvo2_type  = lo_ixml->get_attribute_ns( 'type' ).
          io_style_cond->mode_databar-cfvo2_value = lo_ixml->get_attribute_ns( 'val' ).

        WHEN OTHERS.
          EXIT.
      ENDCASE.

      lo_ixml ?= lo_ixml_iterator->get_next( ).
    ENDWHILE.


  ENDMETHOD.