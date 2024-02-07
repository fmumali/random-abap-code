  METHOD load_worksheet_cond_format_is.
    DATA: lo_ixml_nodes        TYPE REF TO if_ixml_node_collection,
          lo_ixml_iterator     TYPE REF TO if_ixml_node_iterator,
          lo_ixml              TYPE REF TO if_ixml_element,
          lo_ixml_rule_iconset TYPE REF TO if_ixml_element.

    lo_ixml_rule_iconset ?= io_ixml_rule->get_first_child( ).
    io_style_cond->mode_iconset-iconset   = lo_ixml_rule_iconset->get_attribute_ns( 'iconSet' ).
    io_style_cond->mode_iconset-showvalue = lo_ixml_rule_iconset->get_attribute_ns( 'showValue' ).
    lo_ixml_nodes ?= lo_ixml_rule_iconset->get_elements_by_tag_name_ns( name = 'cfvo' uri = namespace-main ).
    lo_ixml_iterator = lo_ixml_nodes->create_iterator( ).
    lo_ixml ?= lo_ixml_iterator->get_next( ).
    WHILE lo_ixml IS BOUND.

      CASE sy-index.
        WHEN 1.
          io_style_cond->mode_iconset-cfvo1_type  = lo_ixml->get_attribute_ns( 'type' ).
          io_style_cond->mode_iconset-cfvo1_value = lo_ixml->get_attribute_ns( 'val' ).

        WHEN 2.
          io_style_cond->mode_iconset-cfvo2_type  = lo_ixml->get_attribute_ns( 'type' ).
          io_style_cond->mode_iconset-cfvo2_value = lo_ixml->get_attribute_ns( 'val' ).

        WHEN 3.
          io_style_cond->mode_iconset-cfvo3_type  = lo_ixml->get_attribute_ns( 'type' ).
          io_style_cond->mode_iconset-cfvo3_value = lo_ixml->get_attribute_ns( 'val' ).

        WHEN 4.
          io_style_cond->mode_iconset-cfvo4_type  = lo_ixml->get_attribute_ns( 'type' ).
          io_style_cond->mode_iconset-cfvo4_value = lo_ixml->get_attribute_ns( 'val' ).

        WHEN 5.
          io_style_cond->mode_iconset-cfvo5_type  = lo_ixml->get_attribute_ns( 'type' ).
          io_style_cond->mode_iconset-cfvo5_value = lo_ixml->get_attribute_ns( 'val' ).

        WHEN OTHERS.
          EXIT.
      ENDCASE.

      lo_ixml ?= lo_ixml_iterator->get_next( ).
    ENDWHILE.

  ENDMETHOD.