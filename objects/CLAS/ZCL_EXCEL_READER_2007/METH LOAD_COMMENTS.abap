  METHOD load_comments.
    DATA: lo_comments_xml       TYPE REF TO if_ixml_document,
          lo_node_comment       TYPE REF TO if_ixml_element,
          lo_node_comment_child TYPE REF TO if_ixml_element,
          lo_node_r_child_t     TYPE REF TO if_ixml_element,
          lo_attr               TYPE REF TO if_ixml_attribute,
          lo_comment            TYPE REF TO zcl_excel_comment,
          lv_comment_text       TYPE string,
          lv_node_value         TYPE string,
          lv_attr_value         TYPE string.

    lo_comments_xml = me->get_ixml_from_zip_archive( ip_path ).

    lo_node_comment ?= lo_comments_xml->find_from_name_ns( name = 'comment' uri = namespace-main ).
    WHILE lo_node_comment IS BOUND.

      CLEAR lv_comment_text.
      lo_attr = lo_node_comment->get_attribute_node_ns( name = 'ref' ).
      lv_attr_value  = lo_attr->get_value( ).

      lo_node_comment_child ?= lo_node_comment->get_first_child( ).
      WHILE lo_node_comment_child IS BOUND.
        " There will be rPr nodes here, but we do not support them
        " in comments right now; see 'load_shared_strings' for handling.
        " Extract the <t>...</t> part of each <r>-tag
        lo_node_r_child_t ?= lo_node_comment_child->find_from_name_ns( name = 't' uri = namespace-main ).
        IF lo_node_r_child_t IS BOUND.
          lv_node_value = lo_node_r_child_t->get_value( ).
          CONCATENATE lv_comment_text lv_node_value INTO lv_comment_text RESPECTING BLANKS.
        ENDIF.
        lo_node_comment_child ?= lo_node_comment_child->get_next( ).
      ENDWHILE.

      CREATE OBJECT lo_comment.
      lo_comment->set_text( ip_ref = lv_attr_value ip_text = lv_comment_text ).
      io_worksheet->add_comment( lo_comment ).

      lo_node_comment ?= lo_node_comment->get_next( ).
    ENDWHILE.

  ENDMETHOD.