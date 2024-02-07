  METHOD add_1_val_child_node.

    DATA: lo_child TYPE REF TO if_ixml_element.

    lo_child = io_document->create_simple_element( name   = iv_elem_name
                                                   parent = io_document ).
    IF iv_attr_name IS NOT INITIAL.
      lo_child->set_attribute_ns( name  = iv_attr_name
                                  value = iv_attr_value ).
    ENDIF.
    io_parent->append_child( new_child = lo_child ).

  ENDMETHOD.