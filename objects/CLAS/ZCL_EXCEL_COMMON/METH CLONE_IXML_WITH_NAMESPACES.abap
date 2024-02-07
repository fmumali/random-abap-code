  METHOD clone_ixml_with_namespaces.

    DATA: iterator    TYPE REF TO if_ixml_node_iterator,
          node        TYPE REF TO if_ixml_node,
          xmlns       TYPE ihttpnvp,
          xmlns_table TYPE TABLE OF ihttpnvp.
    FIELD-SYMBOLS:
      <xmlns> TYPE ihttpnvp.

    iterator = element->create_iterator( ).
    result ?= element->clone( ).
    node = iterator->get_next( ).
    WHILE node IS BOUND.
      xmlns-name = node->get_namespace_prefix( ).
      xmlns-value = node->get_namespace_uri( ).
      COLLECT xmlns INTO xmlns_table.
      node = iterator->get_next( ).
    ENDWHILE.

    LOOP AT xmlns_table ASSIGNING <xmlns>.
      result->set_attribute_ns( prefix = 'xmlns' name = <xmlns>-name value = <xmlns>-value ).
    ENDLOOP.

  ENDMETHOD.