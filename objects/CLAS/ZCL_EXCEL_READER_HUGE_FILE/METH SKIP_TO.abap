  METHOD skip_to.

    DATA: lv_error TYPE string.

* Skip forward to given element
    WHILE io_reader->name NE iv_element_name OR
          io_reader->node_type NE c_element_open.
      io_reader->next_node( ).
      IF io_reader->node_type = c_end_of_stream.
        CONCATENATE 'XML error: Didn''t find element <' iv_element_name '>' INTO lv_error.
        RAISE EXCEPTION TYPE lcx_not_found
          EXPORTING
            error = lv_error.
      ENDIF.
    ENDWHILE.


  ENDMETHOD.