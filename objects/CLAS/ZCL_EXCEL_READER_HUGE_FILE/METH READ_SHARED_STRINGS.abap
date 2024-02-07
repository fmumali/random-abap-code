  METHOD read_shared_strings.

    DATA lv_value TYPE string.

    WHILE io_reader->node_type NE c_end_of_stream.
      io_reader->next_node( ).
      CASE io_reader->name.
        WHEN 'si'.
          CASE io_reader->node_type .
            WHEN c_element_open .
              CLEAR lv_value .
            WHEN c_element_close .
              APPEND lv_value TO et_shared_strings.
          ENDCASE .
        WHEN 't'.
          CASE io_reader->node_type .
            WHEN c_node_value .
              lv_value = lv_value && io_reader->value .
          ENDCASE .
      ENDCASE .
    ENDWHILE.

  ENDMETHOD.