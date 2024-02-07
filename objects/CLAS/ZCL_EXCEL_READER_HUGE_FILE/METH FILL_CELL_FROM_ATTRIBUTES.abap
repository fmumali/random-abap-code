  METHOD fill_cell_from_attributes.

    WHILE io_reader->node_type NE c_end_of_stream.
      io_reader->next_attribute( ).
      IF io_reader->node_type NE c_attribute.
        EXIT.
      ENDIF.
      CASE io_reader->name.
        WHEN `t`.
          es_cell-datatype = io_reader->value.
        WHEN `s`.
          IF io_reader->value IS NOT INITIAL.
            es_cell-style = get_style( io_reader->value ).
          ENDIF.
        WHEN `r`.
          es_cell-coord = get_cell_coord( io_reader->value ).
      ENDCASE.
    ENDWHILE.

  ENDMETHOD.