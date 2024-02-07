  METHOD read_worksheet_data.

    DATA: ls_cell   TYPE t_cell.

* Skip to <sheetData> element
    skip_to(  iv_element_name = `sheetData`  io_reader = io_reader ).

* Main loop: Evaluate the <c> elements and its children
    WHILE io_reader->node_type NE c_end_of_stream.
      io_reader->next_node( ).
      CASE io_reader->node_type.
        WHEN c_element_open.
          IF io_reader->name EQ `c`.
            ls_cell = fill_cell_from_attributes( io_reader ).
          ENDIF.
        WHEN c_node_value.
          CASE io_reader->name.
            WHEN `f`.
              ls_cell-formula = io_reader->value.
            WHEN `v`.
              IF ls_cell-datatype EQ `s`.
                ls_cell-value = get_shared_string( io_reader->value ).
              ELSE.
                ls_cell-value = io_reader->value.
              ENDIF.
            WHEN `t` OR `is`.
              ls_cell-value = io_reader->value.
          ENDCASE.
        WHEN c_element_close.
          CASE io_reader->name.
            WHEN `c`.
              put_cell_to_worksheet( is_cell = ls_cell io_worksheet = io_worksheet ).
            WHEN `sheetData`.
              EXIT.
          ENDCASE.
      ENDCASE.
    ENDWHILE.

  ENDMETHOD.