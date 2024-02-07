  METHOD set_table_reference.

    FIELD-SYMBOLS: <ls_sheet_content> TYPE zexcel_s_cell_data.

    READ TABLE sheet_content ASSIGNING <ls_sheet_content> WITH KEY cell_row    = ip_row
                                                                   cell_column = ip_column.
    IF sy-subrc = 0.
      <ls_sheet_content>-table           = ir_table.
      <ls_sheet_content>-table_fieldname = ip_fieldname.
      <ls_sheet_content>-table_header    = ip_header.
    ELSE.
      zcx_excel=>raise_text( 'Cell not found' ).
    ENDIF.

  ENDMETHOD.