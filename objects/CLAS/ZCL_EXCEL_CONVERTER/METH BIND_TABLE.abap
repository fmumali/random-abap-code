  METHOD bind_table.
    DATA: lt_field_catalog TYPE zexcel_t_fieldcatalog,
          ls_field_catalog TYPE zexcel_s_fieldcatalog,
          ls_fcat          TYPE zexcel_s_converter_fcat,
          lo_column        TYPE REF TO zcl_excel_column,
          lv_col_int       TYPE zexcel_cell_column,
          lv_col_alpha     TYPE zexcel_cell_column_alpha,
          ls_settings      TYPE zexcel_s_table_settings,
          lv_line          TYPE i.

    FIELD-SYMBOLS: <fs_tab>         TYPE ANY TABLE.

    ASSIGN wo_data->* TO <fs_tab> .

    ls_settings-table_style      = i_style_table.
    ls_settings-top_left_column  = zcl_excel_common=>convert_column2alpha( ip_column = w_col_int ).
    ls_settings-top_left_row     = w_row_int.
    ls_settings-show_row_stripes = ws_layout-is_stripped.

    DESCRIBE TABLE  wt_fieldcatalog  LINES lv_line.
    lv_line = lv_line + 1 + w_col_int.
    ls_settings-bottom_right_column = zcl_excel_common=>convert_column2alpha( ip_column = lv_line ).

    DESCRIBE TABLE <fs_tab> LINES lv_line.
    ls_settings-bottom_right_row = lv_line + 1 + w_row_int.
    SORT wt_fieldcatalog BY position.
    LOOP AT wt_fieldcatalog INTO ls_fcat.
      MOVE-CORRESPONDING ls_fcat TO ls_field_catalog.
      ls_field_catalog-dynpfld = abap_true.
      INSERT ls_field_catalog INTO TABLE lt_field_catalog.
    ENDLOOP.

    wo_worksheet->bind_table(
      EXPORTING
        ip_table          = <fs_tab>
        it_field_catalog  = lt_field_catalog
        is_table_settings = ls_settings
      IMPORTING
        es_table_settings = ls_settings
           ).
    LOOP AT wt_fieldcatalog INTO ls_fcat.
      lv_col_int = w_col_int + ls_fcat-position - 1.
      lv_col_alpha = zcl_excel_common=>convert_column2alpha( lv_col_int ).
* Freeze panes
      IF ls_fcat-fix_column = abap_true.
        ADD 1 TO r_freeze_col.
      ENDIF.
* Now let's check for optimized
      IF ls_fcat-is_optimized = abap_true.
        lo_column = wo_worksheet->get_column( ip_column = lv_col_alpha ).
        lo_column->set_auto_size( ip_auto_size = abap_true ) .
      ENDIF.
* Now let's check for visible
      IF ls_fcat-is_hidden = abap_true.
        lo_column = wo_worksheet->get_column( ip_column = lv_col_alpha ).
        lo_column->set_visible( ip_visible = abap_false ) .
      ENDIF.
    ENDLOOP.

  ENDMETHOD.