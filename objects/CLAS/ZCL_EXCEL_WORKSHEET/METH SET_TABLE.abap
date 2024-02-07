  METHOD set_table.

    DATA: lo_tabdescr     TYPE REF TO cl_abap_structdescr,
          lr_data         TYPE REF TO data,
          lt_dfies        TYPE ddfields,
          lv_row_int      TYPE zexcel_cell_row,
          lv_column_int   TYPE zexcel_cell_column,
          lv_column_alpha TYPE zexcel_cell_column_alpha,
          lv_cell_value   TYPE zexcel_cell_value.


    FIELD-SYMBOLS: <fs_table_line> TYPE any,
                   <fs_fldval>     TYPE any,
                   <fs_dfies>      TYPE dfies.

    lv_column_int = zcl_excel_common=>convert_column2int( ip_top_left_column ).
    lv_row_int    = ip_top_left_row.

    CREATE DATA lr_data LIKE LINE OF ip_table.

    lo_tabdescr ?= cl_abap_structdescr=>describe_by_data_ref( lr_data ).

    lt_dfies = lo_tabdescr->get_ddic_field_list( ).

* It is better to loop column by column
    LOOP AT lt_dfies ASSIGNING <fs_dfies>.
      lv_column_alpha = zcl_excel_common=>convert_column2alpha( lv_column_int ).

      IF ip_no_header = abap_false.
        " First of all write column header
        lv_cell_value = <fs_dfies>-scrtext_m.
        me->set_cell( ip_column = lv_column_alpha
                      ip_row    = lv_row_int
                      ip_value  = lv_cell_value
                      ip_style  = ip_hdr_style ).
        IF ip_transpose = abap_true.
          ADD 1 TO lv_column_int.
        ELSE.
          ADD 1 TO lv_row_int.
        ENDIF.
      ENDIF.

      LOOP AT ip_table ASSIGNING <fs_table_line>.
        lv_column_alpha = zcl_excel_common=>convert_column2alpha( lv_column_int ).
        ASSIGN COMPONENT <fs_dfies>-fieldname OF STRUCTURE <fs_table_line> TO <fs_fldval>.
        lv_cell_value = <fs_fldval>.
        me->set_cell( ip_column = lv_column_alpha
                      ip_row    = lv_row_int
                      ip_value  = <fs_fldval>   "lv_cell_value
                      ip_style  = ip_body_style ).
        IF ip_transpose = abap_true.
          ADD 1 TO lv_column_int.
        ELSE.
          ADD 1 TO lv_row_int.
        ENDIF.
      ENDLOOP.
      IF ip_transpose = abap_true.
        lv_column_int = zcl_excel_common=>convert_column2int( ip_top_left_column ).
        ADD 1 TO lv_row_int.
      ELSE.
        lv_row_int = ip_top_left_row.
        ADD 1 TO lv_column_int.
      ENDIF.
    ENDLOOP.

  ENDMETHOD.                    "SET_TABLE