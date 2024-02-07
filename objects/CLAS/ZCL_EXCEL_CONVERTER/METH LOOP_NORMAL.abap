  METHOD loop_normal.
    DATA: l_row_int_end TYPE zexcel_cell_row,
          l_row_int     TYPE zexcel_cell_row,
          l_col_int     TYPE zexcel_cell_column,
          l_col_alpha   TYPE zexcel_cell_column_alpha,
          l_cell_value  TYPE zexcel_cell_value,
          l_s_color     TYPE abap_bool,
          lo_column     TYPE REF TO zcl_excel_column,
          l_formula     TYPE zexcel_cell_formula,
          l_style       TYPE zexcel_cell_style,
          l_cells       TYPE i,
          l_count       TYPE i,
          l_table_row   TYPE i.

    FIELD-SYMBOLS: <fs_stab>   TYPE any,
                   <fs_tab>    TYPE STANDARD TABLE,
                   <fs_sfcat>  TYPE zexcel_s_converter_fcat,
                   <fs_fldval> TYPE any.

    ASSIGN wo_data->* TO <fs_tab> .

    DESCRIBE TABLE wt_fieldcatalog LINES l_cells.
    DESCRIBE TABLE <fs_tab> LINES l_count.
    l_cells = l_cells * l_count.

* It is better to loop column by column
    LOOP AT wt_fieldcatalog ASSIGNING <fs_sfcat>.
      l_row_int = i_row_int.
      l_col_int = i_col_int + <fs_sfcat>-position - 1.

*   Freeze panes
      IF <fs_sfcat>-fix_column = abap_true.
        ADD 1 TO r_freeze_col.
      ENDIF.
      l_s_color = abap_true.

      l_col_alpha = zcl_excel_common=>convert_column2alpha( l_col_int ).

*   Only if the Header is required create it.
      IF ws_option-hidehd IS INITIAL.
        " First of all write column header
        l_cell_value = <fs_sfcat>-scrtext_m.
        wo_worksheet->set_cell( ip_column    = l_col_alpha
                                ip_row       = l_row_int
                                ip_value     = l_cell_value
                                ip_style     = <fs_sfcat>-style_hdr ).
        ADD 1 TO l_row_int.
      ENDIF.
      LOOP AT <fs_tab> ASSIGNING <fs_stab>.
        l_table_row = sy-tabix.
* Now the cell values
        ASSIGN COMPONENT <fs_sfcat>-columnname OF STRUCTURE <fs_stab> TO <fs_fldval>.
* Now let's write the cell values
        IF ws_layout-is_stripped = abap_true AND l_s_color = abap_true.
          l_style = get_color_style( i_row       = l_table_row
                                     i_fieldname = <fs_sfcat>-columnname
                                     i_style     = <fs_sfcat>-style_stripped  ).
          wo_worksheet->set_cell( ip_column    = l_col_alpha
                                  ip_row       = l_row_int
                                  ip_value     = <fs_fldval>
                                  ip_style     = l_style
                                  ip_conv_exit_length = ws_option-conv_exit_length ).
          CLEAR l_s_color.
        ELSE.
          l_style = get_color_style( i_row       = l_table_row
                                     i_fieldname = <fs_sfcat>-columnname
                                     i_style     = <fs_sfcat>-style_normal  ).
          wo_worksheet->set_cell( ip_column    = l_col_alpha
                                  ip_row       = l_row_int
                                  ip_value     = <fs_fldval>
                                  ip_style     = l_style
                                  ip_conv_exit_length = ws_option-conv_exit_length  ).
          l_s_color = abap_true.
        ENDIF.
        READ TABLE wt_filter TRANSPORTING NO FIELDS WITH TABLE KEY rownumber  = l_table_row
                                                                   columnname = <fs_sfcat>-columnname.
        IF sy-subrc = 0.
          wo_worksheet->get_cell( EXPORTING
                                     ip_column    = l_col_alpha
                                     ip_row       = l_row_int
                                  IMPORTING
                                     ep_value     = l_cell_value ).
          wo_autofilter->set_value( i_column = l_col_int
                                    i_value  = l_cell_value ).
        ENDIF.
        ADD 1 TO l_row_int.
      ENDLOOP.
* Now let's check for optimized
      IF <fs_sfcat>-is_optimized = abap_true .
        lo_column = wo_worksheet->get_column( ip_column = l_col_alpha ).
        lo_column->set_auto_size( ip_auto_size = abap_true ) .
      ENDIF.
* Now let's check for visible
      IF <fs_sfcat>-is_hidden = abap_true.
        lo_column = wo_worksheet->get_column( ip_column = l_col_alpha ).
        lo_column->set_visible( ip_visible = abap_false ) .
      ENDIF.
* Now let's check for total versus subtotal.
      IF <fs_sfcat>-totals_function IS NOT INITIAL.
        l_row_int_end = l_row_int - 1.

        l_formula = create_formular_total( i_row_int         = l_row_int_end
                                           i_column          = l_col_alpha
                                           i_totals_function = <fs_sfcat>-totals_function ).
        wo_worksheet->set_cell( ip_column    = l_col_alpha
                                ip_row       = l_row_int
                                ip_formula   = l_formula
                                ip_style     = <fs_sfcat>-style_total ).
      ENDIF.
    ENDLOOP.
  ENDMETHOD.