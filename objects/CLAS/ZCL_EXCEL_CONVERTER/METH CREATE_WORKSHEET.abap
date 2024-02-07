  METHOD create_worksheet.
    DATA: l_freeze_col TYPE i.

    IF wo_data IS BOUND AND wo_worksheet IS BOUND.

      wo_worksheet->zif_excel_sheet_properties~summarybelow = zif_excel_sheet_properties=>c_below_on. " By default is on

      IF wt_fieldcatalog IS INITIAL.
        set_fieldcatalog( ) .
      ELSE.
        clean_fieldcatalog( ) .
      ENDIF.

      IF i_table = abap_true.
        l_freeze_col = bind_table( i_style_table = i_style_table ) .
      ELSEIF wt_filter IS NOT INITIAL.
* Let's check for filter.
        wo_autofilter = wo_excel->add_new_autofilter( io_sheet = wo_worksheet ).
        l_freeze_col = bind_cells( ) .
        set_autofilter_area( ) .
      ELSE.
        l_freeze_col = bind_cells( ) .
      ENDIF.

* Check for freeze panes
      IF ws_layout-is_fixed = abap_true.
        IF l_freeze_col = 0.
          l_freeze_col = w_col_int.
        ENDIF.
        wo_worksheet->freeze_panes( EXPORTING ip_num_columns = l_freeze_col
                                              ip_num_rows    = w_row_int ) .
      ENDIF.
    ENDIF.

  ENDMETHOD.