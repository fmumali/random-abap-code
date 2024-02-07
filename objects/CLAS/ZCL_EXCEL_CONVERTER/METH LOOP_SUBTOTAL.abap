  METHOD loop_subtotal.

    DATA: l_row_int_start   TYPE zexcel_cell_row,
          l_row_int_end     TYPE zexcel_cell_row,
          l_row_int         TYPE zexcel_cell_row,
          l_col_int         TYPE zexcel_cell_column,
          l_col_alpha       TYPE zexcel_cell_column_alpha,
          l_col_alpha_start TYPE zexcel_cell_column_alpha,
          l_cell_value      TYPE zexcel_cell_value,
          l_s_color         TYPE abap_bool,
          lo_column         TYPE REF TO zcl_excel_column,
          lo_row            TYPE REF TO zcl_excel_row,
          l_formula         TYPE zexcel_cell_formula,
          l_style           TYPE zexcel_cell_style,
          l_text            TYPE string,
          ls_sort_values    TYPE ts_sort_values,
          ls_subtotal_rows  TYPE ts_subtotal_rows,
          l_sort_level      TYPE int4,
          l_hidden          TYPE int4,
          l_line            TYPE i,
          l_cells           TYPE i,
          l_count           TYPE i,
          l_table_row       TYPE i,
          lt_fcat           TYPE zexcel_t_converter_fcat.

    FIELD-SYMBOLS: <fs_stab>    TYPE any,
                   <fs_tab>     TYPE STANDARD TABLE,
                   <fs_sfcat>   TYPE zexcel_s_converter_fcat,
                   <fs_fldval>  TYPE any,
                   <fs_sortval> TYPE any,
                   <fs_sortv>   TYPE ts_sort_values.

    ASSIGN wo_data->* TO <fs_tab> .

    CLEAR: wt_sort_values,
           wt_subtotal_rows.

    DESCRIBE TABLE wt_fieldcatalog LINES l_cells.
    DESCRIBE TABLE <fs_tab> LINES l_count.
    l_cells = l_cells * l_count.

    READ TABLE <fs_tab> ASSIGNING <fs_stab> INDEX 1.
    IF sy-subrc = 0.
      l_row_int = i_row_int + 1.
      lt_fcat =  wt_fieldcatalog.
      SORT lt_fcat BY sort_level DESCENDING.
      LOOP AT lt_fcat ASSIGNING <fs_sfcat> WHERE is_subtotalled = abap_true.
        ASSIGN COMPONENT <fs_sfcat>-columnname OF STRUCTURE <fs_stab> TO <fs_fldval>.
        ls_sort_values-fieldname    = <fs_sfcat>-columnname.
        ls_sort_values-row_int      = l_row_int.
        ls_sort_values-sort_level   = <fs_sfcat>-sort_level.
        ls_sort_values-is_collapsed = <fs_sfcat>-is_collapsed.
        CREATE DATA ls_sort_values-value LIKE <fs_fldval>.
        ASSIGN ls_sort_values-value->* TO <fs_sortval>.
        <fs_sortval> =  <fs_fldval>.
        INSERT ls_sort_values INTO TABLE wt_sort_values.
      ENDLOOP.
    ENDIF.
    l_row_int = i_row_int.
* Let's check if we need to hide a sort level.
    DESCRIBE TABLE wt_sort_values LINES l_line.
    IF  l_line <= 1.
      CLEAR l_hidden.
    ELSE.
      LOOP AT wt_sort_values INTO ls_sort_values WHERE is_collapsed = abap_false.
        IF l_hidden < ls_sort_values-sort_level.
          l_hidden = ls_sort_values-sort_level.
        ENDIF.
      ENDLOOP.
    ENDIF.
    ADD 1 TO l_hidden. " As this is the first level we show.
* First loop without formular only addtional rows with subtotal text.
    LOOP AT <fs_tab> ASSIGNING <fs_stab>.
      ADD 1 TO l_row_int.  " 1 is for header row.
      l_row_int_start = l_row_int.
      SORT lt_fcat BY sort_level DESCENDING.
      LOOP AT lt_fcat ASSIGNING <fs_sfcat> WHERE is_subtotalled = abap_true.
        l_col_int   = i_col_int + <fs_sfcat>-position - 1.
        l_col_alpha = zcl_excel_common=>convert_column2alpha( l_col_int ).
* Now the cell values
        ASSIGN COMPONENT <fs_sfcat>-columnname OF STRUCTURE <fs_stab> TO <fs_fldval>.
        IF sy-subrc = 0.
          READ TABLE wt_sort_values ASSIGNING <fs_sortv> WITH TABLE KEY fieldname = <fs_sfcat>-columnname.
          IF sy-subrc = 0.
            ASSIGN <fs_sortv>-value->* TO <fs_sortval>.
            IF <fs_sortval> <> <fs_fldval> OR <fs_sortv>-new = abap_true.
* First let's remmember the subtotal values as it has to appear later.
              ls_subtotal_rows-row_int       = l_row_int.
              ls_subtotal_rows-row_int_start = <fs_sortv>-row_int.
              ls_subtotal_rows-columnname    = <fs_sfcat>-columnname.
              INSERT ls_subtotal_rows INTO TABLE wt_subtotal_rows.
* Now let's write the subtotal line
              l_cell_value = create_text_subtotal( i_value = <fs_sortval>
                                     i_totals_function     = <fs_sfcat>-totals_function ).
              wo_worksheet->set_cell( ip_column    = l_col_alpha
                                      ip_row       = l_row_int
                                      ip_value     = l_cell_value
                                      ip_abap_type = cl_abap_typedescr=>typekind_string
                                      ip_style     = <fs_sfcat>-style_subtotal  ).
              lo_row = wo_worksheet->get_row( ip_row = l_row_int ).
              lo_row->set_outline_level( ip_outline_level = <fs_sfcat>-sort_level ) .
              IF <fs_sfcat>-is_collapsed = abap_true.
                IF <fs_sfcat>-sort_level >  l_hidden.
                  lo_row->set_visible( ip_visible =  abap_false ) .
                ENDIF.
                lo_row->set_collapsed( ip_collapsed =  <fs_sfcat>-is_collapsed ) .
              ENDIF.
* Now let's change the key
              ADD 1 TO l_row_int.
              <fs_sortval> =  <fs_fldval>.
              <fs_sortv>-new = abap_false.
              l_line = <fs_sortv>-sort_level.
              LOOP AT wt_sort_values ASSIGNING <fs_sortv> WHERE sort_level >= l_line.
                <fs_sortv>-row_int = l_row_int.
              ENDLOOP.
            ENDIF.
          ENDIF.
        ENDIF.
      ENDLOOP.
    ENDLOOP.
    ADD 1 TO l_row_int.
    l_row_int_start = l_row_int.
    SORT lt_fcat BY sort_level DESCENDING.
    LOOP AT lt_fcat ASSIGNING <fs_sfcat> WHERE is_subtotalled = abap_true.
      l_col_int   = i_col_int + <fs_sfcat>-position - 1.
      l_col_alpha = zcl_excel_common=>convert_column2alpha( l_col_int ).
      READ TABLE wt_sort_values ASSIGNING <fs_sortv> WITH TABLE KEY fieldname = <fs_sfcat>-columnname.
      IF sy-subrc = 0.
        ASSIGN <fs_sortv>-value->* TO <fs_sortval>.
        ls_subtotal_rows-row_int       = l_row_int.
        ls_subtotal_rows-row_int_start = <fs_sortv>-row_int.
        ls_subtotal_rows-columnname    = <fs_sfcat>-columnname.
        INSERT ls_subtotal_rows INTO TABLE wt_subtotal_rows.
* First let's write the value as it has to appear.
        l_cell_value = create_text_subtotal( i_value = <fs_sortval>
                               i_totals_function     = <fs_sfcat>-totals_function ).
        wo_worksheet->set_cell( ip_column    = l_col_alpha
                                ip_row       = l_row_int
                                ip_value     = l_cell_value
                                ip_abap_type = cl_abap_typedescr=>typekind_string
                                ip_style     = <fs_sfcat>-style_subtotal ).

        l_sort_level = <fs_sfcat>-sort_level.
        lo_row = wo_worksheet->get_row( ip_row = l_row_int ).
        lo_row->set_outline_level( ip_outline_level = l_sort_level ) .
        IF <fs_sfcat>-is_collapsed = abap_true.
          IF <fs_sfcat>-sort_level >  l_hidden.
            lo_row->set_visible( ip_visible =  abap_false ) .
          ENDIF.
          lo_row->set_collapsed( ip_collapsed =  <fs_sfcat>-is_collapsed ) .
        ENDIF.
        ADD 1 TO l_row_int.
      ENDIF.
    ENDLOOP.
* Let's write the Grand total
    l_sort_level = 0.
    lo_row = wo_worksheet->get_row( ip_row = l_row_int ).
    lo_row->set_outline_level( ip_outline_level = l_sort_level ) .
*  lo_row_dim->set_collapsed( ip_collapsed =  <fs_sfcat>-is_collapsed ) . Not on grand total

    l_text    = create_text_subtotal( i_value = 'Grand'(002)
                                      i_totals_function     = <fs_sfcat>-totals_function ).

    l_col_alpha_start = zcl_excel_common=>convert_column2alpha( i_col_int ).
    wo_worksheet->set_cell( ip_column    = l_col_alpha_start
                            ip_row       = l_row_int
                            ip_value     = l_text
                            ip_abap_type = cl_abap_typedescr=>typekind_string
                            ip_style     = <fs_sfcat>-style_subtotal ).

* It is better to loop column by column second time around
* Second loop with formular and data.
    LOOP AT wt_fieldcatalog ASSIGNING <fs_sfcat>.
      l_row_int = i_row_int.
      l_col_int = i_col_int + <fs_sfcat>-position - 1.
* Freeze panes
      IF <fs_sfcat>-fix_column = abap_true.
        ADD 1 TO r_freeze_col.
      ENDIF.
      l_s_color = abap_true.
      l_col_alpha = zcl_excel_common=>convert_column2alpha( l_col_int ).
      " First of all write column header
      l_cell_value = <fs_sfcat>-scrtext_m.
      wo_worksheet->set_cell( ip_column    = l_col_alpha
                              ip_row       = l_row_int
                              ip_value     = l_cell_value
                              ip_abap_type = cl_abap_typedescr=>typekind_string
                              ip_style     = <fs_sfcat>-style_hdr ).
      ADD 1 TO l_row_int.
      LOOP AT <fs_tab> ASSIGNING <fs_stab>.
        l_table_row = sy-tabix.
* Now the cell values
        ASSIGN COMPONENT <fs_sfcat>-columnname OF STRUCTURE <fs_stab> TO <fs_fldval>.
* Let's check for subtotal lines
        DO.
          READ TABLE wt_subtotal_rows TRANSPORTING NO FIELDS WITH TABLE KEY row_int = l_row_int.
          IF sy-subrc = 0.
            IF <fs_sfcat>-is_subtotalled = abap_false AND
               <fs_sfcat>-totals_function IS NOT INITIAL.
              DO.
                READ TABLE wt_subtotal_rows INTO ls_subtotal_rows WITH TABLE KEY row_int    = l_row_int.
                IF sy-subrc = 0.
                  l_row_int_start = ls_subtotal_rows-row_int_start.
                  l_row_int_end   = l_row_int - 1.

                  l_formula = create_formular_subtotal( i_row_int_start   = l_row_int_start
                                                        i_row_int_end     = l_row_int_end
                                                        i_column          = l_col_alpha
                                                        i_totals_function = <fs_sfcat>-totals_function ).
                  wo_worksheet->set_cell( ip_column    = l_col_alpha
                                          ip_row       = l_row_int
                                          ip_formula   = l_formula
                                          ip_style     = <fs_sfcat>-style_subtotal ).
                  IF <fs_sfcat>-is_collapsed = abap_true.
                    lo_row = wo_worksheet->get_row( ip_row = l_row_int ).
                    lo_row->set_collapsed( ip_collapsed =  <fs_sfcat>-is_collapsed ).
                    IF <fs_sfcat>-sort_level >  l_hidden.
                      lo_row->set_visible( ip_visible =  abap_false ) .
                    ENDIF.
                  ENDIF.
                  ADD 1 TO l_row_int.
                ELSE.
                  EXIT.
                ENDIF.
              ENDDO.
            ELSE.
              ADD 1 TO l_row_int.
            ENDIF.
          ELSE.
            EXIT.
          ENDIF.
        ENDDO.
* Let's set the row dimension values
        lo_row = wo_worksheet->get_row( ip_row = l_row_int ).
        lo_row->set_outline_level( ip_outline_level = ws_layout-max_subtotal_level ) .
        IF <fs_sfcat>-is_collapsed  = abap_true.
          lo_row->set_visible( ip_visible =  abap_false ) .
          lo_row->set_collapsed( ip_collapsed =  <fs_sfcat>-is_collapsed ) .
        ENDIF.
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
                                  ip_conv_exit_length = ws_option-conv_exit_length   ).
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
* Let's check for subtotal lines
      DO.
        READ TABLE wt_subtotal_rows TRANSPORTING NO FIELDS WITH TABLE KEY row_int = l_row_int.
        IF sy-subrc = 0.
          IF <fs_sfcat>-is_subtotalled = abap_false AND
             <fs_sfcat>-totals_function IS NOT INITIAL.
            DO.
              READ TABLE wt_subtotal_rows INTO ls_subtotal_rows WITH TABLE KEY row_int    = l_row_int.
              IF sy-subrc = 0.
                l_row_int_start = ls_subtotal_rows-row_int_start.
                l_row_int_end   = l_row_int - 1.

                l_formula = create_formular_subtotal( i_row_int_start   = l_row_int_start
                                                      i_row_int_end     = l_row_int_end
                                                      i_column          = l_col_alpha
                                                      i_totals_function = <fs_sfcat>-totals_function ).
                wo_worksheet->set_cell( ip_column    = l_col_alpha
                                        ip_row       = l_row_int
                                        ip_formula   = l_formula
                                        ip_style     = <fs_sfcat>-style_subtotal ).
                IF <fs_sfcat>-is_collapsed = abap_true.
                  lo_row = wo_worksheet->get_row( ip_row = l_row_int ).
                  lo_row->set_collapsed( ip_collapsed =  <fs_sfcat>-is_collapsed ).
                ENDIF.
                ADD 1 TO l_row_int.
              ELSE.
                EXIT.
              ENDIF.
            ENDDO.
          ELSE.
            ADD 1 TO l_row_int.
          ENDIF.
        ELSE.
          EXIT.
        ENDIF.
      ENDDO.
* Now let's check for Grand total
      IF <fs_sfcat>-is_subtotalled = abap_false AND
         <fs_sfcat>-totals_function IS NOT INITIAL.
        l_row_int_start = i_row_int + 1.
        l_row_int_end   = l_row_int - 1.

        l_formula = create_formular_subtotal( i_row_int_start   = l_row_int_start
                                              i_row_int_end     = l_row_int_end
                                              i_column          = l_col_alpha
                                              i_totals_function = <fs_sfcat>-totals_function ).
        wo_worksheet->set_cell( ip_column    = l_col_alpha
                                ip_row       = l_row_int
                                ip_formula   = l_formula
                                ip_style     = <fs_sfcat>-style_subtotal ).
      ENDIF.
* Now let's check for optimized
      IF <fs_sfcat>-is_optimized = abap_true.
        lo_column = wo_worksheet->get_column( ip_column = l_col_alpha ).
        lo_column->set_auto_size( ip_auto_size = abap_true ) .
      ENDIF.
* Now let's check for visible
      IF <fs_sfcat>-is_hidden = abap_true.
        lo_column = wo_worksheet->get_column( ip_column = l_col_alpha ).
        lo_column->set_visible( ip_visible = abap_false ) .
      ENDIF.
    ENDLOOP.

  ENDMETHOD.