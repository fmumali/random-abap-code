  METHOD update_catalog.
    DATA: ls_fieldcatalog TYPE zexcel_s_converter_fcat,
          ls_fcat         TYPE lvc_s_fcat,
          ls_sort         TYPE lvc_s_sort,
          l_decimals      TYPE lvc_decmls.

    FIELD-SYMBOLS: <fs_scat> TYPE zexcel_s_converter_fcat.

    IF ws_layo-zebra IS NOT INITIAL.
      cs_layout-is_stripped = abap_true.
    ENDIF.
    IF ws_layo-no_keyfix IS INITIAL OR
       ws_layo-no_keyfix = '0'.
      cs_layout-is_fixed = abap_true.
    ENDIF.

    LOOP AT wt_fcat INTO ls_fcat.
      CLEAR: ls_fieldcatalog,
             l_decimals.
      CASE ws_option-hidenc.
        WHEN abap_false.     " We make hiden columns visible
          CLEAR ls_fcat-no_out.
        WHEN abap_true.
* We convert column and hide it.
        WHEN abap_undefined. "We don't convert hiden columns
          IF ls_fcat-no_out = abap_true.
            ls_fcat-tech = abap_true.
          ENDIF.
      ENDCASE.
      IF ls_fcat-tech = abap_false.
        ls_fieldcatalog-tabname         = ls_fcat-tabname.
        ls_fieldcatalog-fieldname       = ls_fcat-fieldname .
        ls_fieldcatalog-columnname      = ls_fcat-fieldname .
        ls_fieldcatalog-position        = ls_fcat-col_pos.
        ls_fieldcatalog-col_id          = ls_fcat-col_id.
        ls_fieldcatalog-convexit        = ls_fcat-convexit.
        ls_fieldcatalog-inttype         = ls_fcat-inttype.
        ls_fieldcatalog-scrtext_s       = ls_fcat-scrtext_s .
        ls_fieldcatalog-scrtext_m       = ls_fcat-scrtext_m .
        ls_fieldcatalog-scrtext_l       = ls_fcat-scrtext_l.
        l_decimals = ls_fcat-decimals_o.
        IF l_decimals IS NOT INITIAL.
          ls_fieldcatalog-decimals = l_decimals.
        ELSE.
          ls_fieldcatalog-decimals = ls_fcat-decimals .
        ENDIF.
        CASE ws_option-subtot.
          WHEN abap_false.     " We ignore subtotals
            CLEAR ls_fcat-do_sum.
          WHEN abap_true.      " We convert subtotals and detail

          WHEN abap_undefined. " We should only take subtotals and displayed detail
* for now abap_true
        ENDCASE.
        CASE ls_fcat-do_sum.
          WHEN abap_true.
            ls_fieldcatalog-totals_function =  zcl_excel_table=>totals_function_sum.
          WHEN 'A'.
            ls_fieldcatalog-totals_function =  zcl_excel_table=>totals_function_min.
          WHEN 'B' .
            ls_fieldcatalog-totals_function =  zcl_excel_table=>totals_function_max.
          WHEN 'C' .
            ls_fieldcatalog-totals_function =  zcl_excel_table=>totals_function_average.
          WHEN OTHERS.
            CLEAR ls_fieldcatalog-totals_function .
        ENDCASE.
        ls_fieldcatalog-fix_column       = ls_fcat-fix_column.
        IF ws_layo-cwidth_opt IS INITIAL.
          IF ls_fcat-col_opt IS NOT INITIAL.
            ls_fieldcatalog-is_optimized = abap_true.
          ENDIF.
        ELSE.
          ls_fieldcatalog-is_optimized = abap_true.
        ENDIF.
        IF ls_fcat-no_out IS NOT INITIAL.
          ls_fieldcatalog-is_hidden = abap_true.
          ls_fieldcatalog-position  = ls_fieldcatalog-col_id. " We hide based on orginal data structure
        ENDIF.
* Alignment in each cell
        CASE ls_fcat-just.
          WHEN 'R'.
            ls_fieldcatalog-alignment = zcl_excel_style_alignment=>c_horizontal_right.
          WHEN 'L'.
            ls_fieldcatalog-alignment = zcl_excel_style_alignment=>c_horizontal_left.
          WHEN 'C'.
            ls_fieldcatalog-alignment = zcl_excel_style_alignment=>c_horizontal_center.
          WHEN OTHERS.
            CLEAR ls_fieldcatalog-alignment.
        ENDCASE.
* Check for subtotals.
        READ TABLE wt_sort INTO ls_sort WITH KEY fieldname = ls_fcat-fieldname.
        IF sy-subrc = 0 AND  ws_option-subtot <> abap_false.
          ls_fieldcatalog-sort_level      = 0 .
          ls_fieldcatalog-is_subtotalled  = ls_sort-subtot.
          ls_fieldcatalog-is_collapsed    = ls_sort-expa.
          IF ls_fieldcatalog-is_subtotalled = abap_true.
            ls_fieldcatalog-sort_level      = ls_sort-spos.
            ls_fieldcatalog-totals_function = zcl_excel_table=>totals_function_sum. " we need function for text
          ENDIF.
        ENDIF.
        APPEND ls_fieldcatalog TO ct_fieldcatalog.
      ENDIF.
    ENDLOOP.

    SORT ct_fieldcatalog BY sort_level ASCENDING.
    cs_layout-max_subtotal_level  = 0.
    LOOP AT ct_fieldcatalog ASSIGNING <fs_scat> WHERE sort_level > 0.
      cs_layout-max_subtotal_level = cs_layout-max_subtotal_level + 1.
      <fs_scat>-sort_level = cs_layout-max_subtotal_level.
    ENDLOOP.

  ENDMETHOD.