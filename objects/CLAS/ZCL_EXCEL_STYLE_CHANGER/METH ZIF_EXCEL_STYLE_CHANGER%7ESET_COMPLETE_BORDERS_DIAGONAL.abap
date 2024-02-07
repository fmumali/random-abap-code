  METHOD zif_excel_style_changer~set_complete_borders_diagonal.

    lv_xborder_supplied = boolc( ip_xborders_diagonal IS SUPPLIED ).
    move_supplied_borders(
      EXPORTING
        iv_border_supplied        = lv_border_supplied
        is_border                 = ip_borders_diagonal
        iv_xborder_supplied       = lv_xborder_supplied
        is_xborder                = ip_xborders_diagonal
      CHANGING
        cs_complete_style_border  = complete_style-borders-diagonal
        cs_complete_stylex_border = complete_stylex-borders-diagonal ).
    multiple_change_requested-borders-diagonal = abap_true.

    result = me.

  ENDMETHOD.