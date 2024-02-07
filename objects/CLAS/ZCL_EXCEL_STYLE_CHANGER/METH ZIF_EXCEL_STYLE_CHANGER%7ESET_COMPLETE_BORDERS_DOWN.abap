  METHOD zif_excel_style_changer~set_complete_borders_down.

    lv_xborder_supplied = boolc( ip_xborders_down IS SUPPLIED ).
    move_supplied_borders(
      EXPORTING
        iv_border_supplied        = lv_border_supplied
        is_border                 = ip_borders_down
        iv_xborder_supplied       = lv_xborder_supplied
        is_xborder                = ip_xborders_down
      CHANGING
        cs_complete_style_border  = complete_style-borders-down
        cs_complete_stylex_border = complete_stylex-borders-down ).
    multiple_change_requested-borders-down = abap_true.

    result = me.

  ENDMETHOD.