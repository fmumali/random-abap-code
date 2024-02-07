  METHOD zif_excel_style_changer~set_complete_borders_top.

    lv_xborder_supplied = boolc( ip_xborders_top IS SUPPLIED ).
    move_supplied_borders(
      EXPORTING
        iv_border_supplied        = lv_border_supplied
        is_border                 = ip_borders_top
        iv_xborder_supplied       = lv_xborder_supplied
        is_xborder                = ip_xborders_top
      CHANGING
        cs_complete_style_border  = complete_style-borders-top
        cs_complete_stylex_border = complete_stylex-borders-top ).
    multiple_change_requested-borders-top = abap_true.

    result = me.

  ENDMETHOD.