  METHOD zif_excel_style_changer~set_complete_borders_left.

    lv_xborder_supplied = boolc( ip_xborders_left IS SUPPLIED ).
    move_supplied_borders(
      EXPORTING
        iv_border_supplied        = lv_border_supplied
        is_border                 = ip_borders_left
        iv_xborder_supplied       = lv_xborder_supplied
        is_xborder                = ip_xborders_left
      CHANGING
        cs_complete_style_border  = complete_style-borders-left
        cs_complete_stylex_border = complete_stylex-borders-left ).
    multiple_change_requested-borders-left = abap_true.

    result = me.

  ENDMETHOD.