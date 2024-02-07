  METHOD zif_excel_style_changer~set_complete_borders_right.

    lv_xborder_supplied = boolc( ip_xborders_right IS SUPPLIED ).
    move_supplied_borders(
      EXPORTING
        iv_border_supplied        = lv_border_supplied
        is_border                 = ip_borders_right
        iv_xborder_supplied       = lv_xborder_supplied
        is_xborder                = ip_xborders_right
      CHANGING
        cs_complete_style_border  = complete_style-borders-right
        cs_complete_stylex_border = complete_stylex-borders-right ).
    multiple_change_requested-borders-right = abap_true.

    result = me.

  ENDMETHOD.