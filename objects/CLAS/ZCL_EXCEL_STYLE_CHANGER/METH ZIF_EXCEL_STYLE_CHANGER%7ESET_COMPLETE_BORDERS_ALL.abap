  METHOD zif_excel_style_changer~set_complete_borders_all.

    lv_xborder_supplied = boolc( ip_xborders_allborders IS SUPPLIED ).
    move_supplied_borders(
      EXPORTING
        iv_border_supplied        = lv_border_supplied
        is_border                 = ip_borders_allborders
        iv_xborder_supplied       = lv_xborder_supplied
        is_xborder                = ip_xborders_allborders
      CHANGING
        cs_complete_style_border  = complete_style-borders-allborders
        cs_complete_stylex_border = complete_stylex-borders-allborders ).
    multiple_change_requested-borders-allborders = abap_true.

    result = me.

  ENDMETHOD.