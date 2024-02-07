  METHOD zif_excel_style_changer~get_guid.

    result = excel->get_static_cellstyle_guid( ip_cstyle_complete  = complete_style
                                               ip_cstylex_complete = complete_stylex  ).

  ENDMETHOD.