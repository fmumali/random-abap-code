  METHOD create_new_argb.

    CONCATENATE zcl_excel_style_color=>c_alpha ip_red ip_green ip_blu INTO ep_color_argb.

  ENDMETHOD.