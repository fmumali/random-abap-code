  METHOD create_new_arbg_int.
    DATA: lv_red        TYPE int1,
          lv_green      TYPE int1,
          lv_blue       TYPE int1,
          lv_hex        TYPE x,
          lv_char_red   TYPE zexcel_style_color_component,
          lv_char_green TYPE zexcel_style_color_component,
          lv_char_blue  TYPE zexcel_style_color_component.

    lv_red    = iv_red MOD 256.
    lv_green  = iv_green MOD 256.
    lv_blue   = iv_blue  MOD 256.

    lv_hex        = lv_red.
    lv_char_red   = lv_hex.

    lv_hex        = lv_green.
    lv_char_green = lv_hex.

    lv_hex        = lv_blue.
    lv_char_blue  = lv_hex.


    CONCATENATE zcl_excel_style_color=>c_alpha lv_char_red lv_char_green lv_char_blue INTO rv_color_argb.


  ENDMETHOD.