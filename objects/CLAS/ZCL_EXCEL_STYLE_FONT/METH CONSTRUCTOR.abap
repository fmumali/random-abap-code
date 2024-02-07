  METHOD constructor.
    me->color-rgb       = zcl_excel_style_color=>c_black.
    me->color-theme     = zcl_excel_style_color=>c_theme_not_set.
    me->color-indexed   = zcl_excel_style_color=>c_indexed_not_set.
    me->scheme          = zcl_excel_style_font=>c_scheme_minor.
    me->underline_mode  = zcl_excel_style_font=>c_underline_single.
  ENDMETHOD.