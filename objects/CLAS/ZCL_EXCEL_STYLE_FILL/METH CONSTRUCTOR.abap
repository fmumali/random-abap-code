  METHOD constructor.
    filltype = zcl_excel_style_fill=>c_fill_none.
    fgcolor-theme     = zcl_excel_style_color=>c_theme_not_set.
    fgcolor-indexed   = zcl_excel_style_color=>c_indexed_not_set.
    bgcolor-theme     = zcl_excel_style_color=>c_theme_not_set.
    bgcolor-indexed   = zcl_excel_style_color=>c_indexed_sys_foreground.
    rotation = 0.

  ENDMETHOD.                    "CONSTRUCTOR