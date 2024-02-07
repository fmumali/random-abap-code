  METHOD set_color.
    elements->color_scheme->set_color(
      EXPORTING
        iv_type         = iv_type
        iv_srgb         = iv_srgb
        iv_syscolorname = iv_syscolorname
        iv_syscolorlast = iv_syscolorlast
    ).
  ENDMETHOD.                    "set_color