  METHOD emu2pixel.
* suppose 96 DPI
    IF ip_dpi IS SUPPLIED.
      r_pixel = ip_emu * ip_dpi / 914400.
    ELSE.
* suppose 96 DPI
      r_pixel = ip_emu * 96 / 914400.
    ENDIF.
  ENDMETHOD.