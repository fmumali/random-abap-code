  METHOD pixel2emu.
* suppose 96 DPI
    IF ip_dpi IS SUPPLIED.
      r_emu = ip_pixel  * 914400 / ip_dpi.
    ELSE.
* suppose 96 DPI
      r_emu = ip_pixel  * 914400 / 96.
    ENDIF.
  ENDMETHOD.