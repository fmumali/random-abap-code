  METHOD zif_excel_style_changer~set_complete_fill.

    DATA: fillx LIKE ip_xfill.
    IF ip_xfill IS SUPPLIED.
      fillx = ip_xfill.
    ELSE.
      CLEAR fillx WITH 'X'.
      IF ip_fill-filltype IS INITIAL.
        CLEAR fillx-filltype.
      ENDIF.
      clear_initial_colorxfields(
        EXPORTING
          is_color  = ip_fill-fgcolor
        CHANGING
          cs_xcolor = fillx-fgcolor ).
      clear_initial_colorxfields(
        EXPORTING
          is_color  = ip_fill-bgcolor
        CHANGING
          cs_xcolor = fillx-bgcolor ).

    ENDIF.

    complete_style-fill = ip_fill.
    complete_stylex-fill = fillx.
    multiple_change_requested-fill = abap_true.

    result = me.

  ENDMETHOD.