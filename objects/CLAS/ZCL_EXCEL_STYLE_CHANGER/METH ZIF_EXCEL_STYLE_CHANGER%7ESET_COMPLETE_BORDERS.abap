  METHOD zif_excel_style_changer~set_complete_borders.

    DATA: bordersx LIKE ip_xborders.
    IF ip_xborders IS SUPPLIED.
      bordersx = ip_xborders.
    ELSE.
      CLEAR bordersx WITH 'X'.
      IF ip_borders-allborders-border_style IS INITIAL.
        CLEAR bordersx-allborders-border_style.
      ENDIF.
      IF ip_borders-diagonal-border_style IS INITIAL.
        CLEAR bordersx-diagonal-border_style.
      ENDIF.
      IF ip_borders-down-border_style IS INITIAL.
        CLEAR bordersx-down-border_style.
      ENDIF.
      IF ip_borders-left-border_style IS INITIAL.
        CLEAR bordersx-left-border_style.
      ENDIF.
      IF ip_borders-right-border_style IS INITIAL.
        CLEAR bordersx-right-border_style.
      ENDIF.
      IF ip_borders-top-border_style IS INITIAL.
        CLEAR bordersx-top-border_style.
      ENDIF.

      clear_initial_colorxfields(
        EXPORTING
          is_color  = ip_borders-allborders-border_color
        CHANGING
          cs_xcolor = bordersx-allborders-border_color ).

      clear_initial_colorxfields(
        EXPORTING
          is_color  = ip_borders-diagonal-border_color
        CHANGING
          cs_xcolor = bordersx-diagonal-border_color ).

      clear_initial_colorxfields(
        EXPORTING
          is_color  = ip_borders-down-border_color
        CHANGING
          cs_xcolor = bordersx-down-border_color ).

      clear_initial_colorxfields(
        EXPORTING
          is_color  = ip_borders-left-border_color
        CHANGING
          cs_xcolor = bordersx-left-border_color ).

      clear_initial_colorxfields(
        EXPORTING
          is_color  = ip_borders-right-border_color
        CHANGING
          cs_xcolor = bordersx-right-border_color ).

      clear_initial_colorxfields(
        EXPORTING
          is_color  = ip_borders-top-border_color
        CHANGING
          cs_xcolor = bordersx-top-border_color ).

    ENDIF.

    complete_style-borders = ip_borders.
    complete_stylex-borders = bordersx.
    multiple_change_requested-borders-complete = abap_true.

    result = me.

  ENDMETHOD.