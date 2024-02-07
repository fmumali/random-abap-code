  METHOD move_supplied_borders.

    DATA: cs_borderx TYPE zexcel_s_cstylex_border.

    IF iv_border_supplied = abap_true.  " only act if parameter was supplied
      IF iv_xborder_supplied = abap_true. "
        cs_borderx = is_xborder.             " use supplied x-parameter
      ELSE.
        CLEAR cs_borderx WITH 'X'. " <============================== DDIC structure enh. category to set?
        " clear in a way that would be expected to work easily
        IF is_border-border_style IS  INITIAL.
          CLEAR cs_borderx-border_style.
        ENDIF.
        clear_initial_colorxfields(
          EXPORTING
            is_color  = is_border-border_color
          CHANGING
            cs_xcolor = cs_borderx-border_color ).
      ENDIF.
      cs_complete_style_border = is_border.
      cs_complete_stylex_border = cs_borderx.
    ENDIF.

  ENDMETHOD.