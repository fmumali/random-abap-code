  METHOD move_supplied_borders.

    DATA: ls_borderx TYPE zexcel_s_cstylex_border.

    IF iv_border_supplied = abap_true.  " only act if parameter was supplied
      IF iv_xborder_supplied = abap_true. "
        ls_borderx = is_xborder.             " use supplied x-parameter
      ELSE.
        CLEAR ls_borderx WITH 'X'. " <============================== DDIC structure enh. category to set?
        " clear in a way that would be expected to work easily
        IF is_border-border_style IS  INITIAL.
          CLEAR ls_borderx-border_style.
        ENDIF.
        clear_initial_colorxfields(
          EXPORTING
            is_color  = is_border-border_color
          CHANGING
            cs_xcolor = ls_borderx-border_color ).
      ENDIF.
      MOVE-CORRESPONDING is_border  TO cs_complete_style_border.
      MOVE-CORRESPONDING ls_borderx TO cs_complete_stylex_border.
    ENDIF.

  ENDMETHOD.