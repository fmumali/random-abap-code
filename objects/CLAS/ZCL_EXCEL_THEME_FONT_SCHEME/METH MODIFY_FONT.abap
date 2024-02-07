  METHOD modify_font.
    DATA: ls_font TYPE t_font.
    FIELD-SYMBOLS: <font> TYPE t_font.
    ls_font-script = iv_script.
    ls_font-typeface = iv_typeface.
    TRY.
        CASE iv_type.
          WHEN c_major.
            READ TABLE font_scheme-major-fonts WITH KEY script = iv_script ASSIGNING <font>.
            IF sy-subrc EQ 0.
              <font> = ls_font.
            ELSE.
              INSERT ls_font INTO TABLE font_scheme-major-fonts.
            ENDIF.
          WHEN c_minor.
            READ TABLE font_scheme-minor-fonts WITH KEY script = iv_script ASSIGNING <font>.
            IF sy-subrc EQ 0.
              <font> = ls_font.
            ELSE.
              INSERT ls_font INTO TABLE font_scheme-minor-fonts.
            ENDIF.
        ENDCASE.
      CATCH cx_root. "not the best but just to avoid duplicate lines dump
    ENDTRY.
  ENDMETHOD.                    "add_font