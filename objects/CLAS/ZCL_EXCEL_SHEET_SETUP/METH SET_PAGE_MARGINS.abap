  METHOD set_page_margins.
    DATA: lv_coef TYPE f,
          lv_unit TYPE string.

    lv_unit = ip_unit.
    TRANSLATE lv_unit TO UPPER CASE.

    CASE lv_unit.
      WHEN 'IN'. lv_coef = 1.
      WHEN 'CM'. lv_coef = '0.393700787'.
      WHEN 'MM'. lv_coef = '0.0393700787'.
    ENDCASE.

    IF ip_bottom IS SUPPLIED. margin_bottom = lv_coef * ip_bottom. ENDIF.
    IF ip_footer IS SUPPLIED. margin_footer = lv_coef * ip_footer. ENDIF.
    IF ip_header IS SUPPLIED. margin_header = lv_coef * ip_header. ENDIF.
    IF ip_left IS SUPPLIED.   margin_left   = lv_coef * ip_left. ENDIF.
    IF ip_right IS SUPPLIED.  margin_right  = lv_coef * ip_right. ENDIF.
    IF ip_top IS SUPPLIED.    margin_top    = lv_coef * ip_top. ENDIF.

  ENDMETHOD.