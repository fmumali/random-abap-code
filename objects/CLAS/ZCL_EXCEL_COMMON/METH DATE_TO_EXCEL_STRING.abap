  METHOD date_to_excel_string.
    DATA: lv_date_diff         TYPE i.

    CHECK ip_value IS NOT INITIAL
      AND ip_value <> space.
    " Needed hack caused by the problem that:
    " Excel 2000 incorrectly assumes that the year 1900 is a leap year
    " http://support.microsoft.com/kb/214326/en-us
    IF ip_value > c_excel_1900_leap_year.
      lv_date_diff = ip_value - c_excel_baseline_date + 2.
    ELSE.
      lv_date_diff = ip_value - c_excel_baseline_date + 1.
    ENDIF.
    ep_value = zcl_excel_common=>number_to_excel_string( ip_value = lv_date_diff ).
  ENDMETHOD.