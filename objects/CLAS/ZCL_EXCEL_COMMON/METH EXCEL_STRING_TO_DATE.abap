  METHOD excel_string_to_date.
    DATA: lv_date_int TYPE i.

    CHECK ip_value IS NOT INITIAL AND ip_value CN ' 0'.

    TRY.
        lv_date_int = ip_value.
        IF lv_date_int NOT BETWEEN 1 AND 2958465.
          zcx_excel=>raise_text( 'Unable to interpret date' ).
        ENDIF.
        ep_value = lv_date_int + c_excel_baseline_date - 2.
        " Needed hack caused by the problem that:
        " Excel 2000 incorrectly assumes that the year 1900 is a leap year
        " http://support.microsoft.com/kb/214326/en-us
        IF ep_value < c_excel_1900_leap_year.
          ep_value = ep_value + 1.
        ENDIF.
      CATCH cx_sy_conversion_error.
        zcx_excel=>raise_text( 'Index out of bounds' ).
    ENDTRY.
  ENDMETHOD.