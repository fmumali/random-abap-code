  METHOD excel_string_to_time.
    DATA: lv_seconds_in_day TYPE i,
          lv_day_fraction   TYPE f,
          lc_seconds_in_day TYPE i VALUE 86400.

    TRY.

        lv_day_fraction = ip_value.
        lv_seconds_in_day = lv_day_fraction * lc_seconds_in_day.

        ep_value = lv_seconds_in_day.

      CATCH cx_sy_conversion_error.
        zcx_excel=>raise_text( 'Unable to interpret time' ).
    ENDTRY.
  ENDMETHOD.