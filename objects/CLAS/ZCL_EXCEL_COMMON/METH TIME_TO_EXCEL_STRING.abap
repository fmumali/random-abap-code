  METHOD time_to_excel_string.
    DATA: lv_seconds_in_day TYPE i,
          lv_day_fraction   TYPE f,
          lc_time_baseline  TYPE t VALUE '000000',
          lc_seconds_in_day TYPE i VALUE 86400.

    lv_seconds_in_day = ip_value - lc_time_baseline.
    lv_day_fraction = lv_seconds_in_day / lc_seconds_in_day.
    ep_value = zcl_excel_common=>number_to_excel_string( ip_value = lv_day_fraction ).
  ENDMETHOD.