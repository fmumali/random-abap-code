  METHOD get_default_excel_time_format.
    DATA: l_timefm TYPE xutimefm.

    IF default_excel_time_format IS NOT INITIAL.
      ep_default_excel_time_format = default_excel_time_format.
      RETURN.
    ENDIF.

* Let's get default
    l_timefm = cl_abap_timefm=>get_environment_timefm( ).
    CASE l_timefm.
      WHEN 0.
*0  24 Hour Format (Example: 12:05:10)
        default_excel_time_format = zcl_excel_style_number_format=>c_format_date_time6.
      WHEN 1.
*1  12 Hour Format (Example: 12:05:10 PM)
        default_excel_time_format = zcl_excel_style_number_format=>c_format_date_time2.
      WHEN 2.
*2  12 Hour Format (Example: 12:05:10 pm) for now all the same. no chnage upper lower
        default_excel_time_format = zcl_excel_style_number_format=>c_format_date_time2.
      WHEN 3.
*3  Hours from 0 to 11 (Example: 00:05:10 PM)  for now all the same. no chnage upper lower
        default_excel_time_format = zcl_excel_style_number_format=>c_format_date_time2.
      WHEN 4.
*4  Hours from 0 to 11 (Example: 00:05:10 pm)  for now all the same. no chnage upper lower
        default_excel_time_format = zcl_excel_style_number_format=>c_format_date_time2.
      WHEN OTHERS.
        " and fallback to fixed format
        default_excel_time_format = zcl_excel_style_number_format=>c_format_date_time6.
    ENDCASE.

    ep_default_excel_time_format = default_excel_time_format.
  ENDMETHOD.                    "GET_DEFAULT_EXCEL_TIME_FORMAT