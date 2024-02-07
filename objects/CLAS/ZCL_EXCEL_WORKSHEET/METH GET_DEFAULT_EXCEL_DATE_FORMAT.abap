  METHOD get_default_excel_date_format.
    CONSTANTS: c_lang_e TYPE lang VALUE 'E'.

    IF default_excel_date_format IS NOT INITIAL.
      ep_default_excel_date_format = default_excel_date_format.
      RETURN.
    ENDIF.

    "try to get defaults
    TRY.
        cl_abap_datfm=>get_date_format_des( EXPORTING im_langu = c_lang_e
                                            IMPORTING ex_dateformat = default_excel_date_format ).
      CATCH cx_abap_datfm_format_unknown.

    ENDTRY.

    " and fallback to fixed format
    IF default_excel_date_format IS INITIAL.
      default_excel_date_format = zcl_excel_style_number_format=>c_format_date_ddmmyyyydot.
    ENDIF.

    ep_default_excel_date_format = default_excel_date_format.
  ENDMETHOD.                    "GET_DEFAULT_EXCEL_DATE_FORMAT