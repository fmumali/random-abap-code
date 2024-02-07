  METHOD set_default_excel_date_format.

    IF ip_default_excel_date_format IS INITIAL.
      zcx_excel=>raise_text( 'Default date format cannot be blank' ).
    ENDIF.

    default_excel_date_format = ip_default_excel_date_format.
  ENDMETHOD.                    "SET_DEFAULT_EXCEL_DATE_FORMAT