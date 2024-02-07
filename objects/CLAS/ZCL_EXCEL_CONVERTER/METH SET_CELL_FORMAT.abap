  METHOD set_cell_format.
    DATA:  l_format    TYPE zexcel_number_format.

    CLEAR r_format.
    CASE i_inttype.
      WHEN cl_abap_typedescr=>typekind_date.
        r_format = wo_worksheet->get_default_excel_date_format( ).
      WHEN cl_abap_typedescr=>typekind_time.
        r_format = wo_worksheet->get_default_excel_time_format( ).
      WHEN cl_abap_typedescr=>typekind_float OR cl_abap_typedescr=>typekind_packed OR
           cl_abap_typedescr=>typekind_decfloat OR
           cl_abap_typedescr=>typekind_decfloat16 OR
           cl_abap_typedescr=>typekind_decfloat34.
        IF i_decimals > 0 .
          l_format = '#,##0.'.
          DO i_decimals TIMES.
            CONCATENATE l_format '0' INTO l_format.
          ENDDO.
          r_format = l_format.
        ENDIF.
      WHEN cl_abap_typedescr=>typekind_int OR cl_abap_typedescr=>typekind_int1 OR cl_abap_typedescr=>typekind_int2.
        r_format = '#,##0'.
    ENDCASE.

  ENDMETHOD.