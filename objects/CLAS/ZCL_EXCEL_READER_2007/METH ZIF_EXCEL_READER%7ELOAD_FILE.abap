  METHOD zif_excel_reader~load_file.

    DATA: lv_excel_data TYPE xstring.

*--------------------------------------------------------------------*
* Read file into binary string
*--------------------------------------------------------------------*
    IF i_from_applserver = abap_true.
      lv_excel_data = read_from_applserver( i_filename ).
    ELSE.
      lv_excel_data = read_from_local_file( i_filename ).
    ENDIF.

*--------------------------------------------------------------------*
* Parse Excel data into ZCL_EXCEL object from binary string
*--------------------------------------------------------------------*
    r_excel = zif_excel_reader~load( i_excel2007            = lv_excel_data
                                     i_use_alternate_zip    = i_use_alternate_zip
                                     iv_zcl_excel_classname = iv_zcl_excel_classname ).

  ENDMETHOD.