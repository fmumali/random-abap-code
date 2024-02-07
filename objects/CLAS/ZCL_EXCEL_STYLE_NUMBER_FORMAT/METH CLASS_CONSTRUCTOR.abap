  METHOD class_constructor.

    CLEAR mt_built_in_num_formats.

    add_format( id = '1' code = zcl_excel_style_number_format=>c_format_number ).               " '0'.
    add_format( id = '2' code = zcl_excel_style_number_format=>c_format_number_00 ).            " '0.00'.
    add_format( id = '3' code = zcl_excel_style_number_format=>c_format_number_comma_sep0 ).    " '#,##0'.
    add_format( id = '4' code = zcl_excel_style_number_format=>c_format_number_comma_sep1 ).    " '#,##0.00'.
    add_format( id = '5' code = zcl_excel_style_number_format=>c_format_currency_simple ).      " '$#,##0_);($#,##0)'.
    add_format( id = '6' code = zcl_excel_style_number_format=>c_format_currency_simple_red ).  " '$#,##0_);[Red]($#,##0)'.
    add_format( id = '7' code = zcl_excel_style_number_format=>c_format_currency_simple2 ).     " '$#,##0.00_);($#,##0.00)'.
    add_format( id = '8' code = zcl_excel_style_number_format=>c_format_currency_simple_red2 ). " '$#,##0.00_);[Red]($#,##0.00)'.
    add_format( id = '9' code = zcl_excel_style_number_format=>c_format_percentage ).           " '0%'.
    add_format( id = '10' code = zcl_excel_style_number_format=>c_format_percentage_00 ).        " '0.00%'.
    add_format( id = '11' code = zcl_excel_style_number_format=>c_format_scientific ).           " '0.00E+00'.
    add_format( id = '12' code = zcl_excel_style_number_format=>c_format_fraction_1 ).           " '# ?/?'.
    add_format( id = '13' code = zcl_excel_style_number_format=>c_format_fraction_2 ).           " '# ??/??'.
    add_format( id = '14' code = zcl_excel_style_number_format=>c_format_date_xlsx14 ).          "'m/d/yyyy'.  <--  should have been 'mm-dd-yy' like constant in zcl_excel_style_number_format
    add_format( id = '15' code = zcl_excel_style_number_format=>c_format_date_xlsx15 ).          "'d-mmm-yy'.
    add_format( id = '16' code = zcl_excel_style_number_format=>c_format_date_xlsx16 ).          "'d-mmm'.
    add_format( id = '17' code = zcl_excel_style_number_format=>c_format_date_xlsx17 ).          "'mmm-yy'.
    add_format( id = '18' code = zcl_excel_style_number_format=>c_format_date_time1 ).           " 'h:mm AM/PM'.
    add_format( id = '19' code = zcl_excel_style_number_format=>c_format_date_time2 ).           " 'h:mm:ss AM/PM'.
    add_format( id = '20' code = zcl_excel_style_number_format=>c_format_date_time3 ).           " 'h:mm'.
    add_format( id = '21' code = zcl_excel_style_number_format=>c_format_date_time4 ).           " 'h:mm:ss'.
    add_format( id = '22' code = zcl_excel_style_number_format=>c_format_date_xlsx22 ).          " 'm/d/yyyy h:mm'.


    add_format( id = '37' code = zcl_excel_style_number_format=>c_format_xlsx37 ).               " '#,##0_);(#,##0)'.
    add_format( id = '38' code = zcl_excel_style_number_format=>c_format_xlsx38 ).               " '#,##0_);[Red](#,##0)'.
    add_format( id = '39' code = zcl_excel_style_number_format=>c_format_xlsx39 ).               " '#,##0.00_);(#,##0.00)'.
    add_format( id = '40' code = zcl_excel_style_number_format=>c_format_xlsx40 ).               " '#,##0.00_);[Red](#,##0.00)'.
    add_format( id = '41' code = zcl_excel_style_number_format=>c_format_xlsx41 ).               " '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)'.
    add_format( id = '42' code = zcl_excel_style_number_format=>c_format_xlsx42 ).               " '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)'.
    add_format( id = '43' code = zcl_excel_style_number_format=>c_format_xlsx43 ).               " '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'.
    add_format( id = '44' code = zcl_excel_style_number_format=>c_format_xlsx44 ).               " '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'.
    add_format( id = '45' code = zcl_excel_style_number_format=>c_format_date_xlsx45 ).          " 'mm:ss'.
    add_format( id = '46' code = zcl_excel_style_number_format=>c_format_date_xlsx46 ).          " '[h]:mm:ss'.
    add_format( id = '47' code = zcl_excel_style_number_format=>c_format_date_xlsx47 ).          "  'mm:ss.0'.
    add_format( id = '48' code = zcl_excel_style_number_format=>c_format_special_01 ).           " '##0.0E+0'.
    add_format( id = '49' code = zcl_excel_style_number_format=>c_format_text ).                 " '@'.

  ENDMETHOD.