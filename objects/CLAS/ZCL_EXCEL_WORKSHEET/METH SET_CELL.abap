  METHOD set_cell.

    DATA: lv_column        TYPE zexcel_cell_column,
          ls_sheet_content TYPE zexcel_s_cell_data,
          lv_row           TYPE zexcel_cell_row,
          lv_value         TYPE zexcel_cell_value,
          lv_data_type     TYPE zexcel_cell_data_type,
          lv_value_type    TYPE abap_typekind,
          lv_style_guid    TYPE zexcel_cell_style,
          lo_addit         TYPE REF TO cl_abap_elemdescr,
          lt_rtf           TYPE zexcel_t_rtf,
          lo_value         TYPE REF TO data,
          lo_value_new     TYPE REF TO data.

    FIELD-SYMBOLS: <fs_sheet_content> TYPE zexcel_s_cell_data,
                   <fs_numeric>       TYPE numeric,
                   <fs_date>          TYPE d,
                   <fs_time>          TYPE t,
                   <fs_value>         TYPE simple,
                   <fs_typekind_int8> TYPE abap_typekind.
    FIELD-SYMBOLS: <fs_column_formula> TYPE mty_s_column_formula.
    FIELD-SYMBOLS: <ls_fieldcat>       TYPE zexcel_s_fieldcatalog.

    IF ip_value  IS NOT SUPPLIED
        AND ip_formula IS NOT SUPPLIED
        AND ip_column_formula_id = 0.
      zcx_excel=>raise_text( 'Please provide the value or formula' ).
    ENDIF.

    normalize_columnrow_parameter( EXPORTING ip_columnrow = ip_columnrow
                                             ip_column    = ip_column
                                             ip_row       = ip_row
                                   IMPORTING ep_column    = lv_column
                                             ep_row       = lv_row ).

* Begin of change issue #152 - don't touch exisiting style if only value is passed
    IF ip_column_formula_id <> 0.
      check_cell_column_formula(
          it_column_formulas   = column_formulas
          ip_column_formula_id = ip_column_formula_id
          ip_formula           = ip_formula
          ip_value             = ip_value
          ip_row               = lv_row
          ip_column            = lv_column ).
    ENDIF.
    READ TABLE sheet_content ASSIGNING <fs_sheet_content> WITH TABLE KEY cell_row    = lv_row      " Changed to access via table key , Stefan Schmöcker, 2013-08-03
                                                                         cell_column = lv_column.
    IF sy-subrc = 0.
      IF ip_style IS INITIAL.
        " If no style is provided as method-parameter and cell is found use cell's current style
        lv_style_guid = <fs_sheet_content>-cell_style.
      ELSE.
        " Style provided as method-parameter --> use this
        lv_style_guid = normalize_style_parameter( ip_style ).
      ENDIF.
    ELSE.
      " No cell found --> use supplied style even if empty
      lv_style_guid = normalize_style_parameter( ip_style ).
    ENDIF.
* End of change issue #152 - don't touch exisiting style if only value is passed

    IF ip_value IS SUPPLIED.
      "if data type is passed just write the value. Otherwise map abap type to excel and perform conversion
      "IP_DATA_TYPE is passed by excel reader so source types are preserved
*First we get reference into local var.
      IF ip_conv_exit_length = abap_true.
        lo_value = create_data_conv_exit_length( ip_value ).
      ELSE.
        CREATE DATA lo_value LIKE ip_value.
      ENDIF.
      ASSIGN lo_value->* TO <fs_value>.
      <fs_value> = ip_value.
      IF ip_data_type IS SUPPLIED.
        IF ip_abap_type IS NOT SUPPLIED.
          get_value_type( EXPORTING ip_value      = ip_value
                          IMPORTING ep_value      = <fs_value> ) .
        ENDIF.
        lv_value = <fs_value>.
        lv_data_type = ip_data_type.
      ELSE.
        IF ip_abap_type IS SUPPLIED.
          lv_value_type = ip_abap_type.
        ELSE.
          get_value_type( EXPORTING ip_value      = ip_value
                          IMPORTING ep_value      = <fs_value>
                                    ep_value_type = lv_value_type ).
        ENDIF.

        ASSIGN ('CL_ABAP_TYPEDESCR=>TYPEKIND_INT8') TO <fs_typekind_int8>.
        IF sy-subrc <> 0.
          ASSIGN space TO <fs_typekind_int8>. "not used as typekind!
        ENDIF.

        CASE lv_value_type.
          WHEN cl_abap_typedescr=>typekind_int OR cl_abap_typedescr=>typekind_int1 OR cl_abap_typedescr=>typekind_int2
            OR <fs_typekind_int8>. "Allow INT8 types columns
            IF lv_value_type = <fs_typekind_int8>.
              CALL METHOD cl_abap_elemdescr=>('GET_INT8') RECEIVING p_result = lo_addit.
            ELSE.
              lo_addit = cl_abap_elemdescr=>get_i( ).
            ENDIF.
            CREATE DATA lo_value_new TYPE HANDLE lo_addit.
            ASSIGN lo_value_new->* TO <fs_numeric>.
            IF sy-subrc = 0.
              <fs_numeric> = <fs_value>.
              lv_value = zcl_excel_common=>number_to_excel_string( ip_value = <fs_numeric> ).
            ENDIF.

          WHEN cl_abap_typedescr=>typekind_float OR cl_abap_typedescr=>typekind_packed OR
               cl_abap_typedescr=>typekind_decfloat OR
               cl_abap_typedescr=>typekind_decfloat16 OR
               cl_abap_typedescr=>typekind_decfloat34.
            IF lv_value_type = cl_abap_typedescr=>typekind_packed
                AND ip_currency IS NOT INITIAL.
              lv_value = zcl_excel_common=>number_to_excel_string( ip_value    = <fs_value>
                                                                   ip_currency = ip_currency ).
            ELSE.
              lo_addit = cl_abap_elemdescr=>get_f( ).
              CREATE DATA lo_value_new TYPE HANDLE lo_addit.
              ASSIGN lo_value_new->* TO <fs_numeric>.
              IF sy-subrc = 0.
                <fs_numeric> = <fs_value>.
                lv_value = zcl_excel_common=>number_to_excel_string( ip_value = <fs_numeric> ).
              ENDIF.
            ENDIF.

          WHEN cl_abap_typedescr=>typekind_char OR cl_abap_typedescr=>typekind_string OR cl_abap_typedescr=>typekind_num OR
               cl_abap_typedescr=>typekind_hex.
            lv_value = <fs_value>.
            lv_data_type = 's'.

          WHEN cl_abap_typedescr=>typekind_date.
            lo_addit = cl_abap_elemdescr=>get_d( ).
            CREATE DATA lo_value_new TYPE HANDLE lo_addit.
            ASSIGN lo_value_new->* TO <fs_date>.
            IF sy-subrc = 0.
              <fs_date> = <fs_value>.
              lv_value = zcl_excel_common=>date_to_excel_string( ip_value = <fs_date> ) .
            ENDIF.
* Begin of change issue #152 - don't touch exisiting style if only value is passed
* Moved to end of routine - apply date-format even if other styleinformation is passed
*          IF ip_style IS NOT SUPPLIED. "get default date format in case parameter is initial
*            lo_style = excel->add_new_style( ).
*            lo_style->number_format->format_code = get_default_excel_date_format( ).
*            lv_style_guid = lo_style->get_guid( ).
*          ENDIF.
* End of change issue #152 - don't touch exisiting style if only value is passed

          WHEN cl_abap_typedescr=>typekind_time.
            lo_addit = cl_abap_elemdescr=>get_t( ).
            CREATE DATA lo_value_new TYPE HANDLE lo_addit.
            ASSIGN lo_value_new->* TO <fs_time>.
            IF sy-subrc = 0.
              <fs_time> = <fs_value>.
              lv_value = zcl_excel_common=>time_to_excel_string( ip_value = <fs_time> ).
            ENDIF.
* Begin of change issue #152 - don't touch exisiting style if only value is passed
* Moved to end of routine - apply time-format even if other styleinformation is passed
*          IF ip_style IS NOT SUPPLIED. "get default time format for user in case parameter is initial
*            lo_style = excel->add_new_style( ).
*            lo_style->number_format->format_code = zcl_excel_style_number_format=>c_format_date_time6.
*            lv_style_guid = lo_style->get_guid( ).
*          ENDIF.
* End of change issue #152 - don't touch exisiting style if only value is passed

          WHEN OTHERS.
            zcx_excel=>raise_text( 'Invalid data type of input value' ).
        ENDCASE.
      ENDIF.

      IF <fs_sheet_content> IS ASSIGNED AND <fs_sheet_content>-table_header IS NOT INITIAL AND lv_value IS NOT INITIAL.
        READ TABLE <fs_sheet_content>-table->fieldcat ASSIGNING <ls_fieldcat> WITH KEY fieldname = <fs_sheet_content>-table_fieldname.
        IF sy-subrc = 0.
          <ls_fieldcat>-column_name = lv_value.
          IF <ls_fieldcat>-column_name <> lv_value.
            zcx_excel=>raise_text( 'Cell is table column header - this value is not allowed' ).
          ENDIF.
        ENDIF.
      ENDIF.

    ENDIF.

    IF ip_hyperlink IS BOUND.
      ip_hyperlink->set_cell_reference( ip_column = lv_column
                                        ip_row = lv_row ).
      me->hyperlinks->add( ip_hyperlink ).
    ENDIF.

    IF lv_value CS '_x'.
      " Issue #761 value "_x0041_" rendered as "A".
      " "_x...._", where "." is 0-9 a-f or A-F (case insensitive), is an internal value in sharedStrings.xml
      " that Excel uses to store special characters, it's interpreted like Unicode character U+....
      " for instance "_x0041_" is U+0041 which is "A".
      " To not interpret such text, the first underscore is replaced with "_x005f_".
      " The value "_x0041_" is to be stored internally "_x005f_x0041_" so that it's rendered like "_x0041_".
      " Note that REGEX is time consuming, it's why "CS" is used above to improve the performance.
      REPLACE ALL OCCURRENCES OF REGEX '_(x[0-9a-fA-F]{4}_)' IN lv_value WITH '_x005f_$1' RESPECTING CASE.
    ENDIF.

* Begin of change issue #152 - don't touch exisiting style if only value is passed
* Read table moved up, so that current style may be evaluated

    IF <fs_sheet_content> IS ASSIGNED.
* End of change issue #152 - don't touch exisiting style if only value is passed
      <fs_sheet_content>-cell_value   = lv_value.
      <fs_sheet_content>-cell_formula = ip_formula.
      <fs_sheet_content>-column_formula_id = ip_column_formula_id.
      <fs_sheet_content>-cell_style   = lv_style_guid.
      <fs_sheet_content>-data_type    = lv_data_type.
    ELSE.
      ls_sheet_content-cell_row     = lv_row.
      ls_sheet_content-cell_column  = lv_column.
      ls_sheet_content-cell_value   = lv_value.
      ls_sheet_content-cell_formula = ip_formula.
      ls_sheet_content-column_formula_id = ip_column_formula_id.
      ls_sheet_content-cell_style   = lv_style_guid.
      ls_sheet_content-data_type    = lv_data_type.
      ls_sheet_content-cell_coords  = zcl_excel_common=>convert_column_a_row2columnrow( i_column = lv_column i_row = lv_row ).
      INSERT ls_sheet_content INTO TABLE sheet_content ASSIGNING <fs_sheet_content>. "ins #152 - Now <fs_sheet_content> always holds the data

    ENDIF.

    IF ip_formula IS INITIAL AND lv_value IS NOT INITIAL AND it_rtf IS NOT INITIAL.
      lt_rtf = it_rtf.
      check_rtf( EXPORTING ip_value = lv_value
                           ip_style = lv_style_guid
                 CHANGING  ct_rtf   = lt_rtf ).
      <fs_sheet_content>-rtf_tab = lt_rtf.
    ENDIF.

* Begin of change issue #152 - don't touch exisiting style if only value is passed
* For Date- or Timefields change the formatcode if nothing is set yet
* Enhancement option:  Check if existing formatcode is a date/ or timeformat
*                      If not, use default
    DATA: lo_format_code_datetime TYPE zexcel_number_format.
    DATA: stylemapping    TYPE zexcel_s_stylemapping.
    IF <fs_sheet_content>-cell_style IS INITIAL.
      <fs_sheet_content>-cell_style = me->excel->get_default_style( ).
    ENDIF.
    CASE lv_value_type.
      WHEN cl_abap_typedescr=>typekind_date.
        TRY.
            stylemapping = me->excel->get_style_to_guid( <fs_sheet_content>-cell_style ).
          CATCH zcx_excel .
        ENDTRY.
        IF stylemapping-complete_stylex-number_format-format_code IS INITIAL OR
           stylemapping-complete_style-number_format-format_code IS INITIAL.
          lo_format_code_datetime = zcl_excel_style_number_format=>c_format_date_std.
        ELSE.
          lo_format_code_datetime = stylemapping-complete_style-number_format-format_code.
        ENDIF.
        me->change_cell_style( ip_column                      = lv_column
                               ip_row                         = lv_row
                               ip_number_format_format_code   = lo_format_code_datetime ).

      WHEN cl_abap_typedescr=>typekind_time.
        TRY.
            stylemapping = me->excel->get_style_to_guid( <fs_sheet_content>-cell_style ).
          CATCH zcx_excel .
        ENDTRY.
        IF stylemapping-complete_stylex-number_format-format_code IS INITIAL OR
           stylemapping-complete_style-number_format-format_code IS INITIAL.
          lo_format_code_datetime = zcl_excel_style_number_format=>c_format_date_time6.
        ELSE.
          lo_format_code_datetime = stylemapping-complete_style-number_format-format_code.
        ENDIF.
        me->change_cell_style( ip_column                      = lv_column
                               ip_row                         = lv_row
                               ip_number_format_format_code   = lo_format_code_datetime ).

    ENDCASE.
* End of change issue #152 - don't touch exisiting style if only value is passed

* Fix issue #162
    lv_value = ip_value.
    IF lv_value CS cl_abap_char_utilities=>cr_lf.
      me->change_cell_style( ip_column               = lv_column
                             ip_row                  = lv_row
                             ip_alignment_wraptext   = abap_true ).
    ENDIF.
* End of Fix issue #162

  ENDMETHOD.                    "SET_CELL