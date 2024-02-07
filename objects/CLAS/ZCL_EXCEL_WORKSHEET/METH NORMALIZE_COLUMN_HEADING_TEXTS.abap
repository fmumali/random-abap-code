  METHOD normalize_column_heading_texts.

    DATA: lt_field_catalog      TYPE zexcel_t_fieldcatalog,
          lv_value_lowercase    TYPE string,
          lv_scrtext_l_initial  TYPE zexcel_column_name,
          lv_long_text          TYPE string,
          lv_max_length         TYPE i,
          lv_temp_length        TYPE i,
          lv_syindex            TYPE c LENGTH 3,
          lt_column_name_buffer TYPE SORTED TABLE OF string WITH UNIQUE KEY table_line.
    FIELD-SYMBOLS: <ls_field_catalog> TYPE zexcel_s_fieldcatalog,
                   <scrtxt1>          TYPE any,
                   <scrtxt2>          TYPE any,
                   <scrtxt3>          TYPE any.

    " Due to restrictions in new table object we cannot have two columns with the same name
    " Check if a column with the same name exists, if exists add a counter
    " If no medium description is provided we try to use small or long

    lt_field_catalog = it_field_catalog.

    LOOP AT lt_field_catalog ASSIGNING <ls_field_catalog> WHERE dynpfld EQ abap_true.

      IF <ls_field_catalog>-column_name IS INITIAL.

        CASE iv_default_descr.
          WHEN 'M'.
            ASSIGN <ls_field_catalog>-scrtext_m TO <scrtxt1>.
            ASSIGN <ls_field_catalog>-scrtext_s TO <scrtxt2>.
            ASSIGN <ls_field_catalog>-scrtext_l TO <scrtxt3>.
          WHEN 'S'.
            ASSIGN <ls_field_catalog>-scrtext_s TO <scrtxt1>.
            ASSIGN <ls_field_catalog>-scrtext_m TO <scrtxt2>.
            ASSIGN <ls_field_catalog>-scrtext_l TO <scrtxt3>.
          WHEN 'L'.
            ASSIGN <ls_field_catalog>-scrtext_l TO <scrtxt1>.
            ASSIGN <ls_field_catalog>-scrtext_m TO <scrtxt2>.
            ASSIGN <ls_field_catalog>-scrtext_s TO <scrtxt3>.
          WHEN OTHERS.
            ASSIGN <ls_field_catalog>-scrtext_m TO <scrtxt1>.
            ASSIGN <ls_field_catalog>-scrtext_s TO <scrtxt2>.
            ASSIGN <ls_field_catalog>-scrtext_l TO <scrtxt3>.
        ENDCASE.

        IF <scrtxt1> IS NOT INITIAL.
          <ls_field_catalog>-column_name = <scrtxt1>.
        ELSEIF <scrtxt2> IS NOT INITIAL.
          <ls_field_catalog>-column_name = <scrtxt2>.
        ELSEIF <scrtxt3> IS NOT INITIAL.
          <ls_field_catalog>-column_name = <scrtxt3>.
        ELSE.
          <ls_field_catalog>-column_name = 'Column'.  " default value as Excel does
        ENDIF.
      ENDIF.

      lv_scrtext_l_initial = <ls_field_catalog>-column_name.
      DESCRIBE FIELD <ls_field_catalog>-column_name LENGTH lv_max_length IN CHARACTER MODE.
      DO.
        lv_value_lowercase = <ls_field_catalog>-column_name.
        TRANSLATE lv_value_lowercase TO LOWER CASE.
        READ TABLE lt_column_name_buffer TRANSPORTING NO FIELDS WITH KEY table_line = lv_value_lowercase BINARY SEARCH.
        IF sy-subrc <> 0.
          INSERT lv_value_lowercase INTO TABLE lt_column_name_buffer.
          EXIT.
        ELSE.
          lv_syindex = sy-index.
          CONCATENATE lv_scrtext_l_initial lv_syindex INTO lv_long_text.
          IF strlen( lv_long_text ) <= lv_max_length.
            <ls_field_catalog>-column_name = lv_long_text.
          ELSE.
            lv_temp_length = strlen( lv_scrtext_l_initial ) - 1.
            lv_scrtext_l_initial = substring( val = lv_scrtext_l_initial len = lv_temp_length ).
            CONCATENATE lv_scrtext_l_initial lv_syindex INTO <ls_field_catalog>-column_name.
          ENDIF.
        ENDIF.
      ENDDO.

    ENDLOOP.

    result = lt_field_catalog.

  ENDMETHOD.