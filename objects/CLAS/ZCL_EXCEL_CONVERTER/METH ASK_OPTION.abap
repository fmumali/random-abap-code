  METHOD ask_option.
    DATA: ls_sval      TYPE sval,
          lt_sval      TYPE STANDARD TABLE OF sval,
          l_returncode TYPE string,
          lt_fields    TYPE ddfields,
          ls_fields    TYPE dfies.

    FIELD-SYMBOLS: <fs> TYPE any.

    rs_option = ws_option.

    CALL FUNCTION 'DDIF_FIELDINFO_GET'
      EXPORTING
        tabname        = 'ZEXCEL_S_CONVERTER_OPTION'
      TABLES
        dfies_tab      = lt_fields
      EXCEPTIONS
        not_found      = 1
        internal_error = 2
        OTHERS         = 3.
    IF sy-subrc <> 0.
      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
              WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ENDIF.

    LOOP AT lt_fields INTO ls_fields.
      ASSIGN COMPONENT ls_fields-fieldname OF STRUCTURE ws_option TO <fs>.
      IF sy-subrc = 0.
        CLEAR ls_sval.
        ls_sval-tabname      = ls_fields-tabname.
        ls_sval-fieldname    = ls_fields-fieldname.
        ls_sval-value        = <fs>.
        ls_sval-field_attr   = space.
        ls_sval-field_obl    = space.
        ls_sval-comp_code    = space.
        ls_sval-fieldtext    = ls_fields-scrtext_m.
        ls_sval-comp_tab     = space.
        ls_sval-comp_field   = space.
        ls_sval-novaluehlp   = space.
        INSERT ls_sval INTO TABLE lt_sval.
      ENDIF.
    ENDLOOP.

    CALL FUNCTION 'POPUP_GET_VALUES'
      EXPORTING
        popup_title     = 'Excel creation options'(008)
      IMPORTING
        returncode      = l_returncode
      TABLES
        fields          = lt_sval
      EXCEPTIONS
        error_in_fields = 1
        OTHERS          = 2.
    IF sy-subrc <> 0.
      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
              WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ELSEIF l_returncode = 'A'.
      RAISE EXCEPTION TYPE zcx_excel.
    ELSE.
      LOOP AT lt_sval INTO ls_sval.
        ASSIGN COMPONENT ls_sval-fieldname OF STRUCTURE ws_option TO <fs>.
        IF sy-subrc = 0.
          <fs> = ls_sval-value.
        ENDIF.
      ENDLOOP.
      set_option( is_option = ws_option ) .
      rs_option = ws_option.
    ENDIF.
  ENDMETHOD.