  METHOD process_header_footer.

* ----------------------------------------------------------------------
* Only Basic font/text formatting possible:
* Bold (yes / no), Font Type, Font Size

    DATA:   lv_fname(12) TYPE c
          , lv_string    TYPE string
          .

    FIELD-SYMBOLS:   <lv_value> TYPE string
                   , <ls_font>  TYPE zexcel_s_style_font
                   .

* ----------------------------------------------------------------------
    CONCATENATE ip_side '_VALUE' INTO lv_fname.
    ASSIGN COMPONENT lv_fname OF STRUCTURE ip_header TO <lv_value>.

    CONCATENATE ip_side '_FONT' INTO lv_fname.
    ASSIGN COMPONENT lv_fname OF STRUCTURE ip_header TO <ls_font>.

    IF <ls_font> IS ASSIGNED AND <lv_value> IS ASSIGNED.

      IF <lv_value> = '&G'. "image header
        rv_processed_string = <lv_value>.
      ELSE.

        IF <ls_font>-name IS NOT INITIAL.
          CONCATENATE '&"' <ls_font>-name ',' INTO rv_processed_string.
        ELSE.
          rv_processed_string = '&"-,'.
        ENDIF.

        IF <ls_font>-bold = abap_true.
          CONCATENATE rv_processed_string 'Bold"' INTO rv_processed_string.
        ELSE.
          CONCATENATE rv_processed_string 'Standard"' INTO rv_processed_string.
        ENDIF.

        IF <ls_font>-size IS NOT INITIAL.
          lv_string = <ls_font>-size.
          CONCATENATE rv_processed_string '&' lv_string INTO rv_processed_string.
          CONDENSE rv_processed_string NO-GAPS.
        ENDIF.

        CONCATENATE rv_processed_string <lv_value> INTO rv_processed_string.
      ENDIF.
    ENDIF.
* ----------------------------------------------------------------------

  ENDMETHOD.