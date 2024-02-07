  METHOD check_rtf.

    DATA: lo_style           TYPE REF TO zcl_excel_style,
          lo_iterator        TYPE REF TO zcl_excel_collection_iterator,
          lv_next_rtf_offset TYPE i,
          lv_tabix           TYPE i,
          lv_value           TYPE string,
          lv_val_length      TYPE i,
          ls_rtf             LIKE LINE OF ct_rtf.
    FIELD-SYMBOLS: <rtf> LIKE LINE OF ct_rtf.

    IF ip_style IS NOT SUPPLIED.
      ip_style = excel->get_default_style( ).
    ENDIF.

    lo_iterator = excel->get_styles_iterator( ).
    WHILE lo_iterator->has_next( ) = abap_true.
      lo_style ?= lo_iterator->get_next( ).
      IF lo_style->get_guid( ) = ip_style.
        EXIT.
      ENDIF.
      CLEAR lo_style.
    ENDWHILE.

    lv_next_rtf_offset = 0.
    LOOP AT ct_rtf ASSIGNING <rtf>.
      lv_tabix = sy-tabix.
      IF lv_next_rtf_offset < <rtf>-offset.
        ls_rtf-offset = lv_next_rtf_offset.
        ls_rtf-length = <rtf>-offset - lv_next_rtf_offset.
        ls_rtf-font   = lo_style->font->get_structure( ).
        INSERT ls_rtf INTO ct_rtf INDEX lv_tabix.
      ELSEIF lv_next_rtf_offset > <rtf>-offset.
        RAISE EXCEPTION TYPE zcx_excel
          EXPORTING
            error = 'Gaps or overlaps in RTF data offset/length specs'.
      ENDIF.
      lv_next_rtf_offset = <rtf>-offset + <rtf>-length.
    ENDLOOP.

    lv_value = ip_value.
    lv_val_length = strlen( lv_value ).
    IF lv_val_length > lv_next_rtf_offset.
      ls_rtf-offset = lv_next_rtf_offset.
      ls_rtf-length = lv_val_length - lv_next_rtf_offset.
      ls_rtf-font   = lo_style->font->get_structure( ).
      INSERT ls_rtf INTO TABLE ct_rtf.
    ELSEIF lv_val_length > lv_next_rtf_offset.
      RAISE EXCEPTION TYPE zcx_excel
        EXPORTING
          error = 'RTF specs length is not equal to value length'.
    ENDIF.

  ENDMETHOD.