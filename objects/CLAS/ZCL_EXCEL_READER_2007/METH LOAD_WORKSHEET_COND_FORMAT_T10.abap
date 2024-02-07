  METHOD load_worksheet_cond_format_t10.
    DATA: lv_dxf_style_index TYPE i.

    FIELD-SYMBOLS: <ls_dxf_style> LIKE LINE OF me->mt_dxf_styles.

    io_style_cond->mode_top10-topxx_count  = io_ixml_rule->get_attribute_ns( 'rank' ).        " Top10, Top20, Top 50...

    io_style_cond->mode_top10-percent      = io_ixml_rule->get_attribute_ns( 'percent' ).     " Top10 percent instead of Top10 values
    IF io_style_cond->mode_top10-percent = '1'.
      io_style_cond->mode_top10-percent = 'X'.
    ELSE.
      io_style_cond->mode_top10-percent = ' '.
    ENDIF.

    io_style_cond->mode_top10-bottom       = io_ixml_rule->get_attribute_ns( 'bottom' ).      " Bottom10 instead of Top10
    IF io_style_cond->mode_top10-bottom = '1'.
      io_style_cond->mode_top10-bottom = 'X'.
    ELSE.
      io_style_cond->mode_top10-bottom = ' '.
    ENDIF.
*--------------------------------------------------------------------*
* Cell formatting for top10
*--------------------------------------------------------------------*
    lv_dxf_style_index  = io_ixml_rule->get_attribute_ns( 'dxfId' ).
    READ TABLE me->mt_dxf_styles ASSIGNING <ls_dxf_style> WITH KEY dxf = lv_dxf_style_index.
    IF sy-subrc = 0.
      io_style_cond->mode_top10-cell_style = <ls_dxf_style>-guid.
    ENDIF.

  ENDMETHOD.