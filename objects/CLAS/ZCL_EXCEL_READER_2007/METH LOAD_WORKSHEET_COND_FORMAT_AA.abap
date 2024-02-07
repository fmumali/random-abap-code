  METHOD load_worksheet_cond_format_aa.
    DATA: lv_dxf_style_index TYPE i,
          val                TYPE string.

    FIELD-SYMBOLS: <ls_dxf_style> LIKE LINE OF me->mt_dxf_styles.

*--------------------------------------------------------------------*
* above or below average
*--------------------------------------------------------------------*
    val  = io_ixml_rule->get_attribute_ns( 'aboveAverage' ).
    IF val = '0'.  " 0 = below average
      io_style_cond->mode_above_average-above_average = space.
    ELSE.
      io_style_cond->mode_above_average-above_average = 'X'. " Not present or <> 0 --> we use above average
    ENDIF.

*--------------------------------------------------------------------*
* Equal average also?
*--------------------------------------------------------------------*
    CLEAR val.
    val  = io_ixml_rule->get_attribute_ns( 'equalAverage' ).
    IF val = '1'.  " 0 = below average
      io_style_cond->mode_above_average-equal_average = 'X'.
    ELSE.
      io_style_cond->mode_above_average-equal_average = ' '. " Not present or <> 1 --> we use not equal average
    ENDIF.

*--------------------------------------------------------------------*
* Standard deviation instead of value ( 2nd stddev, 3rd stdev )
*--------------------------------------------------------------------*
    CLEAR val.
    val  = io_ixml_rule->get_attribute_ns( 'stdDev' ).
    CASE val.
      WHEN 1
        OR 2
        OR 3.  " These seem to be supported by excel - don't try anything more
        io_style_cond->mode_above_average-standard_deviation = val.
    ENDCASE.

*--------------------------------------------------------------------*
* Cell formatting for top10
*--------------------------------------------------------------------*
    lv_dxf_style_index  = io_ixml_rule->get_attribute_ns( 'dxfId' ).
    READ TABLE me->mt_dxf_styles ASSIGNING <ls_dxf_style> WITH KEY dxf = lv_dxf_style_index.
    IF sy-subrc = 0.
      io_style_cond->mode_above_average-cell_style = <ls_dxf_style>-guid.
    ENDIF.

  ENDMETHOD.