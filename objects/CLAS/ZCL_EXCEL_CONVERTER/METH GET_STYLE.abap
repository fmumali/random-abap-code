  METHOD get_style.
    DATA: ls_styles TYPE ts_styles,
          lo_style  TYPE REF TO zcl_excel_style.

    CLEAR r_style.

    READ TABLE wt_styles INTO ls_styles WITH TABLE KEY type      = i_type
                                                       alignment = i_alignment
                                                       inttype   = i_inttype
                                                       decimals  = i_decimals.
    IF sy-subrc = 0.
      r_style = ls_styles-guid.
    ELSE.
      CASE i_type.
        WHEN c_type_hdr. " Header
          lo_style = create_style_hdr( i_alignment = i_alignment ).
        WHEN c_type_str. "Stripped
          lo_style   = create_style_stripped( i_alignment = i_alignment
                                              i_inttype   = i_inttype
                                              i_decimals  = i_decimals   ).
        WHEN c_type_nor. "Normal
          lo_style   = create_style_normal( i_alignment = i_alignment
                                            i_inttype   = i_inttype
                                            i_decimals  = i_decimals   ).
        WHEN c_type_sub. "Subtotals
          lo_style   = create_style_subtotal( i_alignment = i_alignment
                                              i_inttype   = i_inttype
                                              i_decimals  = i_decimals   ).
        WHEN c_type_tot. "Totals
          lo_style   = create_style_total( i_alignment = i_alignment
                                           i_inttype   = i_inttype
                                           i_decimals  = i_decimals   ).
      ENDCASE.
      IF lo_style IS NOT INITIAL.
        r_style = lo_style->get_guid( ).
        ls_styles-type       = i_type.
        ls_styles-alignment  = i_alignment.
        ls_styles-inttype    = i_inttype.
        ls_styles-decimals   = i_decimals.
        ls_styles-guid       = r_style.
        ls_styles-style      = lo_style.
        INSERT ls_styles INTO TABLE wt_styles.
      ENDIF.
    ENDIF.
  ENDMETHOD.