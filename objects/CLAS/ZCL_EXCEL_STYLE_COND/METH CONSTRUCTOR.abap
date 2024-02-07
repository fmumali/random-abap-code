  METHOD constructor.

    DATA: ls_iconset TYPE zexcel_conditional_iconset.
    ls_iconset-iconset     = zcl_excel_style_cond=>c_iconset_3trafficlights.
    ls_iconset-cfvo1_type  = zcl_excel_style_cond=>c_cfvo_type_percent.
    ls_iconset-cfvo1_value = '0'.
    ls_iconset-cfvo2_type  = zcl_excel_style_cond=>c_cfvo_type_percent.
    ls_iconset-cfvo2_value = '20'.
    ls_iconset-cfvo3_type  = zcl_excel_style_cond=>c_cfvo_type_percent.
    ls_iconset-cfvo3_value = '40'.
    ls_iconset-cfvo4_type  = zcl_excel_style_cond=>c_cfvo_type_percent.
    ls_iconset-cfvo4_value = '60'.
    ls_iconset-cfvo5_type  = zcl_excel_style_cond=>c_cfvo_type_percent.
    ls_iconset-cfvo5_value = '80'.


    me->rule          = zcl_excel_style_cond=>c_rule_none.
    me->mode_iconset  = ls_iconset.
    me->priority      = 1.

* inizialize dimension range
    me->mv_rule_range     = ip_dimension_range.

    IF ip_guid IS NOT INITIAL.
      me->guid = ip_guid.
    ELSE.
      me->guid = zcl_excel_obsolete_func_wrap=>guid_create( ).
    ENDIF.

  ENDMETHOD.