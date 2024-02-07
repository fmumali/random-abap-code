  METHOD get_static_cellstyle_guid.
    " # issue 139
    DATA: style LIKE LINE OF me->t_stylemapping1.

    READ TABLE me->t_stylemapping1 INTO style
      WITH TABLE KEY dynamic_style_guid = style-guid  " no dynamic style  --> look for initial guid here
                     complete_style     = ip_cstyle_complete
                     complete_stylex    = ip_cstylex_complete.
    IF sy-subrc <> 0.
      style-complete_style  = ip_cstyle_complete.
      style-complete_stylex = ip_cstylex_complete.
      style-guid = zcl_excel_obsolete_func_wrap=>guid_create( ). " ins issue #379 - replacement for outdated function call
      INSERT style INTO TABLE me->t_stylemapping1.
      INSERT style INTO TABLE me->t_stylemapping2.

    ENDIF.

    ep_guid = style-guid.
  ENDMETHOD.