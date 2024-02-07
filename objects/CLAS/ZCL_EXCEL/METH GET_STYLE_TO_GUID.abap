  METHOD get_style_to_guid.
    DATA: lo_style TYPE REF TO zcl_excel_style.
    " # issue 139
    READ TABLE me->t_stylemapping2 INTO ep_stylemapping WITH TABLE KEY guid = ip_guid.
    IF sy-subrc <> 0.
      zcx_excel=>raise_text( 'GUID not found' ).
    ENDIF.

    IF ep_stylemapping-dynamic_style_guid IS NOT INITIAL.
      lo_style = me->get_style_from_guid( ip_guid ).
      zcl_excel_common=>recursive_class_to_struct( EXPORTING i_source = lo_style
                                                   CHANGING  e_target =  ep_stylemapping-complete_style
                                                             e_targetx = ep_stylemapping-complete_stylex ).
    ENDIF.
  ENDMETHOD.