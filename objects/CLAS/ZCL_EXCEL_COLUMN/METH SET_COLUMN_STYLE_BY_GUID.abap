  METHOD set_column_style_by_guid.
    DATA: stylemapping TYPE zexcel_s_stylemapping.

    IF me->excel IS NOT BOUND.
      zcx_excel=>raise_text( 'Internal error - reference to ZCL_EXCEL not bound' ).
    ENDIF.
    TRY.
        stylemapping = me->excel->get_style_to_guid( ip_style_guid ).
        me->style_guid = stylemapping-guid.

      CATCH zcx_excel .
        RETURN.  " leave as is in case of error
    ENDTRY.

  ENDMETHOD.