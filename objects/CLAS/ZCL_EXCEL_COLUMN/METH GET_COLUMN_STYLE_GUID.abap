  METHOD get_column_style_guid.
    IF me->style_guid IS NOT INITIAL.
      ep_style_guid = me->style_guid.
    ELSE.
      ep_style_guid = me->worksheet->zif_excel_sheet_properties~get_style( ).
    ENDIF.
  ENDMETHOD.