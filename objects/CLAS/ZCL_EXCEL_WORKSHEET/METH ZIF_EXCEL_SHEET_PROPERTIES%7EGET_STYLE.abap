  METHOD zif_excel_sheet_properties~get_style.
    IF zif_excel_sheet_properties~style IS NOT INITIAL.
      ep_style = zif_excel_sheet_properties~style.
    ELSE.
      ep_style = me->excel->get_default_style( ).
    ENDIF.
  ENDMETHOD.                    "ZIF_EXCEL_SHEET_PROPERTIES~GET_STYLE