  METHOD add.
    FIELD-SYMBOLS: <ls_data_sheet> TYPE ts_template_data_sheet,
                   <any>           TYPE any.

    APPEND INITIAL LINE TO mt_data ASSIGNING <ls_data_sheet>.
    <ls_data_sheet>-sheet = iv_sheet.
    CREATE DATA  <ls_data_sheet>-data LIKE iv_data.

    ASSIGN <ls_data_sheet>-data->* TO <any>.
    <any> = iv_data.

  ENDMETHOD.