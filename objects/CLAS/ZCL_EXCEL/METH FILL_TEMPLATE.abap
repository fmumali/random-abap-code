  METHOD fill_template.

    DATA: lo_template_filler TYPE REF TO zcl_excel_fill_template.

    FIELD-SYMBOLS:
      <lv_sheet>     TYPE zexcel_sheet_title,
      <lv_data_line> TYPE zcl_excel_template_data=>ts_template_data_sheet.


    lo_template_filler = zcl_excel_fill_template=>create( me ).

    LOOP AT lo_template_filler->mt_sheet ASSIGNING <lv_sheet>.

      READ TABLE iv_data->mt_data ASSIGNING <lv_data_line> WITH KEY sheet = <lv_sheet>.
      CHECK sy-subrc = 0.
      lo_template_filler->fill_sheet( <lv_data_line> ).

    ENDLOOP.

  ENDMETHOD.