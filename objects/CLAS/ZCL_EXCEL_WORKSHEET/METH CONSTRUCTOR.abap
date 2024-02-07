  METHOD constructor.
    DATA: lv_title TYPE zexcel_sheet_title.

    me->excel = ip_excel.

    me->guid = zcl_excel_obsolete_func_wrap=>guid_create( ).        " ins issue #379 - replacement for outdated function call

    IF ip_title IS NOT INITIAL.
      lv_title = ip_title.
    ELSE.
      lv_title = me->generate_title( ). " ins issue #154 - Names of worksheets
    ENDIF.

    me->set_title( ip_title = lv_title ).

    CREATE OBJECT sheet_setup.
    CREATE OBJECT styles_cond.
    CREATE OBJECT data_validations.
    CREATE OBJECT tables.
    CREATE OBJECT columns.
    CREATE OBJECT rows.
    CREATE OBJECT ranges. " issue #163
    CREATE OBJECT mo_pagebreaks.
    CREATE OBJECT drawings
      EXPORTING
        ip_type = zcl_excel_drawing=>type_image.
    CREATE OBJECT charts
      EXPORTING
        ip_type = zcl_excel_drawing=>type_chart.
    me->zif_excel_sheet_protection~initialize( ).
    me->zif_excel_sheet_properties~initialize( ).
    CREATE OBJECT hyperlinks.
    CREATE OBJECT comments. " (+) Issue #180

* initialize active cell coordinates
    active_cell-cell_row = 1.
    active_cell-cell_column = 1.

* inizialize dimension range
    lower_cell-cell_row     = 1.
    lower_cell-cell_column  = 1.
    upper_cell-cell_row     = 1.
    upper_cell-cell_column  = 1.

  ENDMETHOD.                    "CONSTRUCTOR