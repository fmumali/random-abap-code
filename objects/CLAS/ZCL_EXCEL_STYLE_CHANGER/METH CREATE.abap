  METHOD create.

    DATA: style TYPE REF TO zcl_excel_style_changer.

    CREATE OBJECT style.
    style->excel = excel.
    result = style.

  ENDMETHOD.