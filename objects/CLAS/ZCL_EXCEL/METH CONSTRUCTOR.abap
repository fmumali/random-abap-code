  METHOD constructor.
    DATA: lo_style      TYPE REF TO zcl_excel_style.

* Inizialize instance objects
    CREATE OBJECT security.
    CREATE OBJECT worksheets.
    CREATE OBJECT ranges.
    CREATE OBJECT styles.
    CREATE OBJECT drawings
      EXPORTING
        ip_type = zcl_excel_drawing=>type_image.
    CREATE OBJECT charts
      EXPORTING
        ip_type = zcl_excel_drawing=>type_chart.
    CREATE OBJECT comments.
    CREATE OBJECT legacy_palette.
    CREATE OBJECT autofilters.

    me->zif_excel_book_protection~initialize( ).
    me->zif_excel_book_properties~initialize( ).

    TRY.
        me->add_new_worksheet( ).
      CATCH zcx_excel. " suppress syntax check error
        ASSERT 1 = 2.  " some error processing anyway
    ENDTRY.

    me->add_new_style( ). " Standard style
    lo_style = me->add_new_style( ). " Standard style with fill gray125
    lo_style->fill->filltype = zcl_excel_style_fill=>c_fill_pattern_gray125.

  ENDMETHOD.