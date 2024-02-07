  METHOD check_table_overlapping.

    DATA: lv_errormessage TYPE string,
          lv_column_int   TYPE zexcel_cell_column,
          lv_maxcol       TYPE i.
    FIELD-SYMBOLS:
          <ls_table_settings> TYPE zexcel_s_table_settings.

    lv_column_int = zcl_excel_common=>convert_column2int( is_table_settings-top_left_column ).
    lv_maxcol = zcl_excel_common=>convert_column2int( is_table_settings-bottom_right_column ).

    LOOP AT it_other_table_settings ASSIGNING <ls_table_settings>.

      IF  (    (  is_table_settings-top_left_row     GE <ls_table_settings>-top_left_row
              AND is_table_settings-top_left_row     LE <ls_table_settings>-bottom_right_row )
            OR
               (  is_table_settings-bottom_right_row GE <ls_table_settings>-top_left_row
              AND is_table_settings-bottom_right_row LE <ls_table_settings>-bottom_right_row )
          )
        AND
          (    (  lv_column_int GE zcl_excel_common=>convert_column2int( <ls_table_settings>-top_left_column )
              AND lv_column_int LE zcl_excel_common=>convert_column2int( <ls_table_settings>-bottom_right_column ) )
            OR
               (  lv_maxcol     GE zcl_excel_common=>convert_column2int( <ls_table_settings>-top_left_column )
              AND lv_maxcol     LE zcl_excel_common=>convert_column2int( <ls_table_settings>-bottom_right_column ) )
          ).
        lv_errormessage = 'Table overlaps with previously bound table and will not be added to worksheet.'(400).
        zcx_excel=>raise_text( lv_errormessage ).
      ENDIF.

    ENDLOOP.

  ENDMETHOD.