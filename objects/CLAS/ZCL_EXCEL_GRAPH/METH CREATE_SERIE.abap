  METHOD create_serie.
    DATA ls_serie TYPE s_series.

    DATA: lv_start_row_c TYPE c LENGTH 7,
          lv_stop_row_c  TYPE c LENGTH 7.


    IF ip_lbl IS NOT SUPPLIED.
      lv_stop_row_c = ip_lbl_to_row.
      SHIFT lv_stop_row_c RIGHT DELETING TRAILING space.
      SHIFT lv_stop_row_c LEFT DELETING LEADING space.
      lv_start_row_c = ip_lbl_from_row.
      SHIFT lv_start_row_c RIGHT DELETING TRAILING space.
      SHIFT lv_start_row_c LEFT DELETING LEADING space.
      ls_serie-lbl = ip_sheet.
      ls_serie-lbl = zcl_excel_common=>escape_string( ip_value = ls_serie-lbl ).
      CONCATENATE ls_serie-lbl '!$' ip_lbl_from_col '$' lv_start_row_c ':$' ip_lbl_to_col '$' lv_stop_row_c INTO ls_serie-lbl.
      CLEAR: lv_start_row_c, lv_stop_row_c.
    ELSE.
      ls_serie-lbl = ip_lbl.
    ENDIF.
    IF ip_ref IS NOT SUPPLIED.
      lv_stop_row_c = ip_ref_to_row.
      SHIFT lv_stop_row_c RIGHT DELETING TRAILING space.
      SHIFT lv_stop_row_c LEFT DELETING LEADING space.
      lv_start_row_c = ip_ref_from_row.
      SHIFT lv_start_row_c RIGHT DELETING TRAILING space.
      SHIFT lv_start_row_c LEFT DELETING LEADING space.
      ls_serie-ref = ip_sheet.
      ls_serie-ref = zcl_excel_common=>escape_string( ip_value = ls_serie-ref ).
      CONCATENATE ls_serie-ref '!$' ip_ref_from_col '$' lv_start_row_c ':$' ip_ref_to_col '$' lv_stop_row_c INTO ls_serie-ref.
      CLEAR: lv_start_row_c, lv_stop_row_c.
    ELSE.
      ls_serie-ref = ip_ref.
    ENDIF.
    ls_serie-idx = ip_idx.
    ls_serie-order = ip_order.
    ls_serie-invertifnegative = ip_invertifnegative.
    ls_serie-symbol = ip_symbol.
    ls_serie-smooth = ip_smooth.
    ls_serie-sername = ip_sername.
    APPEND ls_serie TO me->series.
    SORT me->series BY order ASCENDING.
  ENDMETHOD.