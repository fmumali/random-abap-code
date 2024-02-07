  METHOD set_value.
    DATA: lv_start_row_c TYPE c LENGTH 7,
          lv_stop_row_c  TYPE c LENGTH 7,
          lv_value       TYPE string.
    lv_stop_row_c = ip_stop_row.
    SHIFT lv_stop_row_c RIGHT DELETING TRAILING space.
    SHIFT lv_stop_row_c LEFT DELETING LEADING space.
    lv_start_row_c = ip_start_row.
    SHIFT lv_start_row_c RIGHT DELETING TRAILING space.
    SHIFT lv_start_row_c LEFT DELETING LEADING space.
    lv_value = ip_sheet_name.
    me->value = zcl_excel_common=>escape_string( ip_value = lv_value ).

    IF ip_stop_column IS INITIAL AND ip_stop_row IS INITIAL.
      CONCATENATE me->value '!$' ip_start_column '$' lv_start_row_c INTO me->value.
    ELSE.
      CONCATENATE me->value '!$' ip_start_column '$' lv_start_row_c ':$' ip_stop_column '$' lv_stop_row_c INTO me->value.
    ENDIF.
  ENDMETHOD.