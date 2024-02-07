  METHOD get_totals_formula.
    CONSTANTS: lc_function_id_sum     TYPE string VALUE '109',
               lc_function_id_min     TYPE string VALUE '105',
               lc_function_id_max     TYPE string VALUE '104',
               lc_function_id_count   TYPE string VALUE '103',
               lc_function_id_average TYPE string VALUE '101'.

    DATA: lv_function_id TYPE string.

    CASE ip_function.
      WHEN zcl_excel_table=>totals_function_sum.
        lv_function_id = lc_function_id_sum.

      WHEN zcl_excel_table=>totals_function_min.
        lv_function_id = lc_function_id_min.

      WHEN zcl_excel_table=>totals_function_max.
        lv_function_id = lc_function_id_max.

      WHEN zcl_excel_table=>totals_function_count.
        lv_function_id = lc_function_id_count.

      WHEN zcl_excel_table=>totals_function_average.
        lv_function_id = lc_function_id_average.

      WHEN zcl_excel_table=>totals_function_custom. " issue #292
        RETURN.

      WHEN OTHERS.
        zcx_excel=>raise_text( 'Invalid totals formula. See ZCL_ for possible values' ).
    ENDCASE.

    CONCATENATE 'SUBTOTAL(' lv_function_id ',[' ip_column '])' INTO ep_formula.
  ENDMETHOD.