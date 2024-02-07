  METHOD create_text_subtotal.
    DATA: l_string(256) TYPE c,
          l_func        TYPE string.

    CASE i_totals_function.
      WHEN zcl_excel_table=>totals_function_sum.     " Total
        l_func = 'Total'(003).
      WHEN zcl_excel_table=>totals_function_min.     " Minimum
        l_func = 'Minimum'(004).
      WHEN zcl_excel_table=>totals_function_max.     " Maximum
        l_func = 'Maximum'(005).
      WHEN zcl_excel_table=>totals_function_average. " Mean Value
        l_func = 'Average'(006).
      WHEN zcl_excel_table=>totals_function_count.   " Count
        l_func = 'Count'(007).
      WHEN OTHERS.
        CLEAR l_func.
    ENDCASE.

    l_string = i_value.

    CONCATENATE l_string l_func INTO r_text SEPARATED BY space.

  ENDMETHOD.