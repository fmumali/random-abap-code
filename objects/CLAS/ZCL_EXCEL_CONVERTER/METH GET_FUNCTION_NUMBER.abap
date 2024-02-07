  METHOD get_function_number.
*Number    Function
*1  AVERAGE
*2  COUNT
*3  COUNTA
*4  MAX
*5  MIN
*6  PRODUCT
*7  STDEV
*8  STDEVP
*9  SUM
*10  VAR
*11  VARP

    CASE i_totals_function.
      WHEN zcl_excel_table=>totals_function_sum.     " Total
        r_function_number = 9.
      WHEN zcl_excel_table=>totals_function_min.     " Minimum
        r_function_number = 5.
      WHEN zcl_excel_table=>totals_function_max.     " Maximum
        r_function_number = 4.
      WHEN zcl_excel_table=>totals_function_average. " Mean Value
        r_function_number = 1.
      WHEN zcl_excel_table=>totals_function_count.   " Count
        r_function_number = 2.
      WHEN OTHERS.
        CLEAR r_function_number.
    ENDCASE.
  ENDMETHOD.