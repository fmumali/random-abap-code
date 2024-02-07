  METHOD is_row_hidden.


    DATA: lr_filter TYPE REF TO ts_filter,
          lv_col    TYPE i.

    FIELD-SYMBOLS: <ls_filter> TYPE ts_filter.

    rv_is_hidden = abap_false.

*--------------------------------------------------------------------*
* 1st row of filter area is never hidden, because here the filter
* symbol is being shown
*--------------------------------------------------------------------*
    IF iv_row = me->filter_area-row_start.
      RETURN.
    ENDIF.


    lv_col = me->filter_area-col_start.


    WHILE lv_col <= me->filter_area-col_end.

      lr_filter = me->get_column_filter( lv_col ).
      ASSIGN lr_filter->* TO <ls_filter>.

      CASE <ls_filter>-rule.

        WHEN mc_filter_rule_single_values.
          rv_is_hidden = me->is_row_hidden_single_values( iv_row    = iv_row
                                                          iv_col    = lv_col
                                                          is_filter = <ls_filter> ).

        WHEN mc_filter_rule_text_pattern.
          rv_is_hidden = me->is_row_hidden_text_pattern(  iv_row    = iv_row
                                                          iv_col    = lv_col
                                                          is_filter = <ls_filter> ).

      ENDCASE.

      IF rv_is_hidden = abap_true.
        RETURN.
      ENDIF.


      ADD 1 TO lv_col.

    ENDWHILE.


  ENDMETHOD.