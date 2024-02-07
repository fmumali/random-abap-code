  METHOD set_value.

    DATA: lr_filter TYPE REF TO ts_filter.

    FIELD-SYMBOLS: <ls_filter> TYPE ts_filter.


    lr_filter = me->get_column_filter(  i_column ).
    ASSIGN lr_filter->* TO <ls_filter>.

    <ls_filter>-rule     = mc_filter_rule_single_values.

    INSERT i_value INTO TABLE <ls_filter>-t_values.

  ENDMETHOD.