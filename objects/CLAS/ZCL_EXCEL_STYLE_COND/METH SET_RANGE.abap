  METHOD set_range.

    CLEAR: me->mv_rule_range.

    me->add_range( ip_start_row    = ip_start_row
                   ip_start_column = ip_start_column
                   ip_stop_row     = ip_stop_row
                   ip_stop_column  = ip_stop_column ).

  ENDMETHOD.