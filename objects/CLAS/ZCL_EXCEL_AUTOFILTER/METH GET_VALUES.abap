  METHOD get_values.

    FIELD-SYMBOLS: <ls_filter> LIKE LINE OF me->mt_filters,
                   <ls_value>  LIKE LINE OF <ls_filter>-t_values.

    DATA: ls_filter LIKE LINE OF rt_filter.

    LOOP AT me->mt_filters ASSIGNING <ls_filter> WHERE rule = mc_filter_rule_single_values.

      ls_filter-column = <ls_filter>-column.
      LOOP AT <ls_filter>-t_values ASSIGNING <ls_value>.
        ls_filter-value = <ls_value>.
        APPEND ls_filter TO rt_filter.
      ENDLOOP.

    ENDLOOP.

  ENDMETHOD.