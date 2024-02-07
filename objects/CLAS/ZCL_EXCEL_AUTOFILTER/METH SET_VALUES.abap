  METHOD set_values.

    FIELD-SYMBOLS: <ls_value> LIKE LINE OF it_values.

    LOOP AT it_values ASSIGNING <ls_value>.

      me->set_value( i_column = <ls_value>-column
                     i_value  = <ls_value>-value ).

    ENDLOOP.

  ENDMETHOD.