  METHOD set_data.

    DATA lr_temp TYPE REF TO data.

    FIELD-SYMBOLS: <lt_table_temp> TYPE ANY TABLE,
                   <lt_table>      TYPE ANY TABLE.

    GET REFERENCE OF ir_data INTO lr_temp.
    ASSIGN lr_temp->* TO <lt_table_temp>.
    CREATE DATA table_data LIKE <lt_table_temp>.
    ASSIGN me->table_data->* TO <lt_table>.
    <lt_table> = <lt_table_temp>.

  ENDMETHOD.