  METHOD class_constructor.

    c_messages-formula_id_only_is_possible = |{ 'If Formula ID is used, value and formula must be empty'(008) }|.
    c_messages-column_formula_id_not_found = |{ 'The Column Formula does not exist'(009) }|.
    c_messages-formula_not_in_this_table = |{ 'The cell uses a Column Formula which should be part of the same table'(010) }|.
    c_messages-formula_in_other_column = |{ 'The cell uses a Column Formula which is in a different column'(011) }|.

  ENDMETHOD.