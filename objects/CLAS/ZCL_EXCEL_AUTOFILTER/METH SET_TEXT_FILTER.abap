  METHOD set_text_filter.
*  see method documentation how to use this

    DATA: lr_filter TYPE REF TO ts_filter,
          ls_value1 TYPE LINE OF ts_filter-tr_textfilter1.

    FIELD-SYMBOLS: <ls_filter> TYPE ts_filter.


    lr_filter = me->get_column_filter(  i_column ).
    ASSIGN lr_filter->* TO <ls_filter>.

    <ls_filter>-rule     = mc_filter_rule_text_pattern.
    CLEAR <ls_filter>-tr_textfilter1.

    IF iv_textfilter1 CA '*+'. " Pattern
      ls_value1-sign   = 'I'.
      ls_value1-option = 'CP'.
      ls_value1-low    = iv_textfilter1.
    ELSE.
      ls_value1-sign   = 'I'.
      ls_value1-option = 'EQ'.
      ls_value1-low    = iv_textfilter1.
    ENDIF.
    APPEND ls_value1 TO <ls_filter>-tr_textfilter1.

  ENDMETHOD.