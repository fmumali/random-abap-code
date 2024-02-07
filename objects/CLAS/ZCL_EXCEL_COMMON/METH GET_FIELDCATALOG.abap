  METHOD get_fieldcatalog.
    DATA: lr_dref_tab           TYPE REF TO data,
          lo_salv_table         TYPE REF TO cl_salv_table,
          lo_salv_columns_table TYPE REF TO cl_salv_columns_table,
          lt_salv_t_column_ref  TYPE salv_t_column_ref,
          ls_salv_t_column_ref  LIKE LINE OF lt_salv_t_column_ref,
          lo_salv_column_table  TYPE REF TO cl_salv_column_table.

    FIELD-SYMBOLS: <tab>          TYPE STANDARD TABLE.
    FIELD-SYMBOLS: <fcat>         LIKE LINE OF ep_fieldcatalog.

* Get copy of IP_TABLE-structure <-- must be changeable to create salv
    CREATE DATA lr_dref_tab LIKE ip_table.
    ASSIGN lr_dref_tab->* TO <tab>.
* Create salv --> implicitly create fieldcat
    TRY.
        cl_salv_table=>factory( IMPORTING
                                  r_salv_table   = lo_salv_table
                                CHANGING
                                  t_table        = <tab>  ).
        lo_salv_columns_table = lo_salv_table->get_columns( ).
        lt_salv_t_column_ref  = lo_salv_columns_table->get( ).
      CATCH cx_root.
* maybe some errorhandling here - just haven't made up my mind yet
    ENDTRY.

* Loop through columns and set relevant fields ( fieldname, texts )
    LOOP AT lt_salv_t_column_ref INTO ls_salv_t_column_ref.

      lo_salv_column_table ?= ls_salv_t_column_ref-r_column.
      APPEND INITIAL LINE TO ep_fieldcatalog ASSIGNING <fcat>.
      <fcat>-position  = sy-tabix.
      <fcat>-fieldname = ls_salv_t_column_ref-columnname.
      <fcat>-scrtext_s = ls_salv_t_column_ref-r_column->get_short_text( ).
      <fcat>-scrtext_m = ls_salv_t_column_ref-r_column->get_medium_text( ).
      <fcat>-scrtext_l = ls_salv_t_column_ref-r_column->get_long_text( ).
      <fcat>-currency_column = ls_salv_t_column_ref-r_column->get_currency_column( ).
      " If currency column not in structure then clear the field again
      IF <fcat>-currency_column IS NOT INITIAL.
        READ TABLE lt_salv_t_column_ref WITH KEY columnname = <fcat>-currency_column TRANSPORTING NO FIELDS.
        IF sy-subrc <> 0.
          CLEAR <fcat>-currency_column.
        ENDIF.
      ENDIF.

      IF ip_conv_exit_length = abap_false.
        <fcat>-abap_type = lo_salv_column_table->get_ddic_inttype( ).
      ENDIF.

      <fcat>-dynpfld   = 'X'.  " What in the world would we exclude here?
      " except for the MANDT-field of most tables ( 1st column that is )
      IF <fcat>-position = 1 AND lo_salv_column_table->get_ddic_datatype( ) = 'CLNT' AND iv_hide_mandt = abap_true.
        CLEAR <fcat>-dynpfld.
      ENDIF.

* For fields that don't a description (  i.e. defined by  "field type i," )
* just use the fieldname as description - that is better than nothing
      IF    <fcat>-scrtext_s IS INITIAL
        AND <fcat>-scrtext_m IS INITIAL
        AND <fcat>-scrtext_l IS INITIAL.
        CONCATENATE 'Col:' <fcat>-fieldname INTO <fcat>-scrtext_l  SEPARATED BY space.
        <fcat>-scrtext_m = <fcat>-scrtext_l.
        <fcat>-scrtext_s = <fcat>-scrtext_l.
      ENDIF.

    ENDLOOP.

  ENDMETHOD.