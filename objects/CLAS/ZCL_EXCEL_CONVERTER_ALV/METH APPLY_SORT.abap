  METHOD apply_sort.
    DATA: lt_otab TYPE abap_sortorder_tab,
          ls_otab TYPE abap_sortorder.

    FIELD-SYMBOLS: <fs_table> TYPE STANDARD TABLE,
                   <fs_sort>  TYPE lvc_s_sort.

    CREATE DATA eo_table LIKE it_table.
    ASSIGN eo_table->* TO <fs_table>.

    <fs_table> = it_table.

    SORT wt_sort BY spos.
    LOOP AT wt_sort ASSIGNING <fs_sort>.
      IF <fs_sort>-up = abap_true.
        ls_otab-name = <fs_sort>-fieldname.
        ls_otab-descending = abap_false.
*     ls_otab-astext     = abap_true. " not only text fields
        INSERT ls_otab INTO TABLE lt_otab.
      ENDIF.
      IF <fs_sort>-down = abap_true.
        ls_otab-name = <fs_sort>-fieldname.
        ls_otab-descending = abap_true.
*     ls_otab-astext     = abap_true. " not only text fields
        INSERT ls_otab INTO TABLE lt_otab.
      ENDIF.
    ENDLOOP.
    IF lt_otab IS NOT INITIAL.
      SORT <fs_table> BY (lt_otab).
    ENDIF.

  ENDMETHOD.