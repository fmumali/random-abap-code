  METHOD get_shared_string_index.


    DATA ls_shared_string TYPE zexcel_s_shared_string.

    IF it_rtf IS INITIAL.
      READ TABLE shared_strings INTO ls_shared_string WITH TABLE KEY string_value = ip_cell_value.
      ep_index = ls_shared_string-string_no.
    ELSE.
      LOOP AT shared_strings INTO ls_shared_string WHERE string_value = ip_cell_value
                                                     AND rtf_tab = it_rtf.

        ep_index = ls_shared_string-string_no.
        EXIT.
      ENDLOOP.
    ENDIF.

  ENDMETHOD.