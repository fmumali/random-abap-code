  METHOD flag2bool.


    IF ip_flag EQ abap_true.
      ep_boolean = 'true'.
    ELSE.
      ep_boolean = 'false'.
    ENDIF.
  ENDMETHOD.