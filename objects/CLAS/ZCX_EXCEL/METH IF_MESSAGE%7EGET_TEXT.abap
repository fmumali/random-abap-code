  METHOD if_message~get_text.

    IF me->error IS NOT INITIAL.
*--------------------------------------------------------------------*
* If message was supplied explicitly use this
*--------------------------------------------------------------------*
      result = me->error .
    ELSEIF me->syst_at_raise IS NOT INITIAL.
*--------------------------------------------------------------------*
* If message was supplied by syst create messagetext now
*--------------------------------------------------------------------*
      MESSAGE ID syst_at_raise-msgid TYPE syst_at_raise-msgty NUMBER syst_at_raise-msgno
           WITH  syst_at_raise-msgv1 syst_at_raise-msgv2 syst_at_raise-msgv3 syst_at_raise-msgv4
           INTO  result.
    ELSE.
*--------------------------------------------------------------------*
* otherwise use standard method to derive text
*--------------------------------------------------------------------*
      CALL METHOD super->if_message~get_text
        RECEIVING
          result = result.
    ENDIF.
  ENDMETHOD.