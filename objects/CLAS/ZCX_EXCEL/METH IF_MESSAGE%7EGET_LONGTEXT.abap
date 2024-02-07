  METHOD if_message~get_longtext.

    IF   me->error         IS NOT INITIAL
      OR me->syst_at_raise IS NOT INITIAL.
*--------------------------------------------------------------------*
* If message was supplied explicitly use this as longtext as well
*--------------------------------------------------------------------*
      result = me->get_text( ).
    ELSE.
*--------------------------------------------------------------------*
* otherwise use standard method to derive text
*--------------------------------------------------------------------*
      result = super->if_message~get_longtext( preserve_newlines = preserve_newlines ).
    ENDIF.
  ENDMETHOD.