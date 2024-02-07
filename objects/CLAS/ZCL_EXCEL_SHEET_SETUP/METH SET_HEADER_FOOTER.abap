  METHOD set_header_footer.

* Only Basic font/text formatting possible:
* Bold (yes / no), Font Type, Font Size
*
* usefull placeholders, which can be used in header/footer value strings
* '&P' - page number
* '&N' - total number of pages
* '&D' - Date
* '&T' - Time
* '&F' - File Name
* '&Z' - Path
* '&A' - Sheet name
* new line via class constant CL_ABAP_CHAR_UTILITIES=>newline
*
* Example Value String 'page &P of &N'
*
* DO NOT USE &L , &C or &R which automatically created as position markers

    me->odd_header = ip_odd_header.
    me->odd_footer = ip_odd_footer.
    me->even_header = ip_even_header.
    me->even_footer = ip_even_footer.

    IF me->even_header IS NOT INITIAL OR me->even_footer IS NOT INITIAL.
      me->diff_oddeven_headerfooter = abap_true.
    ENDIF.


  ENDMETHOD.