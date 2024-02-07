  METHOD unescape_string.

    CONSTANTS   lcv_regex                       TYPE string VALUE `^'[^']`    & `|` &  " Beginning single ' OR
                                                                  `[^']'$`    & `|` &  " Trailing single '  OR
                                                                  `[^']'[^']`.         " Single ' somewhere in between


    DATA:       lv_errormessage                 TYPE string.                          " Can't pass '...'(abc) to exception-class

*--------------------------------------------------------------------*
* This method is used to extract the "real" string from an escaped string.
* An escaped string can be identified by a beginning ' which must be
* accompanied by a trailing '
* All '' in between beginning and trailing ' are treated as single '
*--------------------------------------------------------------------*

*--------------------------------------------------------------------*
* When allowing clike-input parameters we might encounter trailing
* "real" blanks .  These are automatically eliminated when moving
* the input parameter to a string.
*--------------------------------------------------------------------*
    ev_unescaped_string = iv_escaped.           " Pass through if not escaped

    CHECK ev_unescaped_string IS NOT INITIAL.   " Nothing to do if empty
    CHECK ev_unescaped_string(1) = `'`.         " Nothing to do if not escaped

*--------------------------------------------------------------------*
* Remove leading and trailing '
*--------------------------------------------------------------------*
    REPLACE REGEX `^'(.*)'$` IN ev_unescaped_string WITH '$1'.
    IF sy-subrc <> 0.
      lv_errormessage = 'Input not properly escaped - &'(002).
      zcx_excel=>raise_text( lv_errormessage ).
    ENDIF.

*--------------------------------------------------------------------*
* Any remaining single ' should not be here
*--------------------------------------------------------------------*
    FIND REGEX lcv_regex IN ev_unescaped_string.
    IF sy-subrc = 0.
      lv_errormessage = 'Input not properly escaped - &'(002).
      zcx_excel=>raise_text( lv_errormessage ).
    ENDIF.

*--------------------------------------------------------------------*
* Replace '' with '
*--------------------------------------------------------------------*
    REPLACE ALL OCCURRENCES OF `''` IN ev_unescaped_string WITH `'`.


  ENDMETHOD.