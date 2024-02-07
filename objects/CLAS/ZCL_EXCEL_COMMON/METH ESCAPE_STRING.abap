  METHOD escape_string.
*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,      (done)              2012-12-08
*              - ...
* changes: aligning code
*          adding comments to explain what we are trying to achieve
*--------------------------------------------------------------------*
* issue#242 - Support escaping for white-spaces
*           - Escaping also necessary when ' encountered in input
*              - Stefan Schmoecker,                          2012-12-08
* changes: switched check if escaping is necessary to regular expression
*          and moved the "REPLACE"
*--------------------------------------------------------------------*
* issue#155 - lessening restrictions of input parameters
*              - Stefan Schmoecker,                          2012-12-08
* changes: ip_value changed to clike
*--------------------------------------------------------------------*
    DATA:       lv_value                        TYPE string.

*--------------------------------------------------------------------*
* There exist various situations when a space will be used to separate
* different parts of a string. When we have a string consisting spaces
* that will cause errors unless we "escape" the string by putting ' at
* the beginning and at the end of the string.
*--------------------------------------------------------------------*


*--------------------------------------------------------------------*
* When allowing clike-input parameters we might encounter trailing
* "real" blanks .  These are automatically eliminated when moving
* the input parameter to a string.
* Now any remaining spaces ( white-spaces or normal spaces ) should
* trigger the escaping as well as any '
*--------------------------------------------------------------------*
    lv_value = ip_value.


    FIND REGEX `\s|'|-` IN lv_value.  " \s finds regular and white spaces
    IF sy-subrc = 0.
      REPLACE ALL OCCURRENCES OF `'` IN lv_value WITH `''`.
      CONCATENATE `'` lv_value `'` INTO lv_value .
    ENDIF.

    ep_escaped_value = lv_value.

  ENDMETHOD.