  METHOD set_title.
*--------------------------------------------------------------------*
* ToDos:
*        2do §1  The current coding for replacing a named ranges name
*                after renaming a sheet should be checked if it is
*                really working if sheetname should be escaped
*--------------------------------------------------------------------*

*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,      (wip )              2012-12-08
*              - ...
* changes: aligning code
*          message made to support multilinguality
*--------------------------------------------------------------------*
* issue#243 - ' is not allowed as first character in sheet title
*              - Stefan Schmoecker,                          2012-12-02
* changes: added additional check for ' as first character
*--------------------------------------------------------------------*
    DATA: lo_worksheets_iterator TYPE REF TO zcl_excel_collection_iterator,
          lo_worksheet           TYPE REF TO zcl_excel_worksheet,
          errormessage           TYPE string,
          lv_rangesheetname_old  TYPE string,
          lv_rangesheetname_new  TYPE string,
          lo_ranges_iterator     TYPE REF TO zcl_excel_collection_iterator,
          lo_range               TYPE REF TO zcl_excel_range,
          lv_range_value         TYPE zexcel_range_value,
          lv_errormessage        TYPE string.                          " Can't pass '...'(abc) to exception-class


*--------------------------------------------------------------------*
* Check whether title consists only of allowed characters
* Illegal characters are: / \ [ ] * ? : --> http://msdn.microsoft.com/en-us/library/ff837411.aspx
* Illegal characters not in documentation:   ' as first character
*--------------------------------------------------------------------*
    IF ip_title CA '/\[]*?:'.
      lv_errormessage = 'Found illegal character in sheetname. List of forbidden characters: /\[]*?:'(402).
      zcx_excel=>raise_text( lv_errormessage ).
    ENDIF.

    IF ip_title IS NOT INITIAL AND ip_title(1) = `'`.
      lv_errormessage = 'Sheetname may not start with &'(403).   " & used instead of ' to allow fallbacklanguage
      REPLACE '&' IN lv_errormessage WITH `'`.
      zcx_excel=>raise_text( lv_errormessage ).
    ENDIF.


*--------------------------------------------------------------------*
* Check whether title is unique in workbook
*--------------------------------------------------------------------*
    lo_worksheets_iterator = me->excel->get_worksheets_iterator( ).
    WHILE lo_worksheets_iterator->has_next( ) = 'X'.

      lo_worksheet ?= lo_worksheets_iterator->get_next( ).
      CHECK me->guid <> lo_worksheet->get_guid( ).  " Don't check against itself
      IF ip_title = lo_worksheet->get_title( ).  " Not unique --> raise exception
        errormessage = 'Duplicate sheetname &'.
        REPLACE '&' IN errormessage WITH ip_title.
        zcx_excel=>raise_text( errormessage ).
      ENDIF.

    ENDWHILE.

*--------------------------------------------------------------------*
* Remember old sheetname and rename sheet to desired name
*--------------------------------------------------------------------*
    CONCATENATE me->title '!' INTO lv_rangesheetname_old.
    me->title = ip_title.

*--------------------------------------------------------------------*
* After changing this worksheet's title we have to adjust
* all ranges that are referring to this worksheet.
*--------------------------------------------------------------------*
* 2do §1  -  Check if the following quickfix is solid
*           I fear it isn't - but this implementation is better then
*           nothing at all since it handles a supposed majority of cases
*--------------------------------------------------------------------*
    CONCATENATE me->title '!' INTO lv_rangesheetname_new.

    lo_ranges_iterator = me->excel->get_ranges_iterator( ).
    WHILE lo_ranges_iterator->has_next( ) = 'X'.

      lo_range ?= lo_ranges_iterator->get_next( ).
      lv_range_value = lo_range->get_value( ).
      REPLACE ALL OCCURRENCES OF lv_rangesheetname_old IN lv_range_value WITH lv_rangesheetname_new.
      IF sy-subrc = 0.
        lo_range->set_range_value( lv_range_value ).
      ENDIF.

    ENDWHILE.


  ENDMETHOD.                    "SET_TITLE