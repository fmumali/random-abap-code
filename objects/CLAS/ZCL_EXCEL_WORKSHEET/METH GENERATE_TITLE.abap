  METHOD generate_title.
    DATA: lo_worksheets_iterator TYPE REF TO zcl_excel_collection_iterator,
          lo_worksheet           TYPE REF TO zcl_excel_worksheet.

    DATA: t_titles    TYPE HASHED TABLE OF zexcel_sheet_title WITH UNIQUE KEY table_line,
          title       TYPE zexcel_sheet_title,
          sheetnumber TYPE i.

* Get list of currently used titles
    lo_worksheets_iterator = me->excel->get_worksheets_iterator( ).
    WHILE lo_worksheets_iterator->has_next( ) = abap_true.
      lo_worksheet ?= lo_worksheets_iterator->get_next( ).
      title = lo_worksheet->get_title( ).
      INSERT title INTO TABLE t_titles.
      ADD 1 TO sheetnumber.
    ENDWHILE.

* Now build sheetnumber.  Increase counter until we hit a number that is not used so far
    ADD 1 TO sheetnumber.  " Start counting with next number
    DO.
      title = sheetnumber.
      SHIFT title LEFT DELETING LEADING space.
      CONCATENATE 'Sheet'(001) title INTO ep_title.
      INSERT ep_title INTO TABLE t_titles.
      IF sy-subrc = 0.  " Title not used so far --> take it
        EXIT.
      ENDIF.

      ADD 1 TO sheetnumber.
    ENDDO.
  ENDMETHOD.                    "GENERATE_TITLE