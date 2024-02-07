  METHOD get_range.

    DATA:
      lo_worksheets_iterator TYPE REF TO zcl_excel_collection_iterator,
      lo_worksheet           TYPE REF TO zcl_excel_worksheet,
      lo_range_iterator      TYPE REF TO zcl_excel_collection_iterator,
      lo_range               TYPE REF TO zcl_excel_range.


    lo_worksheets_iterator = mo_excel->get_worksheets_iterator( ).


    WHILE lo_worksheets_iterator->has_next( ) = abap_true.
      lo_worksheet ?= lo_worksheets_iterator->get_next( ).
      APPEND lo_worksheet->get_title( ) TO mt_sheet.
    ENDWHILE.

    lo_range_iterator = mo_excel->get_ranges_iterator( ).

    WHILE lo_range_iterator->has_next(  ) = abap_true.
      lo_range ?= lo_range_iterator->get_next( ).
      validate_range( lo_range ).
    ENDWHILE.

  ENDMETHOD.