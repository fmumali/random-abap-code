  METHOD get_collapsed.

    DATA: lt_row_outlines  TYPE zcl_excel_worksheet=>mty_ts_outlines_row,
          lv_previous_row  TYPE i,
          lv_following_row TYPE i.

    r_collapsed = me->collapsed.

    CHECK r_collapsed = abap_false.  " Maybe new method for outlines is being used
    CHECK io_worksheet IS BOUND.

* If an outline is collapsed ( even inside an outer outline ) the line following the last line
* of the group gets the flag "collapsed"
    IF io_worksheet->zif_excel_sheet_properties~summarybelow = zif_excel_sheet_properties=>c_below_off.
      lv_following_row = me->row_index + 1.
      lt_row_outlines = io_worksheet->get_row_outlines( ).
      READ TABLE lt_row_outlines TRANSPORTING NO FIELDS WITH KEY row_from  = lv_following_row " first line of an outline
                                                                 collapsed = abap_true.       " that is collapsed
    ELSE.
      lv_previous_row = me->row_index - 1.
      lt_row_outlines = io_worksheet->get_row_outlines( ).
      READ TABLE lt_row_outlines TRANSPORTING NO FIELDS WITH KEY row_to    = lv_previous_row  " last line of an outline
                                                                 collapsed = abap_true.       " that is collapsed
    ENDIF.
    CHECK sy-subrc = 0.  " ok - we found it
    r_collapsed = abap_true.


  ENDMETHOD.