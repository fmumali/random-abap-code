  METHOD fill_row_outlines.

    TYPES: BEGIN OF lts_row_data,
             row           TYPE i,
             outline_level TYPE i,
           END OF lts_row_data,
           ltt_row_data TYPE SORTED TABLE OF lts_row_data WITH UNIQUE KEY row.

    DATA: lt_row_data             TYPE ltt_row_data,
          ls_row_data             LIKE LINE OF lt_row_data,
          lt_collapse_rows        TYPE HASHED TABLE OF i WITH UNIQUE KEY table_line,

          lv_collapsed            TYPE abap_bool,

          lv_outline_level        TYPE i,
          lv_next_consecutive_row TYPE i,
          lt_outline_rows         TYPE zcl_excel_worksheet=>mty_ts_outlines_row,
          ls_outline_row          LIKE LINE OF lt_outline_rows,
          lo_row                  TYPE REF TO zcl_excel_row,
          lo_row_iterator         TYPE REF TO zcl_excel_collection_iterator,
          lv_row_offset           TYPE i,
          lv_row_collapse_flag    TYPE i.


    FIELD-SYMBOLS: <ls_row_data>      LIKE LINE OF lt_row_data.

* First collect information about outlines ( outline leven and collapsed state )
    lo_row_iterator = io_worksheet->get_rows_iterator( ).
    WHILE lo_row_iterator->has_next( ) = abap_true.
      lo_row ?= lo_row_iterator->get_next( ).
      ls_row_data-row           = lo_row->get_row_index( ).
      ls_row_data-outline_level = lo_row->get_outline_level( ).
      IF ls_row_data-outline_level IS NOT INITIAL.
        INSERT ls_row_data INTO TABLE lt_row_data.
      ENDIF.

      lv_collapsed = lo_row->get_collapsed( ).
      IF lv_collapsed = abap_true.
        INSERT lo_row->get_row_index( ) INTO TABLE lt_collapse_rows.
      ENDIF.
    ENDWHILE.

* Now parse this information - we need consecutive rows - any gap will create a new outline
    DO 7 TIMES.  " max number of outlines allowed
      lv_outline_level = sy-index.
      CLEAR lv_next_consecutive_row.
      CLEAR ls_outline_row.
      LOOP AT lt_row_data ASSIGNING <ls_row_data> WHERE outline_level >= lv_outline_level.

        IF lv_next_consecutive_row    <> <ls_row_data>-row   " A gap --> close all open outlines
          AND lv_next_consecutive_row IS NOT INITIAL.        " First time in loop.
          INSERT ls_outline_row INTO TABLE lt_outline_rows.
          CLEAR: ls_outline_row.
        ENDIF.

        IF ls_outline_row-row_from IS INITIAL.
          ls_outline_row-row_from = <ls_row_data>-row.
        ENDIF.
        ls_outline_row-row_to = <ls_row_data>-row.

        lv_next_consecutive_row = <ls_row_data>-row + 1.

      ENDLOOP.
      IF ls_outline_row-row_from IS NOT INITIAL.
        INSERT ls_outline_row INTO TABLE lt_outline_rows.
      ENDIF.
    ENDDO.

* lt_outline_rows holds all outline information
* we now need to determine whether the outline is collapsed or not
    LOOP AT lt_outline_rows INTO ls_outline_row.

      IF io_worksheet->zif_excel_sheet_properties~summarybelow = zif_excel_sheet_properties=>c_below_off.
        lv_row_collapse_flag = ls_outline_row-row_from - 1.
      ELSE.
        lv_row_collapse_flag = ls_outline_row-row_to + 1.
      ENDIF.
      READ TABLE lt_collapse_rows TRANSPORTING NO FIELDS WITH TABLE KEY table_line = lv_row_collapse_flag.
      IF sy-subrc = 0.
        ls_outline_row-collapsed = abap_true.
      ENDIF.
      io_worksheet->set_row_outline( iv_row_from  = ls_outline_row-row_from
                                     iv_row_to    = ls_outline_row-row_to
                                     iv_collapsed = ls_outline_row-collapsed ).

    ENDLOOP.

* Finally purge outline information ( collapsed state, outline leve)  from row_dimensions, since we want to keep these in the outline-table
    lo_row_iterator = io_worksheet->get_rows_iterator( ).
    WHILE lo_row_iterator->has_next( ) = abap_true.
      lo_row ?= lo_row_iterator->get_next( ).

      lo_row->set_outline_level( 0 ).
      lo_row->set_collapsed( abap_false ).

    ENDWHILE.

  ENDMETHOD.