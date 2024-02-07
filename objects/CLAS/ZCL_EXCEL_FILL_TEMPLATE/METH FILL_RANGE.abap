  METHOD fill_range.

    DATA: lt_tmp_cells_template        TYPE tt_cell_data_no_key,
          lt_cells_result              TYPE tt_cell_data_no_key,
          lt_tmp_cells                 TYPE tt_cell_data_no_key,
          ls_cell                      TYPE zexcel_s_cell_data,
          lt_tmp_merged_cells_template TYPE zcl_excel_worksheet=>mty_ts_merge,
          lt_merged_cells_result       TYPE zcl_excel_worksheet=>mty_ts_merge,
          lt_tmp_merged_cells          TYPE zcl_excel_worksheet=>mty_ts_merge,
          ls_merged_cell               LIKE LINE OF lt_tmp_merged_cells,
          lv_start_row                 TYPE i,
          lv_stop_row                  TYPE i,
          lv_cell_row                  TYPE i,
          lv_column_alpha              TYPE string,
          lt_matches                   TYPE match_result_tab,
          lv_search                    TYPE string,
          lv_var_name                  TYPE string,
          lv_cell_value                TYPE string.

    FIELD-SYMBOLS:
      <table>     TYPE ANY TABLE,
      <line>      TYPE any,
      <ls_range>  TYPE ts_range,
      <ls_cell>   TYPE zexcel_s_cell_data,
      <ls_match>  TYPE match_result,
      <var_value> TYPE any.


    cv_diff = cv_diff +  iv_range_length .

    lv_start_row = 1.


* recursive fill nested range

    LOOP AT mt_range ASSIGNING <ls_range> WHERE sheet = iv_sheet
                                            AND parent = iv_parent.


      lv_stop_row = <ls_range>-start - 1.

*      update cells before any range

      LOOP AT ct_cells INTO ls_cell  WHERE cell_row >= lv_start_row AND cell_row <= lv_stop_row .
        ls_cell-cell_row =  ls_cell-cell_row + cv_diff.
        lv_column_alpha = zcl_excel_common=>convert_column2alpha( ls_cell-cell_column ).

        ls_cell-cell_coords = ls_cell-cell_row.
        CONCATENATE lv_column_alpha ls_cell-cell_coords INTO ls_cell-cell_coords.
        CONDENSE ls_cell-cell_coords NO-GAPS.

        APPEND ls_cell TO lt_cells_result.
      ENDLOOP.



*      update merged cells before range

      LOOP AT ct_merged_cells INTO ls_merged_cell WHERE row_from >=  lv_start_row AND row_to <= lv_stop_row.
        ls_merged_cell-row_from = ls_merged_cell-row_from + cv_diff.
        ls_merged_cell-row_to = ls_merged_cell-row_to + cv_diff.

        APPEND ls_merged_cell TO lt_merged_cells_result.

      ENDLOOP.



      lv_start_row = <ls_range>-stop + 1.



      CLEAR:
       lt_tmp_cells_template,
       lt_tmp_merged_cells_template.


*copy cell template
      LOOP AT ct_cells INTO ls_cell WHERE cell_row >= <ls_range>-start AND cell_row <= <ls_range>-stop.
        APPEND ls_cell TO lt_tmp_cells_template.
      ENDLOOP.

      LOOP AT ct_merged_cells INTO ls_merged_cell WHERE row_from >= <ls_range>-start AND row_to <= <ls_range>-stop.
        APPEND ls_merged_cell TO lt_tmp_merged_cells_template.
      ENDLOOP.


      ASSIGN COMPONENT <ls_range>-name OF STRUCTURE iv_data TO <table>.
      CHECK sy-subrc = 0.

      cv_diff = cv_diff - <ls_range>-length.

*merge each line of data table with template
      LOOP AT <table> ASSIGNING <line>.
*        make local copy
        lt_tmp_cells = lt_tmp_cells_template.
        lt_tmp_merged_cells = lt_tmp_merged_cells_template.

*fill data

        fill_range(
          EXPORTING
            io_sheet        = io_sheet
            iv_sheet        = iv_sheet
            iv_parent       = <ls_range>-id
            iv_data         = <line>
            iv_range_length = <ls_range>-length
          CHANGING
            ct_cells        = lt_tmp_cells
            ct_merged_cells = lt_tmp_merged_cells
            cv_diff         = cv_diff ).

*collect data

        APPEND LINES OF lt_tmp_cells TO lt_cells_result.
        APPEND LINES OF lt_tmp_merged_cells TO lt_merged_cells_result.

      ENDLOOP.

    ENDLOOP.


    IF <ls_range> IS ASSIGNED.

      LOOP AT ct_cells INTO ls_cell WHERE cell_row > <ls_range>-stop .
        ls_cell-cell_row =  ls_cell-cell_row + cv_diff.
        lv_column_alpha = zcl_excel_common=>convert_column2alpha( ls_cell-cell_column ).

        ls_cell-cell_coords = ls_cell-cell_row.
        CONCATENATE lv_column_alpha ls_cell-cell_coords INTO ls_cell-cell_coords.
        CONDENSE ls_cell-cell_coords NO-GAPS.

        APPEND ls_cell TO lt_cells_result.
      ENDLOOP.

      ct_cells = lt_cells_result.

      LOOP AT ct_merged_cells INTO ls_merged_cell WHERE row_from > <ls_range>-stop.
        ls_merged_cell-row_from = ls_merged_cell-row_from + cv_diff.
        ls_merged_cell-row_to = ls_merged_cell-row_to + cv_diff.

        APPEND ls_merged_cell TO lt_merged_cells_result.
      ENDLOOP.

      ct_merged_cells = lt_merged_cells_result.

    ELSE.


      LOOP AT ct_cells ASSIGNING <ls_cell>.
        <ls_cell>-cell_row =  <ls_cell>-cell_row + cv_diff.
        lv_column_alpha = zcl_excel_common=>convert_column2alpha( <ls_cell>-cell_column ).

        <ls_cell>-cell_coords = <ls_cell>-cell_row.
        CONCATENATE lv_column_alpha <ls_cell>-cell_coords INTO <ls_cell>-cell_coords.
        CONDENSE <ls_cell>-cell_coords NO-GAPS.
      ENDLOOP.

      LOOP AT ct_merged_cells INTO ls_merged_cell .
        ls_merged_cell-row_from = ls_merged_cell-row_from + cv_diff.
        ls_merged_cell-row_to = ls_merged_cell-row_to + cv_diff.

        APPEND ls_merged_cell TO lt_merged_cells_result.
      ENDLOOP.

      ct_merged_cells = lt_merged_cells_result.

    ENDIF.


*check if variables in this range
    READ TABLE mt_var TRANSPORTING NO FIELDS WITH KEY sheet = iv_sheet parent = iv_parent.

    IF sy-subrc = 0.

*      replace variables of current range with data
      LOOP AT ct_cells ASSIGNING <ls_cell>.

        CLEAR lt_matches.

        lv_cell_value = <ls_cell>-cell_value.

        FIND ALL OCCURRENCES OF REGEX '\[[^\]]*\]' IN <ls_cell>-cell_value  RESULTS lt_matches.

        SORT lt_matches BY offset DESCENDING .

        LOOP AT lt_matches ASSIGNING <ls_match>.
          lv_search = <ls_cell>-cell_value+<ls_match>-offset(<ls_match>-length).
          lv_var_name = lv_search.

          TRANSLATE lv_var_name TO UPPER CASE.
          TRANSLATE lv_var_name USING '[ ] '.
          CONDENSE lv_var_name .

          ASSIGN COMPONENT lv_var_name OF STRUCTURE iv_data TO <var_value>.
          CHECK sy-subrc = 0.

          " Use SET_CELL to format correctly
          io_sheet->set_cell( ip_column = <ls_cell>-cell_column ip_row = <ls_cell>-cell_row - cv_diff ip_value = <var_value> ).
          lv_cell_row = <ls_cell>-cell_row - cv_diff.
          READ TABLE io_sheet->sheet_content INTO ls_cell
            WITH KEY cell_column = <ls_cell>-cell_column
                     cell_row    = lv_cell_row.
          REPLACE ALL OCCURRENCES OF lv_search IN <ls_cell>-cell_value WITH ls_cell-cell_value.
        ENDLOOP.

        IF lines( lt_matches ) = 1.
          lv_cell_row = <ls_cell>-cell_row - cv_diff.
          READ TABLE io_sheet->sheet_content INTO ls_cell
            WITH KEY cell_column = <ls_cell>-cell_column
                     cell_row    = lv_cell_row.
          <ls_cell>-data_type = ls_cell-data_type.
        ENDIF.

      ENDLOOP.
    ENDIF.

  ENDMETHOD.