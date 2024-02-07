  METHOD find_var.

    DATA: lv_row            TYPE i,
          lv_column         TYPE i,
          lv_column_alpha   TYPE string,
          lv_value          TYPE string,
          ls_name_style     TYPE ts_name_style,
          lo_style          TYPE REF TO zcl_excel_style,
          lo_worksheet      TYPE REF TO zcl_excel_worksheet,
          ls_variable       TYPE ts_variable,
          lv_highest_column TYPE zexcel_cell_column,
          lv_highest_row    TYPE int4,
          lt_matches        TYPE match_result_tab,
          lv_search         TYPE string,
          lv_replace        TYPE string.

    FIELD-SYMBOLS:
      <ls_match>      TYPE match_result,
      <ls_range>      TYPE ts_range,
      <lv_sheet>      TYPE zexcel_sheet_title,
      <ls_name_style> TYPE ts_name_style.


    LOOP AT mt_sheet ASSIGNING <lv_sheet>.

      lo_worksheet ?= mo_excel->get_worksheet_by_name(  <lv_sheet> ).
      lv_row = 1.

      lv_highest_column = lo_worksheet->get_highest_column( ).
      lv_highest_row    = lo_worksheet->get_highest_row( ).

      WHILE lv_row <= lv_highest_row.
        lv_column = 1.
        WHILE lv_column <= lv_highest_column.
          lv_column_alpha = zcl_excel_common=>convert_column2alpha( lv_column ).
          CLEAR lo_style.
          lo_worksheet->get_cell(
            EXPORTING
              ip_column = lv_column_alpha
              ip_row    = lv_row
            IMPORTING
              ep_value = lv_value
              ep_style = lo_style ).

          FIND ALL OCCURRENCES OF REGEX '\[[^\]]*\]' IN lv_value RESULTS lt_matches.

          LOOP AT lt_matches ASSIGNING <ls_match>.
            lv_search = lv_value+<ls_match>-offset(<ls_match>-length).
            lv_replace = lv_search.

            TRANSLATE lv_replace TO UPPER CASE.

            CLEAR ls_variable.

            ls_variable-sheet = <lv_sheet>.
            ls_variable-name = lv_replace.
            TRANSLATE ls_variable-name USING '[ ] '.
            CONDENSE ls_variable-name .

            LOOP AT mt_range ASSIGNING <ls_range> WHERE sheet = <lv_sheet>
                                                    AND start <= lv_row
                                                    AND stop >= lv_row.
              ls_variable-parent = <ls_range>-id.
              EXIT.
            ENDLOOP.

            READ TABLE mt_var TRANSPORTING NO FIELDS WITH KEY sheet = ls_variable-sheet name = ls_variable-name parent = ls_variable-parent.
            IF sy-subrc NE 0.
              APPEND ls_variable TO mt_var.
            ENDIF.

            READ TABLE mt_name_styles WITH KEY sheet = ls_variable-sheet name = ls_variable-name parent = ls_variable-parent ASSIGNING <ls_name_style>.
            IF sy-subrc NE 0.
              CLEAR ls_name_style.
              ls_name_style-sheet = <lv_sheet>.
              ls_name_style-name = ls_variable-name.
              ls_name_style-parent = ls_variable-parent.
              APPEND ls_name_style TO mt_name_styles ASSIGNING <ls_name_style>.
            ENDIF.
            IF lo_style IS NOT BOUND.
              <ls_name_style>-text_counter = <ls_name_style>-text_counter + 1.
            ELSE.
              IF lo_style->number_format->format_code CA '0'
                   AND lo_style->number_format->format_code NS '0]'.
                <ls_name_style>-numeric_counter = <ls_name_style>-numeric_counter + 1.
              ELSEIF lo_style->number_format->format_code CA 'm'
                 AND lo_style->number_format->format_code CA 'd'
                 AND lo_style->number_format->format_code NA 'h'.
                <ls_name_style>-date_counter = <ls_name_style>-date_counter + 1.
              ELSEIF ( lo_style->number_format->format_code CA 'h' OR lo_style->number_format->format_code CA 's' )
                 AND lo_style->number_format->format_code NA 'd'.
                <ls_name_style>-time_counter = <ls_name_style>-time_counter + 1.
              ELSE.
                <ls_name_style>-text_counter = <ls_name_style>-text_counter + 1.
              ENDIF.
            ENDIF.

          ENDLOOP.
          lv_column = lv_column + 1.
        ENDWHILE.
        lv_row = lv_row + 1.
      ENDWHILE.
    ENDLOOP.

    SORT mt_range BY id .
  ENDMETHOD.