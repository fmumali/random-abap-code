  METHOD get_table.
*--------------------------------------------------------------------*
* Comment D. Rauchenstein
* With this method, we get a fully functional Excel Upload, which solves
* a few issues of the other excel upload tools
* ZBCABA_ALSM_EXCEL_UPLOAD_EXT: Reads only up to 50 signs per Cell, Limit
* in row-Numbers. Other have Limitations of Lines, or you are not able
* to ignore filters or choosing the right tab.
*
* To get a fully functional XLSX Upload, you can use it e.g. with method
* CL_EXCEL_READER_2007->ZIF_EXCEL_READER~LOAD_FILE()
*--------------------------------------------------------------------*

    FIELD-SYMBOLS: <ls_line> TYPE data.
    FIELD-SYMBOLS: <lv_value> TYPE data.

    DATA lv_actual_row TYPE int4.
    DATA lv_actual_row_string TYPE string.
    DATA lv_actual_col TYPE int4.
    DATA lv_actual_col_string TYPE string.
    DATA lv_errormessage TYPE string.
    DATA lv_max_col TYPE zexcel_cell_column.
    DATA lv_max_row TYPE int4.
    DATA lv_delta_col TYPE int4.
    DATA lv_value  TYPE zexcel_cell_value.
    DATA lv_rc  TYPE sysubrc.
    DATA lx_conversion_error TYPE REF TO cx_sy_conversion_error.
    DATA lv_float TYPE f.
    DATA lv_type.
    DATA lv_tabix TYPE i.

    lv_max_col =  me->get_highest_column( ).
    IF iv_max_col IS SUPPLIED AND iv_max_col < lv_max_col.
      lv_max_col = iv_max_col.
    ENDIF.
    lv_max_row =  me->get_highest_row( ).
    IF iv_max_row IS SUPPLIED AND iv_max_row < lv_max_row.
      lv_max_row = iv_max_row.
    ENDIF.

*--------------------------------------------------------------------*
* The row counter begins with 1 and should be corrected with the skips
*--------------------------------------------------------------------*
    lv_actual_row =  iv_skipped_rows + 1.
    lv_actual_col =  iv_skipped_cols + 1.


    TRY.
*--------------------------------------------------------------------*
* Check if we the basic features are possible with given "any table"
*--------------------------------------------------------------------*
        APPEND INITIAL LINE TO et_table ASSIGNING <ls_line>.
        IF sy-subrc <> 0 OR <ls_line> IS NOT ASSIGNED.

          lv_errormessage = 'Error at inserting new Line to internal Table'(002).
          zcx_excel=>raise_text( lv_errormessage ).

        ELSE.
          lv_delta_col = lv_max_col - iv_skipped_cols.
          ASSIGN COMPONENT lv_delta_col OF STRUCTURE <ls_line> TO <lv_value>.
          IF sy-subrc <> 0 OR <lv_value> IS NOT ASSIGNED.
            lv_errormessage = 'Internal table has less columns than excel'(003).
            zcx_excel=>raise_text( lv_errormessage ).
          ELSE.
*--------------------------------------------------------------------*
*now we are ready for handle the table data
*--------------------------------------------------------------------*
            CLEAR et_table.
*--------------------------------------------------------------------*
* Handle each Row until end on right side
*--------------------------------------------------------------------*
            WHILE lv_actual_row <= lv_max_row .

*--------------------------------------------------------------------*
* Handle each Column until end on bottom
* First step is to step back on first column
*--------------------------------------------------------------------*
              lv_actual_col =  iv_skipped_cols + 1.

              UNASSIGN <ls_line>.
              APPEND INITIAL LINE TO et_table ASSIGNING <ls_line>.
              IF sy-subrc <> 0 OR <ls_line> IS NOT ASSIGNED.
                lv_errormessage = 'Error at inserting new Line to internal Table'(002).
                zcx_excel=>raise_text( lv_errormessage ).
              ENDIF.
              WHILE lv_actual_col <= lv_max_col.

                lv_delta_col = lv_actual_col - iv_skipped_cols.
                ASSIGN COMPONENT lv_delta_col OF STRUCTURE <ls_line> TO <lv_value>.
                IF sy-subrc <> 0.
                  lv_actual_col_string = lv_actual_col.
                  lv_actual_row_string = lv_actual_row.
                  CONCATENATE 'Error at assigning field (Col:'(004) lv_actual_col_string ' Row:'(005) lv_actual_row_string INTO lv_errormessage.
                  zcx_excel=>raise_text( lv_errormessage ).
                ENDIF.

                me->get_cell(
                  EXPORTING
                    ip_column  = lv_actual_col    " Cell Column
                    ip_row     = lv_actual_row    " Cell Row
                  IMPORTING
                    ep_value   = lv_value    " Cell Value
                    ep_rc      = lv_rc    " Return Value of ABAP Statements
                ).
                IF lv_rc <> 0
                  AND lv_rc <> 4                                                   "No found error means, zero/no value in cell
                  AND lv_rc <> 8. "rc is 8 when the last row contains cells with zero / no values
                  lv_actual_col_string = lv_actual_col.
                  lv_actual_row_string = lv_actual_row.
                  CONCATENATE 'Error at reading field value (Col:'(007) lv_actual_col_string ' Row:'(005) lv_actual_row_string INTO lv_errormessage.
                  zcx_excel=>raise_text( lv_errormessage ).
                ENDIF.

                TRY.
                    DESCRIBE FIELD <lv_value> TYPE lv_type.
                    IF lv_type = 'D'.
                      <lv_value> = zcl_excel_common=>excel_string_to_date( ip_value = lv_value ).
                    ELSE.
                      <lv_value> = lv_value. "Will raise exception if data type of <lv_value> is not float (or decfloat16/34) and excel delivers exponential number e.g. -2.9398924194538267E-2
                    ENDIF.
                  CATCH cx_sy_conversion_error INTO lx_conversion_error.
                    "Another try with conversion to float...
                    IF lv_type = 'P'.
                      <lv_value> = lv_float = lv_value.
                    ELSE.
                      RAISE EXCEPTION lx_conversion_error. "Pass on original exception
                    ENDIF.
                ENDTRY.

*  CATCH zcx_excel.    "
                ADD 1 TO lv_actual_col.
              ENDWHILE.
              ADD 1 TO lv_actual_row.
            ENDWHILE.

            IF iv_skip_bottom_empty_rows = abap_true.
              lv_tabix = lines( et_table ).
              WHILE lv_tabix >= 1.
                READ TABLE et_table INDEX lv_tabix ASSIGNING <ls_line>.
                ASSERT sy-subrc = 0.
                IF <ls_line> IS NOT INITIAL.
                  EXIT.
                ENDIF.
                DELETE et_table INDEX lv_tabix.
                lv_tabix = lv_tabix - 1.
              ENDWHILE.
            ENDIF.

          ENDIF.


        ENDIF.

      CATCH cx_sy_assign_cast_illegal_cast.
        lv_actual_col_string = lv_actual_col.
        lv_actual_row_string = lv_actual_row.
        CONCATENATE 'Error at assigning field (Col:'(004) lv_actual_col_string ' Row:'(005) lv_actual_row_string INTO lv_errormessage.
        zcx_excel=>raise_text( lv_errormessage ).
      CATCH cx_sy_assign_cast_unknown_type.
        lv_actual_col_string = lv_actual_col.
        lv_actual_row_string = lv_actual_row.
        CONCATENATE 'Error at assigning field (Col:'(004) lv_actual_col_string ' Row:'(005) lv_actual_row_string INTO lv_errormessage.
        zcx_excel=>raise_text( lv_errormessage ).
      CATCH cx_sy_assign_out_of_range.
        lv_errormessage = 'Internal table has less columns than excel'(003).
        zcx_excel=>raise_text( lv_errormessage ).
      CATCH cx_sy_conversion_error.
        lv_actual_col_string = lv_actual_col.
        lv_actual_row_string = lv_actual_row.
        CONCATENATE 'Error at converting field value (Col:'(006) lv_actual_col_string ' Row:'(005) lv_actual_row_string INTO lv_errormessage.
        zcx_excel=>raise_text( lv_errormessage ).

    ENDTRY.
  ENDMETHOD.                    "get_table