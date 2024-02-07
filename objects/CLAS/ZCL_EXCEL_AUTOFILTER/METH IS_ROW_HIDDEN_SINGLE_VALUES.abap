  METHOD is_row_hidden_single_values.


    DATA: lv_value TYPE string.

    FIELD-SYMBOLS: <ls_sheet_content> LIKE LINE OF me->worksheet->sheet_content.

    rv_is_hidden = abap_false.   " Default setting is NOT HIDDEN = is in filter range

*--------------------------------------------------------------------*
* No filter values --> only symbol should be shown but nothing is being hidden
*--------------------------------------------------------------------*
    IF is_filter-t_values IS INITIAL.
      RETURN.
    ENDIF.

*--------------------------------------------------------------------*
* Get value of cell
*--------------------------------------------------------------------*
    READ TABLE me->worksheet->sheet_content ASSIGNING <ls_sheet_content> WITH TABLE KEY cell_row    = iv_row
                                                                                        cell_column = iv_col.
    IF sy-subrc = 0.
      lv_value = <ls_sheet_content>-cell_value.
    ELSE.
      CLEAR lv_value.
    ENDIF.

*--------------------------------------------------------------------*
* Check whether it is affected by filter
* this needs to be extended if we support other filtertypes
* other than single values
*--------------------------------------------------------------------*
    READ TABLE is_filter-t_values TRANSPORTING NO FIELDS WITH TABLE KEY table_line =  lv_value.
    IF sy-subrc <> 0.
      rv_is_hidden = abap_true.
    ENDIF.

  ENDMETHOD.