  METHOD set_cell_formula.
    DATA:
      lv_column        TYPE zexcel_cell_column,
      lv_row           TYPE zexcel_cell_row,
      ls_sheet_content LIKE LINE OF me->sheet_content.

    FIELD-SYMBOLS:
                <sheet_content>                 LIKE LINE OF me->sheet_content.

*--------------------------------------------------------------------*
* Get cell to set formula into
*--------------------------------------------------------------------*
    normalize_columnrow_parameter( EXPORTING ip_columnrow = ip_columnrow
                                             ip_column    = ip_column
                                             ip_row       = ip_row
                                   IMPORTING ep_column    = lv_column
                                             ep_row       = lv_row ).

    READ TABLE me->sheet_content ASSIGNING <sheet_content> WITH TABLE KEY cell_row    = lv_row
                                                                          cell_column = lv_column.
    IF sy-subrc <> 0.                   " Create new entry in sheet_content if necessary
      CHECK ip_formula IS NOT INITIAL.  " only create new entry in sheet_content when a formula is passed
      ls_sheet_content-cell_row    = lv_row.
      ls_sheet_content-cell_column = lv_column.
      ls_sheet_content-cell_coords = zcl_excel_common=>convert_column_a_row2columnrow( i_column = lv_column i_row = lv_row ).
      INSERT ls_sheet_content INTO TABLE me->sheet_content ASSIGNING <sheet_content>.
    ENDIF.

*--------------------------------------------------------------------*
* Fieldsymbol now holds the relevant cell
*--------------------------------------------------------------------*
    <sheet_content>-cell_formula = ip_formula.


  ENDMETHOD.                    "SET_CELL_FORMULA