  METHOD create_xl_sheet_column_formula.

    TYPES: ls_column_formula_used     TYPE mty_column_formula_used,
           lv_column_alpha            TYPE zexcel_cell_column_alpha,
           lv_top_cell_coords         TYPE zexcel_cell_coords,
           lv_bottom_cell_coords      TYPE zexcel_cell_coords,
           lv_cell_coords             TYPE zexcel_cell_coords,
           lv_ref_value               TYPE string,
           lv_test_shared             TYPE string,
           lv_si                      TYPE i,
           lv_1st_line_shared_formula TYPE abap_bool.
    DATA: lv_value                   TYPE string,
          ls_column_formula_used     TYPE mty_column_formula_used,
          lv_column_alpha            TYPE zexcel_cell_column_alpha,
          lv_top_cell_coords         TYPE zexcel_cell_coords,
          lv_bottom_cell_coords      TYPE zexcel_cell_coords,
          lv_cell_coords             TYPE zexcel_cell_coords,
          lv_ref_value               TYPE string,
          lv_1st_line_shared_formula TYPE abap_bool.
    FIELD-SYMBOLS: <ls_column_formula>      TYPE zcl_excel_worksheet=>mty_s_column_formula,
                   <ls_column_formula_used> TYPE mty_column_formula_used.


    READ TABLE it_column_formulas WITH TABLE KEY id = is_sheet_content-column_formula_id ASSIGNING <ls_column_formula>.
    ASSERT sy-subrc = 0.

    lv_value = <ls_column_formula>-formula.
    lv_1st_line_shared_formula = abap_false.
    eo_element = io_document->create_simple_element( name   = 'f'
                                                     parent = io_document ).
    READ TABLE ct_column_formulas_used WITH TABLE KEY id = is_sheet_content-column_formula_id ASSIGNING <ls_column_formula_used>.
    IF sy-subrc <> 0.
      CLEAR ls_column_formula_used.
      ls_column_formula_used-id = is_sheet_content-column_formula_id.
      IF is_formula_shareable( ip_formula = lv_value ) = abap_true.
        ls_column_formula_used-t = 'shared'.
        ls_column_formula_used-si = cv_si.
        CONDENSE ls_column_formula_used-si.
        cv_si = cv_si + 1.
        lv_1st_line_shared_formula = abap_true.
      ENDIF.
      INSERT ls_column_formula_used INTO TABLE ct_column_formulas_used ASSIGNING <ls_column_formula_used>.
    ENDIF.

    IF lv_1st_line_shared_formula = abap_true OR <ls_column_formula_used>-t <> 'shared'.
      lv_column_alpha = zcl_excel_common=>convert_column2alpha( ip_column = is_sheet_content-cell_column ).
      lv_top_cell_coords = |{ lv_column_alpha }{ <ls_column_formula>-table_top_left_row + 1 }|.
      lv_bottom_cell_coords = |{ lv_column_alpha }{ <ls_column_formula>-table_bottom_right_row + 1 }|.
      lv_cell_coords = |{ lv_column_alpha }{ is_sheet_content-cell_row }|.
      IF lv_top_cell_coords = lv_cell_coords.
        lv_ref_value = |{ lv_top_cell_coords }:{ lv_bottom_cell_coords }|.
      ELSE.
        lv_ref_value = |{ lv_cell_coords }:{ lv_bottom_cell_coords }|.
        lv_value = zcl_excel_common=>shift_formula(
            iv_reference_formula = lv_value
            iv_shift_cols        = 0
            iv_shift_rows        = is_sheet_content-cell_row - <ls_column_formula>-table_top_left_row - 1 ).
      ENDIF.
    ENDIF.

    IF <ls_column_formula_used>-t = 'shared'.
      eo_element->set_attribute( name  = 't'
                                 value = <ls_column_formula_used>-t ).
      eo_element->set_attribute( name  = 'si'
                                 value = <ls_column_formula_used>-si ).
      IF lv_1st_line_shared_formula = abap_true.
        eo_element->set_attribute( name  = 'ref'
                                   value = lv_ref_value ).
        eo_element->set_value( value = lv_value ).
      ENDIF.
    ELSE.
      eo_element->set_value( value = lv_value ).
    ENDIF.

  ENDMETHOD.