  METHOD put_cell_to_worksheet.
    CHECK is_cell-value IS NOT INITIAL
       OR is_cell-formula IS NOT INITIAL
       OR is_cell-style IS NOT INITIAL.
    CALL METHOD io_worksheet->set_cell
      EXPORTING
        ip_column    = is_cell-column
        ip_row       = is_cell-row
        ip_value     = is_cell-value
        ip_formula   = is_cell-formula
        ip_data_type = is_cell-datatype
        ip_style     = is_cell-style.
  ENDMETHOD.