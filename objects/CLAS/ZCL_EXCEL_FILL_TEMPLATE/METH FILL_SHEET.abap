  METHOD fill_sheet.

    DATA: lo_worksheet      TYPE REF TO zcl_excel_worksheet,
          lt_sheet_cells    TYPE tt_cell_data_no_key,
          lt_merged_cells   TYPE zcl_excel_worksheet=>mty_ts_merge,
          lt_merged_cells_2 TYPE zcl_excel_worksheet=>mty_ts_merge,
          lv_initial_diff   TYPE i.
    FIELD-SYMBOLS:
      <any_data>       TYPE any,
      <ls_sheet_cell>  TYPE zexcel_s_cell_data,
      <ls_merged_cell> TYPE zcl_excel_worksheet=>mty_merge.


    lo_worksheet = mo_excel->get_worksheet_by_name( iv_data-sheet ).

    lt_sheet_cells = lo_worksheet->sheet_content.
    lt_merged_cells = lo_worksheet->mt_merged_cells.

    ASSIGN iv_data-data->* TO <any_data>.

    fill_range(
      EXPORTING
        io_sheet        = lo_worksheet
        iv_range_length = 0
        iv_sheet        = iv_data-sheet
        iv_parent       = 0
        iv_data         = <any_data>
      CHANGING
        ct_cells        = lt_sheet_cells
        ct_merged_cells = lt_merged_cells
        cv_diff         = lv_initial_diff  ).


    CLEAR lo_worksheet->sheet_content.

    LOOP AT lt_sheet_cells ASSIGNING <ls_sheet_cell>.
      INSERT <ls_sheet_cell> INTO TABLE lo_worksheet->sheet_content.
    ENDLOOP.

    lt_merged_cells_2 = lo_worksheet->mt_merged_cells.
    LOOP AT lt_merged_cells_2 ASSIGNING <ls_merged_cell>.
      lo_worksheet->delete_merge( ip_cell_column = <ls_merged_cell>-col_from ip_cell_row = <ls_merged_cell>-row_from ).
    ENDLOOP.

    LOOP AT lt_merged_cells ASSIGNING <ls_merged_cell>.
      lo_worksheet->set_merge(
          ip_column_start = <ls_merged_cell>-col_from
          ip_column_end   = <ls_merged_cell>-col_to
          ip_row          = <ls_merged_cell>-row_from
          ip_row_to       = <ls_merged_cell>-row_to ).
    ENDLOOP.

  ENDMETHOD.