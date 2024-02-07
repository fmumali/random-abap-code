  METHOD freeze_panes.

    IF ip_num_columns IS NOT SUPPLIED AND ip_num_rows IS NOT SUPPLIED.
      zcx_excel=>raise_text( 'Pleas provide number of rows and/or columns to freeze' ).
    ENDIF.

    IF ip_num_columns IS SUPPLIED AND ip_num_columns <= 0.
      zcx_excel=>raise_text( 'Number of columns to freeze should be positive' ).
    ENDIF.

    IF ip_num_rows IS SUPPLIED AND ip_num_rows <= 0.
      zcx_excel=>raise_text( 'Number of rows to freeze should be positive' ).
    ENDIF.

    freeze_pane_cell_column = ip_num_columns + 1.
    freeze_pane_cell_row = ip_num_rows + 1.
  ENDMETHOD.                    "FREEZE_PANES