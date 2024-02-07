  METHOD zif_excel_sheet_protection~initialize.

    me->zif_excel_sheet_protection~protected = zif_excel_sheet_protection=>c_unprotected.
    CLEAR me->zif_excel_sheet_protection~password.
    me->zif_excel_sheet_protection~auto_filter            = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~delete_columns         = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~delete_rows            = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~format_cells           = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~format_columns         = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~format_rows            = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~insert_columns         = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~insert_hyperlinks      = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~insert_rows            = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~objects                = zif_excel_sheet_protection=>c_noactive.
*  me->zif_excel_sheet_protection~password               = zif_excel_sheet_protection=>c_noactive. "issue #68
    me->zif_excel_sheet_protection~pivot_tables           = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~protected              = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~scenarios              = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~select_locked_cells    = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~select_unlocked_cells  = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~sheet                  = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~sort                   = zif_excel_sheet_protection=>c_noactive.

  ENDMETHOD.                    "ZIF_EXCEL_SHEET_PROTECTION~INITIALIZE