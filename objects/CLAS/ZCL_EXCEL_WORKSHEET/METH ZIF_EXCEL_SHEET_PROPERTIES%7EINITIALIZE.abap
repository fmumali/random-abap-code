  METHOD zif_excel_sheet_properties~initialize.

    zif_excel_sheet_properties~show_zeros   = zif_excel_sheet_properties=>c_showzero.
    zif_excel_sheet_properties~summarybelow = zif_excel_sheet_properties=>c_below_on.
    zif_excel_sheet_properties~summaryright = zif_excel_sheet_properties=>c_right_on.

* inizialize zoomscale values
    zif_excel_sheet_properties~zoomscale = 100.
    zif_excel_sheet_properties~zoomscale_normal = 100.
    zif_excel_sheet_properties~zoomscale_pagelayoutview = 100 .
    zif_excel_sheet_properties~zoomscale_sheetlayoutview = 100 .
  ENDMETHOD.                    "ZIF_EXCEL_SHEET_PROPERTIES~INITIALIZE