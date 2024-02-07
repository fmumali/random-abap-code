  PRIVATE SECTION.

    DATA worksheet TYPE REF TO zcl_excel_worksheet .
    DATA mt_filters TYPE tt_filters .

    METHODS validate_area
      RAISING
        zcx_excel .