  PRIVATE SECTION.

    TYPES:
      BEGIN OF ts_autofilter,
        worksheet  TYPE REF TO zcl_excel_worksheet,
        autofilter TYPE REF TO zcl_excel_autofilter,
      END OF ts_autofilter .
    TYPES:
      tt_autofilters TYPE HASHED TABLE OF ts_autofilter WITH UNIQUE KEY worksheet .

    DATA mt_autofilters TYPE tt_autofilters .