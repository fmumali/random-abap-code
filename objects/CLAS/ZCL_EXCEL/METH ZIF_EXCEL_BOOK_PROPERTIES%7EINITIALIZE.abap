  METHOD zif_excel_book_properties~initialize.
    DATA: lv_timestamp TYPE timestampl.

    me->zif_excel_book_properties~application     = 'Microsoft Excel'.
    me->zif_excel_book_properties~appversion      = '12.0000'.

    GET TIME STAMP FIELD lv_timestamp.
    me->zif_excel_book_properties~created         = lv_timestamp.
    me->zif_excel_book_properties~creator         = sy-uname.
    me->zif_excel_book_properties~description     = zcl_excel=>version.
    me->zif_excel_book_properties~modified        = lv_timestamp.
    me->zif_excel_book_properties~lastmodifiedby  = sy-uname.
  ENDMETHOD.