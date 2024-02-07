  METHOD zif_excel_book_protection~initialize.
    me->zif_excel_book_protection~protected      = zif_excel_book_protection=>c_unprotected.
    me->zif_excel_book_protection~lockrevision   = zif_excel_book_protection=>c_unlocked.
    me->zif_excel_book_protection~lockstructure  = zif_excel_book_protection=>c_unlocked.
    me->zif_excel_book_protection~lockwindows    = zif_excel_book_protection=>c_unlocked.
    CLEAR me->zif_excel_book_protection~workbookpassword.
    CLEAR me->zif_excel_book_protection~revisionspassword.
  ENDMETHOD.