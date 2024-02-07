  METHOD add_pagebreak.
    DATA: ls_pagebreak      LIKE LINE OF me->mt_pagebreaks.

    ls_pagebreak-cell_row    = ip_row.
    ls_pagebreak-cell_column = zcl_excel_common=>convert_column2int( ip_column ).

    INSERT ls_pagebreak INTO TABLE me->mt_pagebreaks.


  ENDMETHOD.