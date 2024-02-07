  METHOD create.

* .csv format with ; delimiter

* Start of insertion # issue 1134 - Dateretention of cellstyles(issue #139)
    me->excel->add_static_styles( ).
* End of insertion # issue 1134 - Dateretention of cellstyles(issue #139)

    ep_excel = me->create_csv( ).

  ENDMETHOD.