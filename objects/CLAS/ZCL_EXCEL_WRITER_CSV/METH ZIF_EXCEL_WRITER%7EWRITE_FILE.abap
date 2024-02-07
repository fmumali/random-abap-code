  METHOD zif_excel_writer~write_file.
    me->excel = io_excel.
    ep_file = me->create( ).
  ENDMETHOD.